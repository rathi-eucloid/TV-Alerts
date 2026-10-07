#!/usr/bin/env python3
"""
script for scraping product pages from Samsung and Walmart using
Playwright; Amazon, BestBuy and Lowes prices come from Apify actors (tokens in the
APIFY_API_TOKENS env var, see the APIFY section).
- Saves each page's HTML to a unique file in outputs/ directory.
- Parses each HTML to extract product price and model number (or SKU for Samsung).

This version additionally saves the run's result as a single row in an Excel file:
- Columns = keys (flattened per-site/per-index keys like "amazon_1_url")
- Row = values for this run
- On next run the script appends a new row (does not overwrite previous data).

"""
from zoneinfo import ZoneInfo
import datetime
import asyncio
import json
import os
import random
import re
import socket
import shutil
import tempfile
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request

from urllib.parse import quote_plus
from playwright.async_api import async_playwright, TimeoutError

from openpyxl.utils import column_index_from_string, get_column_letter
# new imports for Excel writing
from openpyxl import Workbook, load_workbook

# =========================================================================
# script_2.py fixes (vs script.py)
# -------------------------------------------------------------------------
# 1. Redirect detection: many product pages now silently redirect to a
#    *different* product (e.g. S26 Ultra -> S25 FE, Z Fold 7 -> Z Flip 7) when
#    the item is out of stock/unavailable. We compare the product identifier we
#    REQUESTED (ASIN / BestBuy code / Samsung SKU) against the identifier in the
#    page's <link rel="canonical">. On mismatch we record "not available".
# 2. Duplicate-element fix: the old hardcoded selectors matched many elements on
#    the page (related items, sponsored, carousels), causing wrong reads. We now
#    scope price extraction to the MAIN price container only.
# 3. Updated selectors for current UI: Samsung price -> JSON-LD offer keyed by SKU.
# 4. Amazon, BestBuy and Lowes are no longer scraped with a browser; they come
#    from Apify actors (see the APIFY section). Redirect detection for Amazon /
#    BestBuy compares the requested ASIN / product code with what the actor
#    returned.
# 5. results.xlsx now uses the same layout as "Price Comparisons_v3_WIP":
#    51 product groups x 9 columns starting at column C, timestamp in column B.
# Everything else (URLs, delays, user agents, cookies logic) is unchanged.
# =========================================================================
NOT_AVAILABLE = "not available"

# Retry policy for transient failures (network errors, timeouts, or a page that
# navigated OK but didn't render its price in time). 1 initial try + 2 retries.
# Genuine outcomes (a redirect to another product, or a page that explicitly says
# it's unavailable) are treated as FINAL and are NOT retried.
MAX_ATTEMPTS = 3
RETRY_BACKOFF_SEC = 5

# ---- Walmart pacing ----
# Wait before re-attempting the SAME url / after finishing one url, before
# the next. Kept separate from RETRY_BACKOFF_SEC (Samsung) so tuning one
# site never changes another.
WALMART_RETRY_BACKOFF_SEC = 20
WALMART_BETWEEN_URL_GAP_SEC = 20

# Excel layout (per product group of 15 columns):
#   +0 Amazon price   +1 Samsung price   +2 BestBuy price   +3 Lowes price   +4 Walmart price
#   +5 SKU_Amazon     +6 SKU_Samsung     +7 SKU_BestBuy     +8 SKU_Lowes     +9 SKU_Walmart
#   +10 vs Amazon  +11 vs Bestbuy  +12 vs Lowes  +13 vs Walmart  (formulas)   +14 blank
# Price columns are written in SITES order; each SKU column sits SKU_OFFSET
# columns to the right of its price column.
SITES = ["amazon", "samsung", "bestbuy", "lowes", "walmart"]
SKU_OFFSET      = len(SITES)          # 5
VS_OFFSET       = 2 * len(SITES)      # 10
FIRST_GROUP_COL = 3        # column C
GROUP_STRIDE    = 15
TIMESTAMP_COL   = 2        # column B

# Product group headers for row 1 (one per slot, same order as the URL lists)
PRODUCT_LABELS = [
    "Samsung Smart LED Indoor Flat screen HDTV (UN43M70HAFXZA)",
    "Samsung Smart OLED Indoor Flat screen HDTV (QN65S90FAFXZA)",
    "Samsung Smart OLED Indoor Flat screen HDTV (QN77S90FAFXZA)",
    "Samsung Smart QLED Indoor Flat screen HDTV (QN43Q7FAAFXZA)",
    "Samsung Smart QLED Indoor Flat screen HDTV (QN65LS03FAFXZA)",
]

SUBHEADERS = ["Amazon price", "Samsung price", "BestBuy.com price",
              "Lowes.com price", "Walmart.com price",
              "SKU_ID_Amazon ", "SKU_ID_Samsung ", "SKU_ID_BestBuy.com",
              "SKU_ID_Lowes.com", "SKU_ID_Walmart.com",
              "vs Amazon", "vs Bestbuy", "vs Lowes", "vs Walmart"]


# ---- product identifiers (used for redirect detection & slot SKUs) ----
def amazon_id_from_url(url):
    m = re.search(r"/dp/([A-Z0-9]{10})", url or "", re.I)
    return m.group(1).upper() if m else None

def bestbuy_id_from_url(url):
    m = re.search(r"/product/[^/]+/([A-Z0-9]+)", url or "", re.I)
    return m.group(1).upper() if m else None

def lowes_id_from_url(url):
    # https://www.lowes.com/pd/<slug>/5013373117  -> "5013373117"
    m = re.search(r"/pd/(?:[^/?#]+/)?(\d+)", url or "")
    return m.group(1) if m else None

def walmart_id_from_url(url):
    # https://www.walmart.com/ip/<slug>/1527054368 -> "1527054368"
    m = re.search(r"/ip/(?:[^/?#]+/)?(\d+)", url or "")
    return m.group(1) if m else None

def get_canonical_href(html):
    """Pull <link rel=canonical href=...> without a full DOM parse."""
    m = re.search(r'<link\b[^>]*\brel=["\']canonical["\'][^>]*>', html, re.I)
    if not m:
        return None
    h = re.search(r'href=["\']([^"\']+)["\']', m.group(0), re.I)
    return h.group(1) if h else None

def _looks_unavailable(html, site):
    """True when the page EXPLICITLY signals the product is unavailable/sold out.

    Used by the retry loop to decide whether a "no price" outcome is FINAL (the
    seller genuinely isn't selling it -> don't waste retries) vs TRANSIENT (the
    price element simply didn't render this time -> retry). Verified against the
    saved pages: Amazon's own out-of-stock listings render the exact phrase
    "Currently unavailable"; only third-party offers carry a price, which we
    deliberately don't scrape.
    """
    if not html:
        return False
    low = html.lower()
    if site == "amazon":
        return "currently unavailable" in low
    if site == "bestbuy":
        return "sold out" in low or "no longer available" in low
    if site == "samsung":
        return ("sold out" in low or "coming soon" in low
                or "out of stock" in low or "notify me" in low)
    return False


def iter_ldjson(html):
    """Yield parsed JSON-LD objects from the HTML (regex-sliced, fast)."""
    for m in re.finditer(
            r'<script\b[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.I | re.S):
        try:
            data = json.loads(m.group(1).strip())
        except Exception:
            continue
        for it in (data if isinstance(data, list) else [data]):
            yield it


# ---- price cleaning (ported from ConvertDirtyTextPriceToNumbers/main2) ----
def clean_price_value(raw):
    """Return a float rounded to 2dp, or None. Mirrors main2_decimalPlaceTill2."""
    if raw is None:
        return None
    s = str(raw).strip()
    if s == "" or s == NOT_AVAILABLE:
        return None
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()
    s = re.sub(r"[^\d\.,\-]", "", s)
    if s == "" or re.fullmatch(r"[-\.,]*", s):
        return None
    s = s.replace("−", "-")
    if "-" in s:
        if s.count("-") > 1:
            s = s.replace("-", "")
        if s.startswith("-"):
            negative = not negative
            s = s.lstrip("-")
    has_dot, has_comma = "." in s, "," in s
    if has_dot and has_comma:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
            if s.count(".") > 1:
                left, right = s.rsplit(".", 1)
                s = left.replace(".", "") + "." + right
    elif has_comma and not has_dot:
        parts = s.split(",")
        if len(parts) >= 2 and len(parts[-1]) == 2:
            s = ",".join(parts[:-1]).replace(",", "") + "." + parts[-1]
        else:
            s = s.replace(",", "")
    elif has_dot and not has_comma:
        if s.count(".") > 1:
            left, right = s.rsplit(".", 1)
            s = left.replace(".", "") + "." + right
    s = re.sub(r"[^\d.]", "", s)
    if s.count(".") > 1:
        left, right = s.rsplit(".", 1)
        s = left.replace(".", "") + "." + right
    if s in ("", "."):
        return None
    try:
        value = round(float(s), 2)
    except Exception:
        return None
    return -value if negative else value


# -----------------------
# Shared helpers
# -----------------------
async def human_delay(min_sec=0.5, max_sec=2.5):
    """Wait for a random time between min_sec and max_sec seconds."""
    delay = random.uniform(min_sec, max_sec)
    await asyncio.sleep(delay)

async def human_delay_short():
    """Small helper to yield control briefly (kept minimal to respect original logic)."""
    await asyncio.sleep(0.1)

async def get_page_content_safe(page, retries=4):
    """Return page HTML, tolerating in-flight client-side navigations.

    BestBuy fires a delayed client-side navigation/reload ~20s after load, which
    made a bare `page.content()` throw:
      "Unable to retrieve content because the page is navigating and changing".
    We wait for the page to settle and retry; as a last resort we read
    document.documentElement.outerHTML via JS (works mid-navigation).
    """
    last_err = None
    for attempt in range(retries):
        try:
            # let any in-flight navigation finish before grabbing content
            try:
                await page.wait_for_load_state("domcontentloaded", timeout=8000)
            except Exception:
                pass
            return await page.content()
        except Exception as e:
            last_err = e
            # brief settle, then retry
            await asyncio.sleep(2.5)
    # final fallback: pull the DOM directly (succeeds even while navigating)
    try:
        return await page.evaluate("() => document.documentElement.outerHTML")
    except Exception:
        raise last_err

def sanitize_filename(s: str, maxlen: int = 200) -> str:
    """Create a filesystem-safe short filename from a string (URL)."""
    if not s:
        return "file"
    s_enc = quote_plus(s, safe="")
    s_clean = re.sub(r'[^A-Za-z0-9._-]', '_', s_enc)
    return s_clean[:maxlen]

def _to_jsonable(v):
    """Convert complex types to JSON strings for Excel storage; leave primitives as-is."""
    if isinstance(v, (dict, list, tuple)):
        return json.dumps(v, ensure_ascii=False)
    return v

def _render_cells(result, slot_sku, site):
    """Given a per-URL result dict, return (price_cell, sku_cell) for the sheet.

    - price -> float when parseable, "not available" for redirect/no-price, else None.
    - sku   -> canonical per-slot SM code (Amazon/BestBuy upper, Samsung lower);
               "not available" mirrors the price when the product wasn't found.
    """
    raw = result.get("price") if result else None
    if raw == NOT_AVAILABLE:
        return NOT_AVAILABLE, NOT_AVAILABLE
    num = clean_price_value(raw)
    if num is None:
        return None, None            # genuine gap (fetch error / empty URL) -> blank
    sku = None
    if slot_sku:
        sku = slot_sku if site == "samsung" else slot_sku.upper()
    return num, sku


def save_results_wip_format(am_res, bb_res, sam_res, lo_res, wm_res,
                            samsung_urls, ts_str,
                            excel_path="outputs/results.xlsx"):
    """Append one row per run to the results workbook:
      - row 1 = product group headers, row 2 = sub-headers, data from row 3
      - len(PRODUCT_LABELS) groups x GROUP_STRIDE (15) columns starting at
        column C; timestamp in column B
      - each group: Amazon/Samsung/BestBuy/Lowes/Walmart price, 5 SKU columns,
        4 'vs' formula columns (each site's price / Samsung price - 1), 1 blank
    Prices are written as numbers; SKU columns filled; 'vs' formulas added per
    row. Existing rows are kept.
    """
    os.makedirs(os.path.dirname(excel_path) or ".", exist_ok=True)

    if os.path.exists(excel_path):
        wb = load_workbook(excel_path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.cell(row=2, column=TIMESTAMP_COL, value="Timestamp EST")
        for s in range(len(PRODUCT_LABELS)):
            gc = FIRST_GROUP_COL + GROUP_STRIDE * s
            ws.cell(row=1, column=gc, value=PRODUCT_LABELS[s])
            for off, name in enumerate(SUBHEADERS):
                ws.cell(row=2, column=gc + off, value=name)

    r = ws.max_row + 1 if ws.max_row >= 2 else 3
    ws.cell(row=r, column=TIMESTAMP_COL, value=ts_str)

    n = len(PRODUCT_LABELS)
    for s in range(n):
        gc = FIRST_GROUP_COL + GROUP_STRIDE * s
        slot_sku = extract_sku_from_url(samsung_urls[s]) if s < len(samsung_urls) else None
        slot_sku = slot_sku.lower() if slot_sku else None

        res_by_site = {"amazon": am_res, "samsung": sam_res, "bestbuy": bb_res,
                       "lowes": lo_res, "walmart": wm_res}
        for off, site in enumerate(SITES):
            res_list = res_by_site[site]
            result = res_list[s] if s < len(res_list) else None
            price_cell, sku_cell = _render_cells(result, slot_sku, site)
            pc = ws.cell(row=r, column=gc + off, value=price_cell)
            if isinstance(price_cell, float):
                pc.number_format = "0.00"
            ws.cell(row=r, column=gc + SKU_OFFSET + off, value=sku_cell)

        # +10..+13 "vs" columns, added on every data row so they're in place as
        # rows accumulate. Each = <site> price / Samsung price - 1, for
        # Amazon, BestBuy, Lowes, Walmart (in that order).
        samsung_col = get_column_letter(gc + SITES.index("samsung"))
        for vs_i, site in enumerate(["amazon", "bestbuy", "lowes", "walmart"]):
            site_col = get_column_letter(gc + SITES.index(site))
            ws.cell(row=r, column=gc + VS_OFFSET + vs_i,
                    value=f"={site_col}{r}/{samsung_col}{r}-1")
        # +14 blank separator : intentionally left untouched

    wb.save(excel_path)
    print(f"✅ Results appended (WIP layout) to {excel_path} at row {r}")


def save_dict_to_excel_row(data: dict, excel_path: str = "outputs/results.xlsx"):
    """
    Save the provided dict as a single row in an Excel file.
    - Keys become column headers (first row).
    - Values become the next available row.
    - If the file exists, new keys are appended as new columns; existing column order is preserved.
    """
    os.makedirs(os.path.dirname(excel_path) or ".", exist_ok=True)
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        headers = list(data.keys())
        ws.append(headers)
        row = [ _to_jsonable(data.get(h)) for h in headers ]
        ws.append(row)
        wb.save(excel_path)
        print(f"✅ Results written to new Excel file: {excel_path}")
        return

    # file exists - load and append
    wb = load_workbook(excel_path)
    ws = wb.active

    # read existing headers from first row
    first_row = next(ws.iter_rows(min_row=1, max_row=1))
    existing_headers = [cell.value for cell in first_row]

    # compute headers union preserving existing order and appending new keys at the end
    new_keys = [k for k in data.keys() if k not in existing_headers]
    if new_keys:
        headers = existing_headers + new_keys
        # rewrite header row with expanded headers
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
    else:
        headers = existing_headers

    # build row in header order
    row = []
    for h in headers:
        v = data.get(h)
        row.append(_to_jsonable(v) if v is not None else None)

    ws.append(row)
    wb.save(excel_path)

    #making changes from here
    column_refs = [
        "blank","a","d","cf","ar","e","cg","as","blank","blank","blank","i","ck","aw","j","cl","ax","blank","blank","blank","n","cp","bb","o","cq","bc","blank","blank","blank","s","cu","bg","t","cv","bh","blank","blank","blank","x","cz","bl","y","da","bm","blank","blank","blank","ac","de","bq","ad","df","br","blank","blank","blank","ah","dj","bv","ai","dk","bw","blank","blank","blank","am","do","ca","an","dp","cb","blank","blank","blank","dt","gb","ex","du","gc","ey","blank","blank","blank","dy","gg","fc","dz","gh","fd","blank","blank","blank","ed","gl","fh","ee","gm","fi","blank","blank","blank","ei","gq","fm","ej","gr","fn","blank","blank","blank","en","gv","fr","eo","gw","fs","blank","blank","blank","es","ha","fw","et","hb","fx", 'blank', 'blank', 'blank', 'hf', 'kr', 'iy', 'hg', 'ks', 'iz', 'blank', 'blank', 'blank', 'hk', 'kw', 'jd', 'hl', 'kx', 'je', 'blank', 'blank', 'blank', 'hp', 'lb', 'ji', 'hq', 'lc', 'jj', 'blank', 'blank', 'blank', 'hu', 'lg', 'jn', 'hv', 'lh', 'jo', 'blank', 'blank', 'blank', 'hz', 'll', 'js', 'ia', 'lm', 'jt', 'blank', 'blank', 'blank', 'ie', 'lq', 'jx', 'if', 'lr', 'jy', 'blank', 'blank', 'blank', 'ij', 'lv', 'kc', 'ik', 'lw', 'kd', 'blank', 'blank', 'blank', 'io', 'ma', 'kh', 'ip', 'mb', 'ki', 'blank', 'blank', 'blank', 'it', 'mf', 'km', 'iu', 'mg', 'kn', 'blank', 'blank', 'blank', 'mk', 'xe', 'ru', 'ml', 'xf', 'rv', 'blank', 'blank', 'blank', 'mp', 'xj', 'rz', 'mq', 'xk', 'sa', 'blank', 'blank', 'blank', 'mu', 'xo', 'se', 'mv', 'xp', 'sf', 'blank', 'blank', 'blank', 'mz', 'xt', 'sj', 'na', 'xu', 'sk', 'blank', 'blank', 'blank', 'ne', 'xy', 'so', 'nf', 'xz', 'sp', 'blank', 'blank', 'blank', 'nj', 'yd', 'st', 'nk', 'ye', 'su', 'blank', 'blank', 'blank', 'no', 'yi', 'sy', 'np', 'yj', 'sz', 'blank', 'blank', 'blank', 'nt', 'yn', 'td', 'nu', 'yo', 'te', 'blank', 'blank', 'blank', 'ny', 'ys', 'ti', 'nz', 'yt', 'tj', 'blank', 'blank', 'blank', 'od', 'yx', 'tn', 'oe', 'yy', 'to', 'blank', 'blank', 'blank', 'oi', 'zc', 'ts', 'oj', 'zd', 'tt', 'blank', 'blank', 'blank', 'on', 'zh', 'tx', 'oo', 'zi', 'ty', 'blank', 'blank', 'blank', 'os', 'zm', 'uc', 'ot', 'zn', 'ud', 'blank', 'blank', 'blank', 'ox', 'zr', 'uh', 'oy', 'zs', 'ui', 'blank', 'blank', 'blank', 'pc', 'zw', 'um', 'pd', 'zx', 'un', 'blank', 'blank', 'blank', 'ph', 'aab', 'ur', 'pi', 'aac', 'us', 'blank', 'blank', 'blank', 'pm', 'aag', 'uw', 'pn', 'aah', 'ux', 'blank', 'blank', 'blank', 'pr', 'aal', 'vb', 'ps', 'aam', 'vc', 'blank', 'blank', 'blank', 'pw', 'aaq', 'vg', 'px', 'aar', 'vh', 'blank', 'blank', 'blank', 'qb', 'aav', 'vl', 'qc', 'aaw', 'vm', 'blank', 'blank', 'blank', 'qg', 'aba', 'vq', 'qh', 'abb', 'vr', 'blank', 'blank', 'blank', 'ql', 'abf', 'vv', 'qm', 'abg', 'vw', 'blank', 'blank', 'blank', 'qq', 'abk', 'wa', 'qr', 'abl', 'wb', 'blank', 'blank', 'blank', 'qv', 'abp', 'wf', 'qw', 'abq', 'wg', 'blank', 'blank', 'blank', 'ra', 'abu', 'wk', 'rb', 'abv', 'wl', 'blank', 'blank', 'blank', 'rf', 'abz', 'wp', 'rg', 'aca', 'wq', 'blank', 'blank', 'blank', 'rk', 'ace', 'wu', 'rl', 'acf', 'wv', 'blank', 'blank', 'blank', 'rp', 'acj', 'wz', 'rq', 'ack', 'xa',
    ]
    new_sheet_base_name="SelectedColumns"
    source_sheet_name=None
    # select source sheet
    if source_sheet_name:
        if source_sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{source_sheet_name}' not found in workbook.")
        src = wb[source_sheet_name]
    else:
        src = wb[wb.sheetnames[0]]

    # create unique new sheet name
    # new_name = new_sheet_base_name
    new_name = "converted"
    if new_name in wb.sheetnames:
        del wb["converted"]
    # i = 1
    # while new_name in wb.sheetnames:
    #     new_name = f"{new_sheet_base_name}_{i}"
    #     i += 1
    tgt = wb.create_sheet(title=new_name)
    # i = 1
    # while new_name in wb.sheetnames:
    #     new_name = f"{new_sheet_base_name}_{i}"
    #     i += 1
    # tgt = wb.create_sheet(title=new_name)

    max_row = src.max_row if src.max_row is not None else 0

    # target column pointer (1-indexed for openpyxl)
    tgt_col_idx = 1

    for token in column_refs:
        is_blank = token is None or (isinstance(token, str) and token.strip().lower() == "blank column")
        if is_blank:
            # leave a blank column (i.e., do nothing but advance tgt_col_idx)
            tgt_col_idx += 1
            continue

        # try to interpret token as Excel column letters
        col_letters = str(token).strip()
        try:
            src_col_idx = column_index_from_string(col_letters.upper())
        except Exception:
            # invalid column reference — create an empty column instead
            for r in range(1, max_row + 1):
                tgt.cell(row=r, column=tgt_col_idx, value=None)
            tgt_col_idx += 1
            continue

        # Copy values from source column to target column
        for r in range(1, max_row + 1):
            src_cell = src.cell(row=r, column=src_col_idx)
            # copy value only (not style/formula). If formula needed, assign src_cell.value (it will copy the formula text)
            tgt.cell(row=r, column=tgt_col_idx, value=src_cell.value)
        tgt_col_idx += 1
    wb.save(excel_path)
    print(f"✅ Results appended to Excel file: {excel_path}")





def copy_columns_by_references(
    file_path: str,
    column_refs: list,
    source_sheet_name: str | None = None,
    new_sheet_base_name: str = "CopiedColumns"
) -> str:
    """
    Copy columns from source sheet to a new sheet using Excel column letters.
    - column_refs: list of strings, column letters like ['A','D','X','AR', ...] or 'blank column'
    - source_sheet_name: None -> first sheet is used
    Returns the name of the created sheet.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    wb = load_workbook(file_path)
    # select source sheet
    if source_sheet_name:
        if source_sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{source_sheet_name}' not found in workbook.")
        src = wb[source_sheet_name]
    else:
        src = wb[wb.sheetnames[0]]

    # create unique new sheet name
    new_name = "converted"
    if new_name in wb.sheetnames:
        del wb["converted"]
    # i = 1
    # while new_name in wb.sheetnames:
    #     new_name = f"{new_sheet_base_name}_{i}"
    #     i += 1
    tgt = wb.create_sheet(title=new_name)

    max_row = src.max_row if src.max_row is not None else 0

    # target column pointer (1-indexed for openpyxl)
    tgt_col_idx = 1

    for token in column_refs:
        is_blank = token is None or (isinstance(token, str) and token.strip().lower() == "blank column")
        if is_blank:
            # leave a blank column (i.e., do nothing but advance tgt_col_idx)
            tgt_col_idx += 1
            continue

        # try to interpret token as Excel column letters
        col_letters = str(token).strip()
        try:
            src_col_idx = column_index_from_string(col_letters.upper())
        except Exception:
            # invalid column reference — create an empty column instead
            for r in range(1, max_row + 1):
                tgt.cell(row=r, column=tgt_col_idx, value=None)
            tgt_col_idx += 1
            continue

        # Copy values from source column to target column
        for r in range(1, max_row + 1):
            src_cell = src.cell(row=r, column=src_col_idx)
            # copy value only (not style/formula). If formula needed, assign src_cell.value (it will copy the formula text)
            tgt.cell(row=r, column=tgt_col_idx, value=src_cell.value)
        tgt_col_idx += 1

    # Save workbook (overwrites existing file)
    wb.save(file_path)
    return new_name


# -----------------------
# APIFY (Amazon + BestBuy + Lowes)
# -----------------------
# Amazon, BestBuy and Lowes are fetched through Apify actors instead of our
# own browser. Each site is ONE actor run that takes every URL at once:
#   Amazon : delicious_zebu/amazon-product-details-scraper  (APIFY_AMAZON_ACTOR)
#   BestBuy: benthepythondev/bestbuy-scraper                (APIFY_BESTBUY_ACTOR)
#   Lowes  : dami_studio/lowes-scraper                      (APIFY_LOWES_ACTOR)
# (Samsung and Walmart are still scraped with Playwright.)
#
# Tokens: env var APIFY_API_TOKENS (a GitHub Actions secret), several tokens
# separated by commas or newlines. They are used strictly in order; a token is
# dropped and the next one used when its account can't run the actor:
#   - before first use we read the account's monthly credit (/users/me/limits)
#     and skip the token if less than APIFY_MIN_REMAINING_USD is left;
#   - if Apify rejects a call with one of the credit/permission errors below
#     (HTTP 401/402/403), we switch to the next token and start again;
#   - if a run ends early (e.g. credit ran out mid-run) its credit is re-checked
#     and the URLs that got no data are re-run, on the next token if needed.
#
# Each per-URL result keeps the same shape the Excel writer expects:
#   {"url", "file", "price", "model", "status"}
#   price = price text        -> written as a number
#   price = NOT_AVAILABLE     -> redirect / unavailable / used-only / no price
#   price = None              -> no data at all (empty slot / Apify failed) -> blank

APIFY_API_BASE = "https://api.apify.com/v2"
APIFY_AMAZON_ACTOR = "U3DyJ7kdhQlYyeQKd"    # delicious_zebu/amazon-product-details-scraper
APIFY_BESTBUY_ACTOR = "pbUZ4z2ORsyKhZshL"   # benthepythondev/bestbuy-scraper
APIFY_LOWES_ACTOR = "Hchq3U0V6uJTxJax0"     # dami_studio/lowes-scraper

APIFY_RUN_TIMEOUT_SEC = 900     # Apify aborts a run that takes longer than this
APIFY_POLL_WAIT_SEC = 60        # each status poll blocks up to this long (Apify max 60)
APIFY_ROUNDS = 3                # 1 run for all URLs + up to 2 re-runs for URLs with no data
APIFY_MIN_REMAINING_USD = 0.20  # skip a token with less monthly credit than this left
APIFY_HTTP_RETRIES = 3          # per API call, for network errors / HTTP 429 / 5xx

# Apify error "type" values meaning THIS token/account can't run the actor, so
# the next token should be tried. Credit exhaustion shows up as:
#   403 platform-feature-disabled          "Monthly usage hard limit exceeded"
#   402 not-enough-usage-to-run-paid-actor  (not enough credit left to start)
# Any other 401/402/403 is treated the same way (see ApifyError.token_unusable).
APIFY_TOKEN_ERROR_TYPES = {
    "platform-feature-disabled",
    "not-enough-usage-to-run-paid-actor",
    "monthly-usage-limit-too-low",
    "limit-reached",
    "x402-payment-required",
    "apify-plan-required-to-use-paid-actor",
    "user-has-no-subscription",
    "actor-is-not-rented",
    "full-permission-actor-not-approved",
    "full-permission-actor-blocked-for-admin",
    "elevated-permissions-needed",
    "insufficient-permissions",
    "actor-memory-limit-exceeded",
    "concurrent-runs-limit-exceeded",
    "invalid-token",
    "token-not-provided",
    "user-or-token-not-found",
    "user-disabled",
}

_APIFY_TERMINAL = {"SUCCEEDED", "FAILED", "ABORTED", "TIMED-OUT"}


class ApifyError(Exception):
    """An error response from the Apify API."""

    def __init__(self, status, err_type, message):
        super().__init__(f"HTTP {status} {err_type or '?'}: {message}")
        self.status = status
        self.type = err_type or ""
        self.message = message or ""

    @property
    def token_unusable(self):
        """True when switching to another token may fix it (credit / permission)."""
        return self.status in (401, 402, 403) or self.type in APIFY_TOKEN_ERROR_TYPES


def _apify_call(method, path, token, body=None, params=None, timeout=90):
    """One Apify API call. Returns the parsed JSON body.

    Network errors, HTTP 429 and HTTP 5xx are retried (APIFY_HTTP_RETRIES);
    any other error status raises ApifyError straight away.
    """
    url = APIFY_API_BASE + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    data = json.dumps(body).encode("utf-8") if body is not None else None
    last_err = None
    for attempt in range(1, APIFY_HTTP_RETRIES + 1):
        req = urllib.request.Request(url, data=data, method=method, headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        })
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8")
                return json.loads(raw) if raw else None
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", "replace")
            try:
                err = (json.loads(raw) or {}).get("error") or {}
            except Exception:
                err = {}
            api_err = ApifyError(e.code, err.get("type"), err.get("message") or raw[:300])
            if e.code != 429 and e.code < 500:
                raise api_err
            last_err = api_err
        except (urllib.error.URLError, OSError, ValueError) as e:
            last_err = e
        if attempt < APIFY_HTTP_RETRIES:
            time.sleep(5 * attempt)
    raise last_err


def _apify_credit_check(token):
    """Return (usable, note) for a token based on its account's monthly credit.

    Only a definite answer disqualifies a token: an invalid token (401) or an
    account whose usage hard limit is already hit / nearly hit. If the limits
    can't be read for any other reason we still try the token.
    """
    try:
        data = (_apify_call("GET", "/users/me/limits", token) or {}).get("data") or {}
    except ApifyError as e:
        if e.status == 401 or e.type == "platform-feature-disabled":
            return False, str(e)
        return True, f"credit not readable ({e}), trying anyway"
    except Exception as e:
        return True, f"credit not readable ({e}), trying anyway"
    max_usd = (data.get("limits") or {}).get("maxMonthlyUsageUsd")
    used_usd = (data.get("current") or {}).get("monthlyUsageUsd")
    if isinstance(max_usd, (int, float)) and isinstance(used_usd, (int, float)) and max_usd > 0:
        left = max_usd - used_usd
        if left < APIFY_MIN_REMAINING_USD:
            return False, f"only ${left:.2f} of ${max_usd:.2f} monthly credit left"
        return True, f"${left:.2f} of ${max_usd:.2f} monthly credit left"
    return True, "credit unknown"


class ApifyTokenPool:
    """The Apify tokens for this run, used in order until each is used up."""

    def __init__(self, tokens):
        self.tokens = tokens
        self.idx = 0            # index of the token currently in use
        self.checked = set()    # indexes whose credit has been checked

    @classmethod
    def from_env(cls):
        raw = os.environ.get("APIFY_API_TOKENS") or os.environ.get("APIFY_API_TOKEN") or ""
        tokens = [t for t in re.split(r"[\s,;]+", raw) if t]
        if not tokens:
            print("⚠️ APIFY_API_TOKENS is not set -> Amazon, BestBuy and Lowes will be left blank")
        else:
            print(f"🔑 {len(tokens)} Apify token(s) loaded")
        return cls(tokens)

    def label(self):
        t = self.tokens[self.idx]
        return f"token #{self.idx + 1} (...{t[-4:]})"

    def current(self):
        """Token to use now (credit-checked on first use), or None if all are used up."""
        while self.idx < len(self.tokens):
            if self.idx not in self.checked:
                self.checked.add(self.idx)
                usable, note = _apify_credit_check(self.tokens[self.idx])
                if not usable:
                    print(f"⚠️ Apify {self.label()} skipped: {note}")
                    self.idx += 1
                    continue
                print(f"🔑 using Apify {self.label()}: {note}")
            return self.tokens[self.idx]
        return None

    def drop_current(self, reason):
        print(f"⚠️ Apify {self.label()} can't be used ({reason}) -> switching to the next token")
        self.idx += 1

    def recheck_current(self):
        """Re-check the current token's credit before its next use."""
        self.checked.discard(self.idx)


def _apify_run_actor(pool, actor_id, actor_input, site):
    """Run an actor once and return (items, run_status).

    Uses the first usable token, moving to the next one on a credit/permission
    error. Returns (None, reason) if the run could not be done at all.
    """
    while True:
        token = pool.current()
        if token is None:
            return None, "no usable Apify token left"
        who = pool.label()
        try:
            print(f"🚀 [{site}] starting Apify actor {actor_id} with {who} ...")
            run = _apify_call("POST", f"/acts/{actor_id}/runs", token, body=actor_input,
                              params={"timeout": APIFY_RUN_TIMEOUT_SEC})["data"]
            run_id = run["id"]
            # Apify stops the run itself at APIFY_RUN_TIMEOUT_SEC; the extra
            # margin only guards against a run stuck in a non-final state.
            deadline = time.monotonic() + APIFY_RUN_TIMEOUT_SEC + 300
            while run.get("status") not in _APIFY_TERMINAL and time.monotonic() < deadline:
                run = _apify_call("GET", f"/actor-runs/{run_id}", token,
                                  params={"waitForFinish": APIFY_POLL_WAIT_SEC})["data"]
            status = run.get("status")
            print(f"[{site}] Apify run {run_id} -> {status} {run.get('statusMessage') or ''}")
            items = _apify_call("GET", f"/datasets/{run['defaultDatasetId']}/items", token,
                                params={"clean": "true", "format": "json"})
            return (items if isinstance(items, list) else []), status
        except ApifyError as e:
            if e.token_unusable:
                pool.drop_current(e)
                continue
            print(f"❌ [{site}] Apify call failed: {e}")
            return None, str(e)
        except Exception as e:
            print(f"❌ [{site}] Apify call failed: {e}")
            return None, str(e)


def _apify_collect(pool, site, actor_id, keyed_urls, build_input, item_key, item_done):
    """Fetch every URL through the actor, re-running only what's still missing.

    keyed_urls: {key: url} for the non-empty slots.
    Returns {key: item} for the keys the actor returned data for.
    """
    found = {}
    for rnd in range(1, APIFY_ROUNDS + 1):
        todo = [u for k, u in keyed_urls.items() if k not in found or not item_done(found[k])]
        if not todo:
            break
        print(f"\n[{site}] Apify round {rnd}/{APIFY_ROUNDS}: {len(todo)} URL(s)")
        items, status = _apify_run_actor(pool, actor_id, build_input(todo), site)
        if items is None:
            if pool.current() is None:
                print(f"❌ [{site}] all Apify tokens used up -> remaining URLs left blank")
                break
            continue
        for it in items:
            k = item_key(it) if isinstance(it, dict) else None
            if k in keyed_urls and (k not in found or not item_done(found[k])):
                found[k] = it
        if status != "SUCCEEDED":
            pool.recheck_current()   # it may have run out of credit mid-run
    return found


def _save_apify_items(items_by_key, site, output_dir):
    """Keep the raw actor output for this run (replaces the old saved HTML pages)."""
    try:
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, f"apify_{site.lower()}_items.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(items_by_key, f, ensure_ascii=False, indent=2)
        print(f"✅ raw {site} data saved to {path}")
        return path
    except Exception as e:
        print(f"⚠️ could not save raw {site} data: {e}")
        return None


# -----------------------
# AMAZON-specific logic (via Apify)
# -----------------------
def _amazon_item_is_used(item):
    """True when the buybox offer is a USED / renewed device (we only track NEW).

    The actor has no condition field; Amazon's used/warehouse seller is
    "Amazon Resale", and renewed listings say so in the seller or title.
    """
    seller = " ".join(str(item.get(k) or "") for k in ("seller_name", "ships_from"))
    title = str(item.get("title") or "")
    return bool(re.search(r"\b(resale|renewed|used)\b", seller, re.I)
                or re.search(r"\brenewed\b", title, re.I))


def _amazon_item_price(item):
    """Buybox price as text, or None."""
    pv = item.get("price_value")
    if isinstance(pv, (int, float)) and pv > 0:
        return f"{pv:.2f}"
    num = clean_price_value(item.get("price"))
    return f"{num:.2f}" if num and num > 0 else None


def _amazon_item_done(item):
    """Whether this record is a final answer (else the URL is re-run)."""
    return bool(_amazon_item_price(item)
                or _looks_unavailable(str(item.get("availability") or ""), "amazon")
                or _amazon_item_is_used(item))


def fetch_amazon_via_apify(urls, pool, output_dir="outputs"):
    """Amazon price per URL via the Apify actor. Returns results in URL order."""
    keyed = {u: u for u in dict.fromkeys(u.strip() for u in urls if u and u.strip())}
    asin_to_url = {amazon_id_from_url(u): u for u in keyed if amazon_id_from_url(u)}

    def build_input(todo):
        return {"Params": todo, "deliverTo": "US", "zipCode": "10001"}

    def item_key(item):
        # search_source echoes the URL we sent; fall back to the ASIN
        src = str(item.get("search_source") or "").strip()
        if src in keyed:
            return src
        return asin_to_url.get(str(item.get("asin") or "").upper())

    found = {}
    if keyed and pool.tokens:
        found = _apify_collect(pool, "Amazon", APIFY_AMAZON_ACTOR, keyed,
                               build_input, item_key, _amazon_item_done)
        _save_apify_items(found, "Amazon", output_dir)

    results = []
    for idx, url in enumerate(urls, start=1):
        if not url or not url.strip():
            print(f"[Amazon {idx}/{len(urls)}] empty URL slot -> skipping")
            results.append({"url": url, "file": None, "price": None, "model": None, "status": "empty"})
            continue
        item = found.get(url.strip())
        if item is None:
            print(f"[Amazon {idx}/{len(urls)}] ❌ no data from Apify -> left blank")
            results.append({"url": url, "file": None, "price": None, "model": None, "status": "error: no data from Apify"})
            continue

        expected = amazon_id_from_url(url)
        got = str(item.get("asin") or "").upper() or None
        price = _amazon_item_price(item)
        if expected and got and got != expected:
            print(f"[Amazon {idx}/{len(urls)}] [REDIRECT] requested {expected} but got {got} -> not available")
            result = {"price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "redirect"}
        elif _amazon_item_is_used(item):
            print(f"[Amazon {idx}/{len(urls)}] buybox is a USED offer (seller {item.get('seller_name')!r}) -> not available")
            result = {"price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "used_offer"}
        elif price:
            print(f"[Amazon {idx}/{len(urls)}] ✅ price {price} ({item.get('availability')})")
            # SKU_Amazon is filled by the writer from the slot's Samsung SM code
            result = {"price": price, "model": None, "status": "ok"}
        else:
            print(f"[Amazon {idx}/{len(urls)}] no price ({item.get('availability')!r}) -> not available")
            result = {"price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "no_price"}
        results.append({"url": url, "file": None, **result})
    return results


# -----------------------
# BESTBUY-specific logic (via Apify)
# -----------------------
def _bestbuy_item_price(item):
    """Current selling price as text, or None."""
    num = clean_price_value(item.get("price"))
    return f"{num:.2f}" if num and num > 0 else None


def _bestbuy_item_is_new(item):
    cond = str(item.get("condition") or "new").strip().lower()
    return cond == "new" and not item.get("openBoxCondition")


def _bestbuy_item_done(item):
    return bool(_bestbuy_item_price(item)) or not _bestbuy_item_is_new(item)


def fetch_bestbuy_via_apify(urls, pool, output_dir="outputs"):
    """BestBuy price per URL via the Apify actor. Returns results in URL order.

    URLs are matched by BestBuy's product code (e.g. JJGRF3TZF2), which the
    actor keeps in the resolved URL it returns (".../JJGRF3TZF2/sku/6681685").
    """
    keyed = {}
    for u in urls:
        code = bestbuy_id_from_url(u) if u and u.strip() else None
        if code:
            keyed.setdefault(code, u.strip())

    def build_input(todo):
        return {"mode": "direct_urls", "productUrls": todo, "maxProducts": len(todo)}

    def item_key(item):
        return bestbuy_id_from_url(str(item.get("url") or ""))

    found = {}
    if keyed and pool.tokens:
        found = _apify_collect(pool, "BestBuy", APIFY_BESTBUY_ACTOR, keyed,
                               build_input, item_key, _bestbuy_item_done)
        _save_apify_items(found, "BestBuy", output_dir)

    results = []
    for idx, url in enumerate(urls, start=1):
        if not url or not url.strip():
            print(f"[BestBuy {idx}/{len(urls)}] empty URL slot -> skipping")
            results.append({"url": url, "file": None, "price": None, "model": None, "status": "empty"})
            continue
        item = found.get(bestbuy_id_from_url(url))
        if item is None:
            print(f"[BestBuy {idx}/{len(urls)}] ❌ no data from Apify -> left blank")
            results.append({"url": url, "file": None, "price": None, "model": None, "status": "error: no data from Apify"})
            continue

        price = _bestbuy_item_price(item)
        model = item.get("modelNumber")
        if not _bestbuy_item_is_new(item):
            print(f"[BestBuy {idx}/{len(urls)}] offer is not new ({item.get('condition')}/{item.get('openBoxCondition')}) -> not available")
            result = {"price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "not_new"}
        elif price:
            print(f"[BestBuy {idx}/{len(urls)}] ✅ price {price} (model {model})")
            result = {"price": price, "model": model, "status": "ok"}
        else:
            print(f"[BestBuy {idx}/{len(urls)}] no price -> not available")
            result = {"price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "no_price"}
        results.append({"url": url, "file": None, **result})
    return results


# -----------------------
# LOWES-specific logic (via Apify)
# -----------------------
# Actor dami_studio/lowes-scraper (APIFY_LOWES_ACTOR). One row per product:
#   productId  = the number at the end of the /pd/ url (used to match rows to URLs)
#   price      = current selling price (number); listPrice = the "was" price,
#                which is never used here
#   modelNumber, priceScope (STORE / NETWORK / SOS), inStock, ...
# Rows with a "charged" field (charged: false) are NOT products: the free
# sample row and diagnostic/note rows (errorCode REFUSED, PRODUCT_NOT_FOUND,
# ...). They are skipped. No storeNumber is passed, so the actor uses its
# default store (0595), the same store on every run.
def _lowes_item_is_product(item):
    return "charged" not in item and bool(item.get("productId"))


def _lowes_item_price(item):
    """Current selling price as text, or None."""
    num = clean_price_value(item.get("price"))
    return f"{num:.2f}" if num and num > 0 else None


def _lowes_item_done(item):
    # A real product row is a final answer, even with price null (the actor
    # returns null for special-order / quote-only items).
    return _lowes_item_is_product(item)


def fetch_lowes_via_apify(urls, pool, output_dir="outputs"):
    """Lowes price per URL via the Apify actor. Returns results in URL order.

    URLs are matched by Lowes' product id (the number at the end of the /pd/
    url, e.g. 5013373117), which the actor returns as "productId".
    """
    keyed = {}
    for u in urls:
        pid = lowes_id_from_url(u) if u and u.strip() else None
        if pid:
            keyed.setdefault(pid, u.strip())

    def build_input(todo):
        # searchQueries=[] explicitly: the actor's console prefills "cordless
        # drill"; an empty list guarantees only our URLs are scraped and billed.
        return {"startUrls": todo, "searchQueries": [], "maxItems": len(todo),
                "detailLevel": "listing"}

    def item_key(item):
        return str(item.get("productId")) if _lowes_item_is_product(item) else None

    found = {}
    if keyed and pool.tokens:
        found = _apify_collect(pool, "Lowes", APIFY_LOWES_ACTOR, keyed,
                               build_input, item_key, _lowes_item_done)
        _save_apify_items(found, "Lowes", output_dir)

    results = []
    for idx, url in enumerate(urls, start=1):
        if not url or not url.strip():
            print(f"[Lowes {idx}/{len(urls)}] empty URL slot -> skipping")
            results.append({"url": url, "file": None, "price": None, "model": None, "status": "empty"})
            continue
        item = found.get(lowes_id_from_url(url))
        if item is None:
            print(f"[Lowes {idx}/{len(urls)}] ❌ no data from Apify -> left blank")
            results.append({"url": url, "file": None, "price": None, "model": None, "status": "error: no data from Apify"})
            continue

        price = _lowes_item_price(item)
        model = item.get("modelNumber")
        if price:
            print(f"[Lowes {idx}/{len(urls)}] ✅ price {price} (model {model}, priceScope {item.get('priceScope')})")
            result = {"price": price, "model": model, "status": "ok"}
        else:
            print(f"[Lowes {idx}/{len(urls)}] no price -> not available")
            result = {"price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "no_price"}
        results.append({"url": url, "file": None, **result})
    return results


# -----------------------
# Real Chrome over CDP (used by the Walmart scraper)
# -----------------------
# Originally written for the old browser-based BestBuy scraper (hence the
# BestBuy wording in the docstrings); BestBuy and Lowes now use Apify, and
# these are kept unchanged for Walmart.
def _find_free_port():
    """Return an OS-assigned free TCP port on localhost."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _find_real_chrome():
    """Locate the REAL Google Chrome binary (NOT Playwright's bundled Chromium).

    BestBuy's bot protection is confirmed to pass with real Chrome; bundled
    Chromium differs in fingerprint (UA-brand "Chromium", no Widevine/H.264,
    different userAgentData) and may be flagged. So for BestBuy we insist on
    real Chrome.

    Resolution order:
      1. $CHROME_BIN / $CHROME_PATH env var (set this on EC2 if Chrome is in a
         non-standard location).
      2. `chrome`/`google-chrome`/`google-chrome-stable` on PATH (Linux; the
         GitHub Actions ubuntu runner ships google-chrome-stable).
      3. Standard Windows install locations (for local Windows 11 runs).
    Returns the path, or None if not found.
    """
    env = os.environ.get("CHROME_BIN") or os.environ.get("CHROME_PATH")
    if env and os.path.exists(env):
        return env

    for name in ("google-chrome-stable", "google-chrome", "chrome",
                 "chromium-browser", "chromium"):
        found = shutil.which(name)
        if found:
            return found

    for candidate in (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/opt/google/chrome/chrome",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    ):
        if candidate and os.path.exists(candidate):
            return candidate

    return None


async def _launch_chrome_cdp(playwright, extra_args=None):
    """Launch a REAL Chrome/Chromium process with a remote-debugging port and
    connect Playwright to it over CDP.

    Why CDP instead of playwright.chromium.launch() for BestBuy: launch() starts
    Chromium with Playwright's automation switches (e.g. --enable-automation,
    AutomationControlled), which BestBuy's bot protection fingerprints. By
    starting Chrome ourselves with only the flags we choose and attaching over
    the DevTools Protocol, the browser looks like an ordinary Chrome instance.

    Each call launches a fresh process with a fresh throwaway --user-data-dir, so
    every attempt gets a brand-new HTTP/2 connection AND a pristine, cookieless
    profile (the two things that previously broke URL 2+).

    Returns (browser, proc, profile_dir). Caller must close browser, kill proc,
    and remove profile_dir.
    """
    port = _find_free_port()
    profile_dir = tempfile.mkdtemp(prefix="bb_cdp_profile_")
    # Use the REAL Google Chrome binary (confirmed to pass BestBuy). Do NOT fall
    # back to bundled Chromium — its fingerprint differs and may be flagged.
    chrome_path = _find_real_chrome()
    if not chrome_path:
        shutil.rmtree(profile_dir, ignore_errors=True)
        raise RuntimeError(
            "Real Google Chrome not found. Install it (Actions ubuntu runner "
            "ships google-chrome-stable; on EC2 install google-chrome-stable) "
            "or set the CHROME_BIN env var to its path.")

    args = [
        chrome_path,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "about:blank",
    ]
    if extra_args:
        args.extend(extra_args)

    # Headful: BestBuy won't serve a headless browser. On a display-less server
    # (EC2 / GitHub Actions) this runs under Xvfb, which supplies $DISPLAY.
    proc = subprocess.Popen(
        args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # Wait for the CDP endpoint to come up, then read the WebSocket URL.
    cdp_http = f"http://127.0.0.1:{port}"
    ws_url = None
    for _ in range(60):  # up to ~30s
        try:
            with urllib.request.urlopen(f"{cdp_http}/json/version", timeout=1) as r:
                ws_url = json.loads(r.read().decode()).get("webSocketDebuggerUrl")
            if ws_url:
                break
        except Exception:
            await asyncio.sleep(0.5)
    if not ws_url:
        # cleanup before raising
        try:
            proc.kill()
        except Exception:
            pass
        shutil.rmtree(profile_dir, ignore_errors=True)
        raise RuntimeError("Chrome CDP endpoint did not come up in time")

    browser = await playwright.chromium.connect_over_cdp(ws_url)
    return browser, proc, profile_dir


# -----------------------
# SAMSUNG-specific logic
# -----------------------
async def wait_network_idle(page, timeout=15000):
    """Wait until network becomes idle (0 active requests)."""
    try:
        await page.wait_for_load_state("networkidle", timeout=timeout)
        await asyncio.sleep(10)  # extra wait to ensure stability | not sure if this is the right approach
    except TimeoutError:
        print("⚠️ networkidle timeout — continuing anyway")

def extract_sku_from_url(url: str):
    """Extract SKU value from the given URL (looks for 'sku-<value>' or 'sm-<value>')."""
    if not url:
        return None
    
    # Try to find sku- first
    m = re.search(r"sku-([A-Za-z0-9-]+)", url, re.IGNORECASE)
    if m:
        return m.group(1)
    
    # If not found, try to find sm-
    m = re.search(r"(sm-[A-Za-z0-9-]+)", url, re.IGNORECASE)
    if m:
        return m.group(1)
    
    return None


def extract_price(filename, expected_url=None):
    """Return (price_text_or_None, redirected_bool) for a saved Samsung page.

    Updated for the current UI: the old #device_info aria-checked radios are gone.
    The reliable source is the JSON-LD Product offer, keyed to the exact SKU, so
    we never pick up a sibling variant's price. Redirects are detected via the
    page's canonical link.
    """
    if not os.path.exists(filename):
        print(f"❌ File not found for parsing: {filename}")
        return None, False

    with open(filename, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()

    expected_sku = extract_sku_from_url(expected_url) if expected_url else None
    expected_sku = expected_sku.lower() if expected_sku else None

    # -------- REDIRECT DETECTION --------
    canonical = get_canonical_href(html)
    if expected_sku and canonical:
        can_sku = extract_sku_from_url(canonical)
        can_sku = can_sku.lower() if can_sku else None
        if can_sku and can_sku != expected_sku:
            print(f"[REDIRECT] requested {expected_sku} but page is {can_sku} -> not available")
            return None, True

    # -------- PRICE from JSON-LD offer keyed to the SKU --------
    # Recorded even when the offer is OutOfStock (Samsung still lists a price).
    price = None
    for off in _samsung_sku_offers(html, expected_sku):
        if off.get("price"):
            price = str(off["price"]); break

    print("🔎 Extracted Price:", price)
    return price, False


def _samsung_sku_offers(html, expected_sku=None):
    """Yield the JSON-LD offer dicts whose product sku matches expected_sku
    (any sku when expected_sku is None).

    The sku/offers pair is either a top-level Product (phones, appliances) or a
    variant nested in a ProductGroup's "hasVariant" list (TV pages), so both
    are checked.
    """
    for it in iter_ldjson(html):
        if not isinstance(it, dict):
            continue
        variants = it.get("hasVariant")
        candidates = [it] + (variants if isinstance(variants, list) else [])
        for c in candidates:
            if not isinstance(c, dict) or not c.get("sku"):
                continue
            if expected_sku and str(c["sku"]).lower() != expected_sku:
                continue
            off = c.get("offers")
            if isinstance(off, dict):
                yield off


def _samsung_offer_unavailable(html, expected_url):
    """True when the JSON-LD offer for THIS sku says it can't be bought
    (OutOfStock / SoldOut / Discontinued).

    Used instead of _looks_unavailable(html, "samsung"): that searches the whole
    page for "out of stock" / "notify me" / "sold out", and every Samsung page
    (in-stock ones too) contains those words in its built-in message templates,
    e.g. "{capacity} in {color} is out of stock" -> it is always True.
    """
    sku = (extract_sku_from_url(expected_url) or "").lower()
    for off in _samsung_sku_offers(html, sku):
        avail = str(off.get("availability", "")).lower()
        return any(k in avail for k in ("outofstock", "soldout", "discontinued"))
    return False

async def save_samsung_htmls(
    urls,
    output_dir="outputs",
    cookies_file="samsung_cookies.json",
    headless=True,
):
    """
    Loop over list of Samsung product URLs, save each page's HTML to output_dir,
    parse price and sku using the same logic you provided, and return results list.
    """
    os.makedirs(output_dir, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)   # keep visible by default per original
        # Create or reuse context
        # if os.path.exists(cookies_file):
        #     print("🍪 Loading existing cookies/session...")
        #     context = await browser.new_context(storage_state=cookies_file)
        # else:
        print("🆕 No cookies found, creating a new session...")

        results = []
        try:
            for idx, url in enumerate(urls, start=1):
                # empty slot (e.g. product not yet listed): keep the position so
                # results stay aligned with the product groups, but skip cleanly.
                if not url or not url.strip():
                    print(f"\n[Samsung {idx}/{len(urls)}] empty URL slot -> skipping")
                    results.append({"url": url, "file": None, "price": None, "sku": None, "status": "empty"})
                    continue
                safe_name = sanitize_filename(url)
                output_file = os.path.join(output_dir, f"samsung_{idx}_{safe_name}.html")
                sku = extract_sku_from_url(url)

                # Retry transient failures (nav error / timeout / price not
                # rendered). A redirect or a JSON-LD offer marked out of stock /
                # discontinued is final and is NOT retried.
                result = None
                for attempt in range(1, MAX_ATTEMPTS + 1):
                    context = None
                    page = None
                    try:
                        # fresh context per attempt (isolates cookies/storage)
                        context = await browser.new_context(
                            user_agent=(
                                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                "AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/120.0.0.0 Safari/537.36"
                            ),
                            viewport={"width": 1600, "height": 900},
                        )

                        page = await context.new_page()
                        print(f"\n[Samsung {idx}/{len(urls)}] (attempt {attempt}/{MAX_ATTEMPTS}) Navigating to {url} ...")
                        try:
                            # 30s hard cap so a slow page can't stall the whole run.
                            await page.goto(url, wait_until="load", timeout=30000)
                        except TimeoutError:
                            print(f"⚠️ navigation timeout for {url} after 30s. Continuing anyway...")

                        print("Waiting for network to be idle...")
                        await wait_network_idle(page, timeout=20000)

                        # The price is read from the JSON-LD Product offer (see
                        # extract_price), so wait for that. (The old wait for
                        # #device_info is gone: that box exists only on Galaxy
                        # phone pages, never on appliance pages like refrigerators,
                        # so it always timed out and the page was never parsed.)
                        print("Waiting for product JSON-LD offer...")
                        try:
                            await page.wait_for_function(
                                """() => {
                                    const s = document.querySelectorAll('script[type="application/ld+json"]');
                                    for (const el of s) {
                                        if (el.textContent && el.textContent.indexOf('"offers"') !== -1) return true;
                                    }
                                    return false;
                                }""",
                                timeout=20000,
                            )
                        except Exception:
                            print("⚠️ product JSON-LD offer not detected within 20s; parsing saved page anyway")

                        # Save HTML (resilient to any mid-load client-side navigation)
                        html = await get_page_content_safe(page)
                        with open(output_file, "w", encoding="utf-8") as f:
                            f.write(html)
                        print(f"✅ HTML saved to {output_file}")

                        # Parse saved HTML (updated: redirect-aware, JSON-LD by SKU;
                        # a price is recorded even if the offer is out of stock)
                        price, redirected = extract_price(output_file, expected_url=url)
                        if redirected:
                            print("[REDIRECT] Samsung redirect -> not available")
                            result = {"url": url, "file": output_file, "price": NOT_AVAILABLE, "sku": NOT_AVAILABLE, "status": "redirect"}
                            break  # final: different product
                        elif price:
                            print("🔎 Final extracted values — Price:", price, "SKU:", sku)
                            result = {"url": url, "file": output_file, "price": price, "sku": sku, "status": "ok"}
                            break  # final: got a price
                        else:
                            result = {"url": url, "file": output_file, "price": NOT_AVAILABLE, "sku": NOT_AVAILABLE, "status": "no_price"}
                            if _samsung_offer_unavailable(html, url):
                                print("ℹ️ JSON-LD offer for this SKU is out of stock / discontinued -> final, not retrying")
                                break
                            print(f"⚠️ price not found & page not marked unavailable (attempt {attempt}/{MAX_ATTEMPTS})")

                        # tiny cooperative yield
                        await human_delay_short()
                    except Exception as e:
                        print(f"❌ Error processing URL {url} (attempt {attempt}/{MAX_ATTEMPTS}): {e}")
                        result = {"url": url, "file": None, "price": None, "sku": None, "status": f"error: {e}"}
                    finally:
                        try:
                            if page:
                                await page.close()
                        except Exception:
                            pass
                        try:
                            if context:
                                await context.close()
                        except Exception:
                            pass
                    if attempt < MAX_ATTEMPTS:
                        print(f"🔁 retrying in {RETRY_BACKOFF_SEC}s ...")
                        await asyncio.sleep(RETRY_BACKOFF_SEC)

                results.append(result)

            # Write cookies/session state once more at the end
            # storage = await context.storage_state()
            # with open(cookies_file, "w", encoding="utf-8") as f:
            #     json.dump(storage, f, indent=2)
            # print(f"\n🍪 Cookies/session state written to {cookies_file}")

        finally:
            await browser.close()

    return results

# -----------------------
# WALMART logic
# -----------------------
# (Lowes used to share this browser flow; it now comes from Apify, see the
# LOWES-specific logic section.)
# Walmart is scraped with the flow the old browser-based Amazon scraper used
# (load -> wait -> human-like
# mouse/scroll -> save HTML -> parse saved HTML -> retry transient failures),
# with the same timings: goto wait_until="load" / 30s cap, 10s settle wait,
# 3-6s random wait, 20s retry backoff, 20s gap between URLs.
#
# Differences from that Amazon flow (verified against the live US site):
#   - Fresh session for EVERY url AND every retry: each attempt launches a new
#     real Chrome process with a brand-new throwaway profile (via
#     _launch_chrome_cdp, the BestBuy launcher). Nothing is shared between
#     URLs and no cookies are ever saved.
#   - Browser: Playwright's own launch (bundled Chromium or Chrome, headless,
#     Amazon's custom user agent) is blocked: Walmart returns its "Robot or
#     human?" page. Even real Chrome is blocked when headless. Only HEADFUL real Chrome attached over CDP
#     passes, so that's what is used. On a display-less server run under Xvfb
#     (same requirement as BestBuy).
#   - No custom user agent: real Chrome sends its own, matching UA; overriding it
#     with Amazon's older Chrome/120 string would contradict Chrome's other
#     headers and look more bot-like.

def _page_title(html):
    m = re.search(r"<title[^>]*>(.*?)</title>", html or "", re.I | re.S)
    return m.group(1).strip() if m else ""


def _looks_blocked(html, site):
    """True when the saved page is a bot-check / block page instead of the
    product page. Treated as TRANSIENT (retried with a fresh session)."""
    title = _page_title(html).lower()
    if site == "walmart":
        return "robot or human" in title
    return False


def parse_walmart_html(html_file_path, expected_url=None):
    """Return (price_text_or_None, redirected_bool, unavailable_bool).

    Walmart embeds the full page data in <script id="__NEXT_DATA__">. The main
    product is props.pageProps.initialData.data.product; its
    priceInfo.currentPrice.price is the buy-box price. The page also holds
    "currentPrice" entries for recommended items, so we only read the main
    product object, never a free-text search.
    - Redirect: requested item id vs the page's canonical link, and vs the
      product's usItemId.
    - Unavailable: product.availabilityStatus == "OUT_OF_STOCK" (FINAL).
    """
    if not os.path.exists(html_file_path):
        print(f"Error: HTML file '{html_file_path}' not found.")
        return None, False, False

    with open(html_file_path, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()

    # -------- REDIRECT DETECTION (canonical) --------
    expected_id = walmart_id_from_url(expected_url) if expected_url else None
    canonical = get_canonical_href(html)
    if expected_id and canonical:
        can_id = walmart_id_from_url(canonical)
        if can_id and can_id != expected_id:
            print(f"[REDIRECT] requested {expected_id} but page is {can_id} -> not available")
            return None, True, False

    # -------- PRICE from __NEXT_DATA__ main product --------
    m = re.search(r'<script[^>]*\bid=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>',
                  html, re.I | re.S)
    if not m:
        print("Price not found: __NEXT_DATA__ missing.")
        return None, False, False
    try:
        product = (json.loads(m.group(1))["props"]["pageProps"]
                   ["initialData"]["data"]["product"]) or {}
    except Exception:
        print("Price not found: product data missing in __NEXT_DATA__.")
        return None, False, False

    # -------- REDIRECT DETECTION (product id) --------
    page_id = str(product.get("usItemId") or "")
    if expected_id and page_id and page_id != expected_id:
        print(f"[REDIRECT] requested {expected_id} but product is {page_id} -> not available")
        return None, True, False

    if str(product.get("availabilityStatus", "")).upper() == "OUT_OF_STOCK":
        print("ℹ️ Walmart product is OUT_OF_STOCK")
        return None, False, True

    price = None
    cur = (product.get("priceInfo") or {}).get("currentPrice") or {}
    if cur.get("price") not in (None, ""):
        price = str(cur["price"])
    elif cur.get("priceString"):
        price = cur["priceString"]

    print(f"🔎 Walmart extracted price: {price}")
    return price, False, False


async def _save_htmls_fresh_chrome(
    site,
    label,
    urls,
    parse_fn,
    retry_backoff_sec,
    between_url_gap_sec,
    output_dir="outputs",
):
    """Browser loop (used by Walmart): for each url (and each retry) launch a FRESH
    real Chrome (new process + new throwaway profile = fresh session), load the
    page Amazon-style, save the HTML, parse it, and tear everything down.
    Returns one result dict per url, in url order (empty slots included)."""
    os.makedirs(output_dir, exist_ok=True)

    async with async_playwright() as p:
        results = []
        for idx, url in enumerate(urls, start=1):
            # empty slot (product not listed on this site): keep the position so
            # results stay aligned with the product groups, but skip cleanly.
            if not url or not url.strip():
                print(f"\n[{label} {idx}/{len(urls)}] empty URL slot -> skipping")
                results.append({"url": url, "file": None, "price": None, "model": None, "status": "empty"})
                continue

            safe_name = sanitize_filename(url)[:120]
            output_file = os.path.join(output_dir, f"{site}_{idx}_{safe_name}.html")
            result = None
            for attempt in range(1, MAX_ATTEMPTS + 1):
                browser = None
                proc = None
                profile_dir = None
                page = None
                try:
                    # FRESH session: new Chrome process + new empty profile.
                    browser, proc, profile_dir = await _launch_chrome_cdp(p)
                    context = (browser.contexts[0] if browser.contexts
                               else await browser.new_context())
                    page = await context.new_page()

                    print(f"\n[{label} {idx}/{len(urls)}] (attempt {attempt}/{MAX_ATTEMPTS}) Navigating to {url} ...")
                    try:
                        # 30s hard cap so a slow page can't stall the whole run.
                        await page.goto(url, wait_until="load", timeout=30000)
                    except TimeoutError:
                        print(f"⚠️ navigation timeout for {url} after 30s. Continuing anyway...")
                        try:
                            await page.wait_for_load_state("domcontentloaded", timeout=7000)
                        except TimeoutError:
                            pass
                    await asyncio.sleep(10)  # Extra wait to ensure dynamic content loads

                    # Wait randomly for page content to settle
                    await human_delay(3, 6)

                    # 🖱️ Simulate random human-like mouse movement
                    for _ in range(3):
                        x = random.randint(200, 800)
                        y = random.randint(200, 600)
                        await page.mouse.move(x, y, steps=random.randint(5, 15))
                        await human_delay(0.3, 1.5)

                    # 🖱️ Random scrolling
                    for _ in range(2):
                        scroll_y = random.randint(400, 1000)
                        await page.mouse.wheel(0, scroll_y)
                        await human_delay(1, 3)

                    # Extract HTML (resilient to any mid-load client-side navigation)
                    html_content = await get_page_content_safe(page)
                    with open(output_file, "w", encoding="utf-8") as f:
                        f.write(html_content)
                    print(f"✅ HTML saved to {output_file}")

                    if _looks_blocked(html_content, site):
                        # bot-check page, not the product -> transient -> retry
                        print(f"⚠️ {label} served a bot-check page ('{_page_title(html_content)}') (attempt {attempt}/{MAX_ATTEMPTS})")
                        result = {"url": url, "file": output_file, "price": None, "model": None, "status": "blocked"}
                    else:
                        price, redirected, unavailable = parse_fn(output_file, expected_url=url)

                        if redirected:
                            result = {"url": url, "file": output_file, "price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "redirect"}
                            break  # final: different product
                        if price:
                            result = {"url": url, "file": output_file, "price": price, "model": None, "status": "ok"}
                            break  # final: got a price
                        result = {"url": url, "file": output_file, "price": NOT_AVAILABLE, "model": NOT_AVAILABLE, "status": "no_price"}
                        if unavailable:
                            print("ℹ️ page marked out of stock / unavailable -> final, not retrying")
                            break
                        print(f"⚠️ price not found & page not marked unavailable (attempt {attempt}/{MAX_ATTEMPTS})")
                except Exception as e:
                    print(f"❌ Error processing URL {url} (attempt {attempt}/{MAX_ATTEMPTS}): {e}")
                    result = {"url": url, "file": None, "price": None, "model": None, "status": f"error: {e}"}
                finally:
                    # Tear down EVERYTHING so the next attempt/url starts with a
                    # brand-new Chrome process + profile (fresh session).
                    for closable in (page, browser):
                        if closable:
                            try:
                                await closable.close()
                            except Exception:
                                pass
                    # browser.close() only disconnects CDP; kill the real Chrome
                    # process and delete its throwaway profile (cookies discarded).
                    if proc:
                        try:
                            proc.kill()
                            proc.wait(timeout=10)
                        except Exception:
                            pass
                    if profile_dir:
                        shutil.rmtree(profile_dir, ignore_errors=True)
                # reached only when the attempt was transient (no break)
                if attempt < MAX_ATTEMPTS:
                    print(f"🔁 retrying in {retry_backoff_sec}s ...")
                    await asyncio.sleep(retry_backoff_sec)

            results.append(result)

            # Politeness gap between consecutive URLs (skipped after the last one).
            if idx < len(urls):
                print(f"⏳ waiting {between_url_gap_sec}s before the next {label} URL ...")
                await asyncio.sleep(between_url_gap_sec)

    return results


async def save_walmart_htmls(urls, output_dir="outputs"):
    return await _save_htmls_fresh_chrome(
        "walmart", "Walmart", urls, parse_walmart_html,
        WALMART_RETRY_BACKOFF_SEC, WALMART_BETWEEN_URL_GAP_SEC, output_dir)

# -----------------------
# Combined main
# -----------------------
async def main():
    # Capture the run's START time in EST. We snap THIS (not the finish time) to
    # the nearest scheduled 6-hour mark, so the Timestamp column reflects the
    # slot the run was launched for and is unaffected by how long scraping takes.
    run_start_est = datetime.datetime.now(datetime.timezone.utc).astimezone(
        ZoneInfo("US/Eastern"))

    # Replace/extend these lists with the product URLs you want to iterate over.
    # Every list has one entry per product slot, in the SAME order as
    # PRODUCT_LABELS. Use "" for a slot where that site doesn't list the product.
    amazon_urls = [
      # (UN43M70HAFXZA)
      "",
  
      # (QN65S90FAFXZA)
      "",
  
      # Samsung 77" 4K TV with a 2 Year Amber Protection Plan (2025) (QN77S90FAFXZA)
      "https://www.amazon.com/dp/B0F6QLH41D?plpRedirect=mhFallback",
  
      # Samsung 43 Inch QLED Q7F 4K Vision AI Smart TV with a HW-LS60D 2.0 Channel Photo Frame Speaker with Dolby Atmos (2025) (QN43Q7FAAFXZA)
      "https://www.amazon.com/dp/B0H2FFWYDG?lv=shuf&channelId=500&plpRedirect=mhFallback",
  
      # (QN65LS03FAFXZA)
      ""
    ]

    bestbuy_urls = [
   # Samsung - 43" Class M70H Series Mini LED 4K UHD SamsungVision AI Smart Tizen TV (2026) (UN43M70HAFXZA)
      "https://www.bestbuy.com/product/samsung-43-class-m70h-series-mini-led-4k-uhd-samsungvision-ai-smart-tizen-tv-2026/JJGRF3QGSH",
  
      # Samsung - 65” Class S90F Series OLED 4K UHD SamsungVision AI Smart Tizen TV (2025) (QN65S90FAFXZA)
      "https://www.bestbuy.com/product/samsung-65-class-s90f-series-oled-4k-uhd-samsungvision-ai-smart-tizen-tv-2025/J3ZYG2V6WP",
  
      # Samsung 4K OLED TV with MI-310L Full Motion Wall Mount (2025) (QN77S90FAFXZA)
      "https://www.bestbuy.com/product/samsung-qn77s90fafxza-4k-oled-tv-with-mi-310l-full-motion-wall-mount-2025/J39TXHR95H",
  
      # Samsung 43 Inch QLED Q7F 4K Vision AI Smart TV with a 2 Year Amber Protection Plan (2025) (QN43Q7FAAFXZA)
      "https://www.bestbuy.com/product/samsung-qn43q7faafxza-43-inch-qled-q7f-4k-vision-ai-smart-tv-with-a-2-year-amber-protection-plan-2025/J39TXHRJK3",
  
      # Samsung - 65” Class LS03FA The Frame Series QLED 4K UHD Art Mode SamsungVision AI Smart TV w/ Wall Mount (2025) (QN65LS03FAFXZA)
      "https://www.bestbuy.com/product/samsung-65-class-ls03fa-the-frame-series-qled-4k-uhd-art-mode-samsungvision-ai-smart-tv-w-wall-mount-2025/J3ZYGXCV3T"
  
    ]

    samsung_urls = [
     # 43 Inch Class Mini LED M70H 4K | Samsung TV | (2026) (UN43M70HAFXZA)
      "https://www.samsung.com/us/tvs/mini-led-tv/43-inch-mini-led-4k-tv-m70h-sku-un43m70hafxza/",
  
      # 65 Inch Class OLED S90F 4K | Samsung Vision AI Smart TV | (2025) (QN65S90FAFXZA)
      "https://www.samsung.com/us/tvs/oled-tv/65-class-oled-tvs90f-sku-qn65s90fafxza/",
  
      # 77 Inch Class OLED S90F 4K | Samsung Vision AI Smart TV | (2025) (QN77S90FAFXZA)
      "https://www.samsung.com/us/tvs/oled-tv/77-class-oled-tvs90f-sku-qn77s90fafxza/",
  
      # 43 Inch Class QLED Q7F 4K | Samsung Vision AI Smart TV | (2025) (QN43Q7FAAFXZA)
      "https://www.samsung.com/us/tvs/qled-tv/43-class-qled-tv-q7f-sku-qn43q7faafxza/",
  
      # 65 Inch Class The Frame QLED 4K | Art Mode Samsung Vision AI Smart TV | (2025) (QN65LS03FAFXZA)
      "https://www.samsung.com/us/lifestyle-tvs/the-frame/65-inch-the-frame-qled-4k-tv-ls03fa-sku-qn65ls03fafxza/"
  
    ]

    lowes_urls = [
      # Samsung M70H 43-in 2160p (4K) Smart LED Indoor Flat screen HDTV (UN43M70HAFXZA)
        "https://www.lowes.com/pd/Samsung-M70H-43-in-2160p-4K-Smart-LED-Indoor-use-only-Flat-screen-HDTV/5021390329",
    
        # Samsung S90F 65-in 2160p (4K) Smart OLED Indoor Flat screen HDTV (QN65S90FAFXZA)
        "https://www.lowes.com/pd/Samsung-Samsung-65in-Class-OLED-4K-QN65S90FAFXZA-Vision-AI-Smart-TV-2025/5016711287",
    
        # Samsung S90F 77-in 2160p (4K) Smart OLED Indoor Flat screen HDTV (QN77S90FAFXZA)
        "https://www.lowes.com/pd/Samsung-Samsung-77in-Class-OLED-4K-QN77S90FAFXZA-Vision-AI-Smart-TV-2025/5016711225",
    
        # Samsung Q7F 43-in 2160p (4K) Smart QLED Indoor Flat screen HDTV (QN43Q7FAAFXZA)
        "https://www.lowes.com/pd/Samsung-Q7F-43-in-2160p-4K-Smart-QLED-Indoor-use-only-Flat-screen-HDTV/5016711209",
    
        # Samsung LS03FA 65-in 2160p (4K) Smart QLED Indoor Flat screen HDTV (QN65LS03FAFXZA)
        "https://www.lowes.com/pd/Samsung-Samsung-65in-Class-The-Frame-LS03FA-QLED-4K-Art-Mode-Vision-AI-Smart-TV-2025/5016711233"
    
    ]

    walmart_urls = [
      # Samsung M70H 43-in Mini LED 4K TV (UN43M70HAFXZA)
        "https://www.walmart.com/ip/Samsung-M70H-43-in-Mini-LED-4K-TV-UN43M70HAFXZA-13-Month-Axiom-Care/20798906398",
    
        #Samsung 4K OLED TV with MI-310L Full Motion Wall Mount (2025) (QN65S90FAFXZA)
        "https://www.walmart.com/ip/Samsung-QN65S90FAFXZA-4K-OLED-TV-with-MI-310L-Full-Motion-Wall-Mount-2025/20967854603",
    
        #Samsung 4K OLED TV with MI-310L Full Motion Wall Mount (2025) (QN77S90FAFXZA)
        "https://www.walmart.com/ip/Samsung-QN65S90FAFXZA-4K-OLED-TV-with-MI-310L-Full-Motion-Wall-Mount-2025/20955064977",
    
        # Samsung 43 Inch QLED Q7F 4K Vision AI Smart TV with a Samsung HW-LS60D 2.0 Channel Photo Frame Speaker with Dolby Atmos (2025) (QN43Q7FAAFXZA)
        "https://www.walmart.com/ip/Samsung-QN43Q7FAAFXZA-43-Inch-QLED-Q7F-4K-Vision-AI-Smart-TV-with-a-Samsung-HW-LS60D-2-0-Channel-Photo-Frame-Speaker-with-Dolby-Atmos-2025/20250800763",
    
        # (QN65LS03FAFXZA)
        "https://www.walmart.com/ip/Samsung-QN65LS03FAFXZA-65-inch-Frame-LS03FA-QLED-4K-Art-Mode-Vision-AI-Smart-TV-2025-Bundle-6FT-HDMI-2-0-Cable-CPS-Exclusive-26-Month-Protection-Pack/20893253329"
    
    ]

    # Amazon, BestBuy and Lowes via Apify (one shared token pool, so a token
    # that runs out during Amazon is not tried again for BestBuy / Lowes)
    apify_pool = ApifyTokenPool.from_env()

    print("\n=== Running Amazon (Apify) ===")
    am_res = fetch_amazon_via_apify(amazon_urls, apify_pool, output_dir="outputs")
    print("\nAmazon Summary:")
    for r in am_res:
        print(r)

    print("\n=== Running BestBuy (Apify) ===")
    bb_res = fetch_bestbuy_via_apify(bestbuy_urls, apify_pool, output_dir="outputs")
    print("\nBestBuy Summary:")
    for r in bb_res:
        print(r)

    print("\n=== Running Samsung scraper ===")
    sam_res = await save_samsung_htmls(samsung_urls, output_dir="outputs", cookies_file="samsung_cookies.json", headless=True)
    print("\nSamsung Summary:")
    for r in sam_res:
        print(r)

    print("\n=== Running Lowes (Apify) ===")
    lo_res = fetch_lowes_via_apify(lowes_urls, apify_pool, output_dir="outputs")
    print("\nLowes Summary:")
    for r in lo_res:
        print(r)

    print("\n=== Running Walmart scraper ===")
    wm_res = await save_walmart_htmls(walmart_urls, output_dir="outputs")
    print("\nWalmart Summary:")
    for r in wm_res:
        print(r)

    # -----------------------
    # Write results (one row per run; prices, SKU columns and 'vs' formulas).
    # am_res / bb_res / sam_res / lo_res / wm_res are in URL order, i.e. slot
    # order, so index s maps directly to product group s.
    # -----------------------
    # The run is scheduled 4x/day to launch on 00:00 / 06:00 / 12:00 / 18:00 EST,
    # but cron / startup jitter means run_start_est is a few minutes off. Snap the
    # START time (captured at the top of main(), NOT this finish time) to the
    # nearest 6-hour mark so the Timestamp column always shows one of the four
    # exact scheduled times, independent of how long scraping took.
    # Rounding to the nearest multiple of 6h naturally yields 0/6/12/18, and rolls
    # over to the next day's 00:00 when the run launches just before midnight.
    _mins = run_start_est.hour * 60 + run_start_est.minute + run_start_est.second / 60
    _snapped = round(_mins / 360) * 360          # nearest 6h (360 min) boundary
    est_snapped = run_start_est.replace(hour=0, minute=0, second=0, microsecond=0) \
        + datetime.timedelta(minutes=_snapped)   # +1440 rolls into the next day
    ts_str = est_snapped.strftime("%d %b %Y, %H:%M")   # e.g. "05 Dec 2025, 06:00"

    # Separate workbook from the old script's results.xlsx: the column layout
    # and products differ, so appending to the old file would misalign rows.
    excel_file = os.path.join("outputs", "results_lowes_walmart.xlsx")
    save_results_wip_format(am_res, bb_res, sam_res, lo_res, wm_res,
                            samsung_urls, ts_str, excel_file)

if __name__ == "__main__":
    asyncio.run(main())
    file_path = "outputs/results.xlsx"

    # column_references = [
    #     "a","d","ar","x","e","as","y","blank column",
    #     "i","aw","ac","j","ax","ad","blank column",
    #     "n","bb","ah","o","bc","ai","blank column",
    #     "s","bg","am","t","bh","an"
    # ]

    # created = copy_columns_by_references(
    #     file_path=file_path,
    #     column_refs=column_references,
    #     source_sheet_name=None,      # None => use first sheet; or set "Sheet1"
    #     new_sheet_base_name="SelectedColumns"
    # )
    # print(f"Created sheet: {created} in {file_path}")
