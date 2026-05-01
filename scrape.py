"""
scrape.py
=========
Downloads the two IATA table images via Selenium (browser-session fetch to
bypass 403), then uses Tesseract OCR with PRECISE column-boundary mapping
derived from actual pixel analysis of the table images.

Column boundaries (at native 1474px width for Table 1, 1025px for Table 2)
are calibrated from real pixel sampling — not guessed from histograms.
"""

import time
import logging
import shutil
import platform
import os

import cv2
import numpy as np
import pytesseract

# ── Tesseract auto-detection (Windows) ────────────────────────────────────────
def _setup_tesseract():
    if shutil.which("tesseract"):
        return
    if platform.system() == "Windows":
        for p in [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        ]:
            if os.path.isfile(p):
                pytesseract.pytesseract.tesseract_cmd = p
                return
        raise RuntimeError(
            "Tesseract not found. Install from "
            "https://github.com/UB-Mannheim/tesseract/wiki"
        )

_setup_tesseract()
log = logging.getLogger(__name__)

# ── IATA image URLs (confirmed from page source) ──────────────────────────────
IATA_URL   = "https://www.iata.org/en/publications/economics/fuel-monitor/"
BASE_URL   = "https://www.iata.org/contentassets/9036deaf9c984009a3515fd6aa1c5e24/"
TABLE1_URL = BASE_URL + "fuel_price_analysis.jpg"
TABLE2_URL = BASE_URL + "jet_fuel_price_devt_recent.jpg"

# ── Table 1: 9 columns, calibrated at 1474px native width ────────────────────
# Boundaries derived from pixel analysis of actual IATA image.
# Format: list of (col_name, x_start_fraction, x_end_fraction)
# Fractions are of total image width so they scale with any resolution.
T1_COL_BOUNDS = [
    ("region",     0.000, 0.197),   # 0 – 290
    ("share",      0.197, 0.299),   # 290 – 440
    ("cts_gal",    0.299, 0.414),   # 440 – 610
    ("bbl",        0.414, 0.516),   # 610 – 760
    ("t",          0.516, 0.617),   # 760 – 910
    ("index_val",  0.617, 0.732),   # 910 – 1079  (0.732 not 0.740: keeps crack's 4.6% in vs_week)
    ("vs_week",    0.732, 0.821),   # 1079 – 1210
    ("vs_month",   0.821, 0.902),   # 1210 – 1330
    ("vs_year",    0.902, 1.000),   # 1330 – 1474
]

# ── Table 2: 5 columns, calibrated at 1025px native width ────────────────────
T2_COL_BOUNDS = [
    ("week_ending",  0.000, 0.280),  # 0 – 287
    ("index_val",    0.280, 0.520),  # 287 – 533
    ("price_bbl",    0.520, 0.690),  # 533 – 707
    ("change",       0.690, 0.860),  # 707 – 881
    ("crack",        0.860, 1.000),  # 881 – 1025
]

# Data rows start below this y-fraction (skip header rows)
T1_DATA_Y_START = 0.185   # below all header rows (~y=113 of 613)
T2_DATA_Y_START = 0.330   # below header (~y=93 of 283)

# Rows to skip entirely (footnote text at bottom)
T1_FOOTNOTE_Y   = 0.880   # y > this fraction = footnote, skip


def _get_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-sync")
    opts.add_argument("--no-first-run")
    opts.add_argument("--mute-audio")
    opts.add_experimental_option("prefs", {
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.media_stream":  2,
    })
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
    )
    # Locate Chrome and Chromedriver — prefer system binaries (Railway/Render)
    # over webdriver-manager downloads. webdriver-manager gets an old version
    # that doesn't match the system Chrome, causing status code 127.

    # 1. Chrome / Chromium binary
    chrome_bin = os.environ.get("CHROME_BIN")
    if not chrome_bin:
        for candidate in ("chromium", "chromium-browser",
                          "google-chrome", "google-chrome-stable"):
            found = shutil.which(candidate)
            if found:
                chrome_bin = found
                log.info(f"Found Chrome at: {found}")
                break

    if chrome_bin:
        opts.binary_location = chrome_bin
    else:
        log.warning("No system Chrome found — falling back to webdriver-manager")

    # 2. Chromedriver — must match the Chrome binary above
    chromedriver_bin = os.environ.get("CHROMEDRIVER_BIN")
    if not chromedriver_bin:
        chromedriver_bin = shutil.which("chromedriver")
        if chromedriver_bin:
            log.info(f"Found chromedriver at: {chromedriver_bin}")

    if chromedriver_bin:
        svc = Service(chromedriver_bin)
    else:
        log.warning("No system chromedriver — using webdriver-manager (may version-mismatch)")
        svc = Service(ChromeDriverManager().install())

    return webdriver.Chrome(service=svc, options=opts)


def _fetch_image(driver, url: str) -> bytes:
    result = driver.execute_async_script(
        """
        var callback = arguments[arguments.length - 1];
        fetch(arguments[0], {headers: {"Referer": "https://www.iata.org/"}})
            .then(r => r.ok ? r.arrayBuffer() : Promise.reject("HTTP " + r.status))
            .then(buf => callback({ok: true,  data: Array.from(new Uint8Array(buf))}))
            .catch(e  => callback({ok: false, err:  String(e)}));
        """,
        url
    )
    if not result or not result.get("ok"):
        raise RuntimeError(f"Could not fetch {url}: {result.get('err', 'unknown')}")
    return bytes(result["data"])


def _ocr_words(img_bytes: bytes) -> tuple[list[tuple], int, int]:
    """
    Returns (words, img_w, img_h) where words = list of (x, y, text).
    x, y are at native image scale (no upscaling applied to coords).
    """
    arr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    if img is None:
        raise RuntimeError("Could not decode image")

    h, w = img.shape[:2]

    # Upscale 2× for better OCR accuracy
    scale  = 2
    up     = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray   = cv2.cvtColor(up, cv2.COLOR_BGR2GRAY)
    _, thr = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)

    tsv = pytesseract.image_to_data(
        thr,
        output_type=pytesseract.Output.DICT,
        config="--psm 6"
    )

    words = [
        # Divide coords back to native scale
        (tsv["left"][i]  // scale,
         tsv["top"][i]   // scale,
         tsv["text"][i].strip())
        for i in range(len(tsv["text"]))
        if tsv["text"][i].strip() and int(tsv["conf"][i]) > 20
    ]
    return words, w, h


def _assign_column(x: int, img_w: int, col_bounds: list) -> int:
    """Return 0-based column index for a word at pixel position x."""
    frac = x / img_w
    for ci, (_, lo, hi) in enumerate(col_bounds):
        if lo <= frac < hi:
            return ci
    # clamp to last column if slightly past edge
    return len(col_bounds) - 1


def _parse_table(img_bytes: bytes,
                 col_bounds: list,
                 data_y_start_frac: float,
                 footnote_y_frac: float = 1.0) -> list[list[str]]:
    """
    Parse an image into rows x columns using calibrated x-boundaries.

    Returns a list of row-lists (strings), data rows only (headers excluded).
    """
    words, img_w, img_h = _ocr_words(img_bytes)

    data_y_min = int(img_h * data_y_start_frac)
    note_y_max = int(img_h * footnote_y_frac)

    # Keep only data-area words
    data_words = [
        (x, y, text)
        for (x, y, text) in words
        if data_y_min <= y < note_y_max
    ]

    if not data_words:
        return []

    # Group into rows by y-coordinate (±10 px tolerance)
    data_words_sorted = sorted(data_words, key=lambda w: w[1])
    raw_rows: list[list[tuple]] = []
    cur_row  = [data_words_sorted[0]]
    cur_y    = data_words_sorted[0][1]

    for word in data_words_sorted[1:]:
        if abs(word[1] - cur_y) <= 10:
            cur_row.append(word)
        else:
            raw_rows.append(sorted(cur_row, key=lambda w: w[0]))
            cur_row = [word]
            cur_y   = word[1]
    raw_rows.append(sorted(cur_row, key=lambda w: w[0]))

    # Assign each word to its column bucket; merge words in same cell
    n_cols = len(col_bounds)
    result = []
    for row_words in raw_rows:
        cells = [""] * n_cols
        for x, _y, text in row_words:
            ci = _assign_column(x, img_w, col_bounds)
            cells[ci] = (cells[ci] + " " + text).strip() if cells[ci] else text

        # Skip completely empty rows
        if any(c for c in cells):
            result.append(cells)

    return result


def _clean_t1(rows: list[list[str]]) -> list[list[str]]:
    """
    Post-process Table 1:
    - Fix OCR region name errors (Africa, Oil Price, Crack Spread)
    - Merge split 'Crack' + 'Spread' rows into one correctly-columned row
    - Replace tilde with minus (OCR sometimes reads - as ~)
    """
    # Pass 1: fix per-cell values
    for row in rows:
        low = row[0].strip().lower()
        if low.startswith("africa"):
            row[0] = "Africa*"
        elif "oil" in low or low.startswith("o11"):
            if "crack" not in low and "spread" not in low:
                row[0] = "Oil Price (Dated Brent)"
        for ci in range(1, 9):
            row[ci] = row[ci].replace("~", "-").replace("`", "-")

    # Pass 2: merge split Crack Spread rows
    cleaned: list[list[str]] = []
    i = 0
    while i < len(rows):
        row    = rows[i]
        # strip leading punctuation OCR sometimes adds
        region = row[0].strip().lower()
        while region and region[0] in ("'", '"', "`", "‘", "’", "“", "”"):
            region = region[1:]

        if region.startswith("crack"):
            merged = ["Crack Spread"] + [""] * 8
            for ci in range(1, 9):
                if row[ci].strip():
                    merged[ci] = row[ci]
            # absorb companion "Spread" row if present
            if i + 1 < len(rows) and "spread" in rows[i + 1][0].lower():
                for ci in range(1, 9):
                    if rows[i + 1][ci].strip() and not merged[ci]:
                        merged[ci] = rows[i + 1][ci]
                i += 1
            cleaned.append(merged)
        elif region == "spread":
            # orphaned spread row — attach numeric values to last crack row
            if cleaned and cleaned[-1][0] == "Crack Spread":
                for ci in range(1, 9):
                    if row[ci].strip() and not cleaned[-1][ci]:
                        cleaned[-1][ci] = row[ci]
        elif not row[0].strip() and not any(row[1:]):
            pass  # discard empty artefact rows
        else:
            cleaned.append(row)
        i += 1

    return cleaned

def _clean_t2(rows: list[list[str]]) -> list[list[str]]:
    """
    Post-process Table 2: merge split date cells (e.g. "24 Apr" + "2026").
    """
    cleaned = []
    i = 0
    while i < len(rows):
        row = rows[i]
        # Detect if date is split: col0 has "24 Apr" and col1 has "2026"
        # This happens because OCR splits the date differently
        # Actually with our boundary fix, col0 should contain the full date
        # Just clean up any stray artefacts
        if row[0].strip() or any(row[1:]):
            cleaned.append(row)
        i += 1
    return cleaned


def scrape_tables() -> tuple[list[list[str]], list[list[str]]]:
    """
    Returns (table1_rows, table2_rows) — data rows only, no headers.
    Each inner list has exactly the right number of columns:
      table1: 9 columns  [region, share, cts/gal, $/bbl, $/t, index, vs_wk, vs_mo, vs_yr]
      table2: 5 columns  [week_ending, index_val, price_bbl, change, crack]
    """
    driver = _get_driver()
    try:
        driver.get(IATA_URL)
        deadline = time.time() + 30
        while time.time() < deadline:
            time.sleep(5)
            try:
                n = driver.execute_script(
                    "return Array.from(document.querySelectorAll('img'))"
                    ".filter(i=>i.naturalWidth>300).length;"
                )
                if n >= 2:
                    break
            except Exception:
                break

        img1 = _fetch_image(driver, TABLE1_URL)
        img2 = _fetch_image(driver, TABLE2_URL)
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    t1 = _parse_table(img1, T1_COL_BOUNDS,
                      data_y_start_frac=T1_DATA_Y_START,
                      footnote_y_frac=T1_FOOTNOTE_Y)
    t2 = _parse_table(img2, T2_COL_BOUNDS,
                      data_y_start_frac=T2_DATA_Y_START)

    t1 = _clean_t1(t1)
    t2 = _clean_t2(t2)

    if not t1:
        raise RuntimeError("Table 1 (Regional Prices) — no data extracted.")
    if not t2:
        raise RuntimeError("Table 2 (Recent Development) — no data extracted.")

    return t1, t2