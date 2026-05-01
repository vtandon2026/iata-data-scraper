# IATA Fuel Price Monitor — Table Scraper

A one-click Flask web app that scrapes the two weekly data tables from the
[IATA Jet Fuel Price Monitor](https://www.iata.org/en/publications/economics/fuel-monitor/)
and exports them to a formatted Excel workbook.

---

## How it works

The two tables on the IATA page are rendered as **static JPG images** — not
HTML. Direct HTTP requests return 403. The scraper works around both
constraints:

1. **Selenium** opens the page in a headless Chrome browser
2. The browser's `fetch()` API downloads each image through the authenticated
   session (bypasses the 403)
3. **Tesseract OCR** reads the image bytes and extracts word positions
4. Words are grouped into rows by y-coordinate and bucketed into columns using
   calibrated pixel boundaries derived from the actual images
5. **openpyxl** writes the data into a formatted Excel workbook

---

## Workbook structure

Every download produces / updates a single file:
`workbooks/IATA_Fuel_Tables.xlsx`

| Sheet | Description |
|---|---|
| **Consolidated T1** | Always first. Accumulates Table 1 across every download. New block appended below the previous one with a merged "Extraction Date" column on the left. |
| **Consolidated T2** | Always second. Same accumulation for Table 2. |
| **YYYY-MM-DD HH.MM** | Snapshot of that download — both tables on one sheet. A new sheet is added each time new data is detected. |

### Consolidated sheet layout

```
Col A              │ Col B                   │ Col C … │
───────────────────┼─────────────────────────┼─────────┤
Extraction Date    │ Week ending / Region    │ …       │  ← header
(merged across     ├─────────────────────────┼─────────┤
 all data rows     │ Jet fuel price  100% …  │         │  ← data rows
 of this block)    │ Asia & Oceania   22% …  │         │
                   │ …                       │         │
───────────────────┼─────────────────────────┼─────────┤
  (blank spacer)   │                         │         │
───────────────────┼─────────────────────────┼─────────┤
Extraction Date    │ Week ending / Region    │ …       │  ← next week's header
(next block)       ├─────────────────────────┼─────────┤
                   │ Jet fuel price  100% …  │         │
                   │ …                       │         │
```

### Duplicate detection

Before writing, the app SHA-256 hashes the raw OCR data and compares it to the
hash stored (invisibly) in the last snapshot sheet. If they match — meaning
IATA hasn't published new data yet — no new sheet is added and the consolidated
sheets are left unchanged. The UI shows which sheet the data was last seen in.

---

## Scraped tables

### Table 1 — Fuel Price Analysis (9 columns)

| Column | Description |
|---|---|
| Week ending / Region | Global + 6 regions + Oil Price + Crack Spread |
| Share in Global Index | Regional weight % |
| cts/gal | Weekly average price in cents per gallon |
| $/bbl | Weekly average price in dollars per barrel |
| $/t | Weekly average price in dollars per tonne |
| Index Value (Year 2000 = 100) | Price index |
| vs prior week's average | % change |
| vs prior month's average | % change |
| vs prior year's average | % change |

### Table 2 — Recent 5-Week Development (5 columns)

| Column | Description |
|---|---|
| Week ending | Date of the weekly report |
| Index Value (Year 2000 = 100) | Price index |
| Weekly Average Price $/bbl | Price in dollars per barrel |
| Change vs prior week | % change week-on-week |
| Weekly Average Crack Spread $/bbl | Refining margin |

---

## File structure

```
iata_app/
├── app.py               Flask backend (routes: /, /download, /status, /reset)
├── scrape.py            Selenium + Tesseract OCR → structured table rows
├── excel.py             openpyxl workbook builder (snapshot + consolidated sheets)
├── requirements.txt
├── .gitignore
├── workbooks/           Auto-created at runtime; Excel workbook lives here (gitignored)
├── templates/
│   └── index.html       Single-page UI
└── static/
    └── style.css
```

---

## Setup

### Prerequisites

- **Python 3.11+**
- **Google Chrome** (any recent version — chromedriver is auto-managed)
- **Tesseract OCR**
  - Windows: download installer from https://github.com/UB-Mannheim/tesseract/wiki
    and install to the default path — detected automatically
  - Mac: `brew install tesseract`
  - Linux: `sudo apt install tesseract-ocr`

### Install & run

```bash
# 1. Create and activate a virtual environment
python -m venv .venv

# Windows
.venv\Scripts\activate
# Mac / Linux
source .venv/bin/activate

# 2. Install dependencies
pip install -r requirements.txt

# 3. Start the app
python app.py
```

Open **http://localhost:5050** in your browser and click **Download Excel**.

---

## Configuration

All tuneable values are at the top of each file:

| Setting | File | Default |
|---|---|---|
| IATA page URL | `scrape.py` → `IATA_URL` | IATA Fuel Monitor |
| Table 1 image filename | `scrape.py` → `TABLE1_FILENAME` | `fuel_price_analysis.jpg` |
| Table 2 image filename | `scrape.py` → `TABLE2_FILENAME` | `jet_fuel_price_devt_recent.jpg` |
| Page load wait (seconds) | `scrape.py` → `deadline = time.time() + 30` | 30 s |
| Workbook output path | `excel.py` → `WORKBOOK_PATH` | `workbooks/IATA_Fuel_Tables.xlsx` |
| Flask port | `app.py` → `app.run(port=...)` | 5050 |

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `Tesseract not found` | Install from link above; restart terminal after installing |
| Browser crashes mid-scrape | Increase timeout: `deadline = time.time() + 45` in `scrape.py` |
| Columns misaligned in Excel | IATA may have changed image dimensions — check `T1_COL_BOUNDS` / `T2_COL_BOUNDS` in `scrape.py` |
| Old bad-formatted sheet in workbook | Delete `workbooks/IATA_Fuel_Tables.xlsx` and re-download |
| "Data unchanged" on first clean download | Expected — the hash matched a previous (possibly broken) run. Delete the workbook and try again |
| Port 5050 already in use | Change `port=5050` to any free port in `app.py` |
| `openpyxl` merge error on open | Workbook may be corrupted — delete it and re-download |