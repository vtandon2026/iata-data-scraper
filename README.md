# IATA Fuel Price Monitor — Table Exporter

A one-click Flask web app that scrapes the two weekly data tables from the
[IATA Jet Fuel Price Monitor](https://www.iata.org/en/publications/economics/fuel-monitor/)
and downloads them as a formatted Excel workbook.

---

## How it works

### Why not `requests` + `pandas.read_html`?
The two tables are **static JPG images** hosted at fixed URLs — not HTML table
elements.  They cannot be parsed by BeautifulSoup or pandas directly.

### Scraping strategy
1. **Selenium** opens the IATA page in a headless Chrome browser.
2. The browser's `fetch()` API is used to download each image through the
   authenticated session (direct HTTP downloads return 403).
3. **Tesseract OCR** reads the image bytes and extracts word positions.
4. Words are grouped into rows by y-coordinate and bucketed into columns by
   x-position, reconstructing the table structure.

### Table identification
| Image file | Content |
|---|---|
| `fuel_price_analysis.jpg` | Table 1 — Regional Prices & Index Values |
| `jet_fuel_price_devt_recent.jpg` | Table 2 — Recent 5-Week Development |

Both files live under:
`https://www.iata.org/contentassets/9036deaf9c984009a3515fd6aa1c5e24/`

These filenames are confirmed from the IATA page source and have been stable.
If they change, update `TABLE1_URL` / `TABLE2_URL` in `scrape.py`.

### Sheet naming & versioning
- Each click creates a new worksheet named `YYYY-MM-DD HH:MM`
  (e.g. `2026-04-30 16:45`).
- The workbook file (`workbooks/IATA_Fuel_Tables.xlsx`) is **never
  overwritten** — sheets accumulate so you have a historical record.
- **Duplicate detection**: before writing, the app SHA-256-hashes the raw OCR
  data.  If the hash matches the last sheet, no new sheet is added and the
  existing workbook is returned with a notice in the UI.

### Excel layout
```
Rows  1–2  : Table 1 merged header (dark navy + grey sub-header)
Rows  3–N  : Table 1 data (regional prices, oil price, crack spread)
Row   N+1  : blank spacer
Row   N+2  : Table 2 header
Rows  N+3… : Table 2 data (last 5 weeks)
```
Styling mirrors the IATA snapshot exactly:
- Dark navy (`#1F3864`) for primary headers
- Royal blue (`#2E4DA7`) for the Index Value column
- Pale blue (`#C5D1EB`) for the "versus" comparison columns
- First data row bold (global jet fuel total)
- Alternating light grey on body rows
- Borders on all cells; top 2 rows frozen

---

## Setup

### Prerequisites
- **Python 3.11+**
- **Google Chrome** (any recent version)
- **Tesseract OCR**
  - Windows: download from https://github.com/UB-Mannheim/tesseract/wiki
    and install to the default path — the app finds it automatically.
  - Mac:   `brew install tesseract`
  - Linux: `sudo apt install tesseract-ocr`

### Install & run

```bash
# 1. Clone / unzip the project
cd iata_app

# 2. Create a virtual environment
python -m venv .venv

# Windows
.venv\Scripts\activate
# Mac/Linux
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the Flask app
python app.py
```

Open **http://localhost:5050** in your browser.

---

## Configuration

| What | Where | Default |
|---|---|---|
| Target URL | `scrape.py` → `IATA_URL` | IATA Fuel Monitor page |
| Image 1 URL | `scrape.py` → `TABLE1_URL` | `fuel_price_analysis.jpg` |
| Image 2 URL | `scrape.py` → `TABLE2_URL` | `jet_fuel_price_devt_recent.jpg` |
| Workbook output path | `excel.py` → `WORKBOOK_PATH` | `workbooks/IATA_Fuel_Tables.xlsx` |
| Page load wait (seconds) | `scrape.py` → deadline | 30 s |
| Flask port | `app.py` → `app.run(port=...)` | 5050 |

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `Tesseract not found` | Not installed or not in PATH | Install from link above; restart terminal |
| `Could not fetch … HTTP 403` | Browser session not ready | Increase `deadline` in `scrape.py` |
| Browser session crashes mid-wait | Heavy IATA page JS | Already handled by poll loop; retry |
| Tables empty / garbled OCR | Image changed resolution | Download image manually and inspect |
| `openpyxl` merge error | Existing workbook corrupted | Delete `workbooks/IATA_Fuel_Tables.xlsx` |
| Port 5050 in use | Another process | Change `port=5050` in `app.py` |

---

## File structure
```
iata_app/
├── app.py              Flask backend (routes)
├── scrape.py           Selenium + OCR extraction
├── excel.py            openpyxl workbook builder
├── requirements.txt
├── workbooks/          Created automatically; workbook lives here
├── templates/
│   └── index.html      UI
└── static/
    └── style.css       Styling
```