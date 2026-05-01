"""
app.py
======
Flask backend for the IATA Fuel Tables downloader.

Routes
------
  GET  /          → renders the UI (index.html)
  POST /download  → scrapes tables, builds/updates workbook, streams file back
  GET  /status    → JSON: sheet count + last update time (for live UI refresh)
"""

import traceback
from pathlib import Path
from flask import Flask, render_template, send_file, jsonify

from scrape import scrape_tables
from excel  import build_or_update, WORKBOOK_PATH

app = Flask(__name__)


@app.route("/")
def index():
    sheets     = _sheet_info()
    return render_template("index.html",
                           sheet_count=sheets["count"],
                           last_updated=sheets["last_updated"])


@app.route("/download", methods=["POST"])
def download():
    try:
        # 1. Scrape
        t1_rows, t2_rows = scrape_tables()

        # 2. Build / append sheet
        wb_path, is_dup, sheet_name = build_or_update(t1_rows, t2_rows)

        # 3. Stream the file back with appropriate headers
        response = send_file(
            wb_path,
            as_attachment=True,
            download_name="IATA_Fuel_Tables.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Expose duplicate flag + sheet name to the front-end via headers
        response.headers["X-Is-Duplicate"] = "1" if is_dup else "0"
        response.headers["X-Sheet-Name"]   = sheet_name
        response.headers["Access-Control-Expose-Headers"] = (
            "X-Is-Duplicate, X-Sheet-Name"
        )
        return response

    except Exception as exc:
        tb = traceback.format_exc()
        app.logger.error(f"Download failed:\n{tb}")
        return jsonify({"error": str(exc), "detail": tb}), 500


@app.route("/status")
def status():
    return jsonify(_sheet_info())


# ── Helper ─────────────────────────────────────────────────────────────────────
def _sheet_info() -> dict:
    if not WORKBOOK_PATH.exists():
        return {"count": 0, "last_updated": None, "sheets": []}
    from openpyxl import load_workbook
    try:
        wb     = load_workbook(WORKBOOK_PATH, read_only=True)
        names  = [s for s in wb.sheetnames]   # filter hidden if needed
        wb.close()
        return {
            "count":        len(names),
            "last_updated": names[-1] if names else None,
            "sheets":       names,
        }
    except Exception:
        return {"count": 0, "last_updated": None, "sheets": []}


if __name__ == "__main__":
    # Ensure workbooks dir exists at startup
    WORKBOOK_PATH.parent.mkdir(parents=True, exist_ok=True)
    app.run(debug=True, port=5050)