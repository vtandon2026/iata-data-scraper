"""
excel.py
========
Builds / updates the Excel workbook to exactly match the IATA snapshot.

Receives clean, column-aligned data from scrape.py:
  table1: 9 cols  [region, share, cts/gal, $/bbl, $/t, index_val,
                   vs_week, vs_month, vs_year]
  table2: 5 cols  [week_ending, index_val, price_bbl, change, crack_spread]

Layout per worksheet
--------------------
  Row 1      : Table 1 super-header  (merged cells)
  Row 2      : Table 1 sub-header    (cts/gal, $/bbl, $/t, prior averages)
  Rows 3+    : Table 1 data
  (blank row)
  Row N+2    : Table 2 header
  Rows N+3+  : Table 2 data

Styling matches the snapshot exactly:
  • Dark navy  #1F3864 → T1 main header, T2 "Week Ending" header
  • Grey       #D9D9D9 → T1 sub-header price cols, T2 price/change/crack headers
  • Royal blue #2E4DA7 → "Index Value" column header (both tables)
  • Pale blue  #C5D1EB → "Weekly Average Price versus" header + those 3 data cols
  • Light blue #DCE6F1 → Index Value data cells (both tables)
  • Bold first data row (Jet fuel price global total)
  • Alternating #F2F2F2 on body rows
"""

import hashlib
import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Colours ───────────────────────────────────────────────────────────────────
C_NAVY      = "1F3864"   # main header bg
C_GREY      = "D9D9D9"   # sub-header / neutral header bg
C_BLUE      = "2E4DA7"   # Index Value header bg
C_PALE_BG   = "C5D1EB"   # "versus" header bg
C_IDX_DATA  = "DCE6F1"   # Index Value data cell bg (light blue tint)
C_ALT_ROW   = "F2F2F2"   # alternating row bg
C_WHITE     = "FFFFFF"
C_BLACK     = "000000"
C_BORDER    = "BFBFBF"

WORKBOOK_PATH = Path("workbooks") / "iata_fuel_tables.xlsx"


# ── Style factories ───────────────────────────────────────────────────────────
def _fill(hex_c: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_c)

def _font(bold=False, color=C_BLACK, size=10) -> Font:
    return Font(bold=bold, color=color, size=size, name="Calibri")

def _align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border() -> Border:
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def _cell(ws, r, c, value="", *,
          bold=False, bg=None, fg=C_BLACK,
          h_align="left", wrap=False):
    """Write a styled cell."""
    cell            = ws.cell(row=r, column=c, value=value)
    cell.font       = _font(bold=bold, color=fg)
    cell.alignment  = _align(h=h_align, wrap=wrap)
    cell.border     = _border()
    if bg:
        cell.fill = _fill(bg)
    return cell

def _hdr(ws, r, c, text, bg, fg=C_WHITE, wrap=True):
    """Header cell: centred, bold, wrapped."""
    _cell(ws, r, c, text, bold=True, bg=bg, fg=fg, h_align="center", wrap=wrap)


# ── Table 1 ───────────────────────────────────────────────────────────────────
def _write_table1(ws, rows: list[list[str]], start_row: int) -> int:
    """
    Write Table 1. Returns the index of the last written row.

    rows: each row has exactly 9 elements:
      [0] region name
      [1] share %
      [2] cts/gal
      [3] $/bbl
      [4] $/t
      [5] index value
      [6] vs prior week
      [7] vs prior month
      [8] vs prior year
    """
    r = start_row

    # ── Row 1: super-header ───────────────────────────────────────────────────
    # Col 1 "Week ending / Region" — merges rows 1+2
    ws.merge_cells(start_row=r, start_column=1, end_row=r+1, end_column=1)
    _hdr(ws, r, 1, "Week ending\n24 Apr 2026\n/ Region", C_NAVY)

    # Col 2 "Share in Global Index" — merges rows 1+2
    ws.merge_cells(start_row=r, start_column=2, end_row=r+1, end_column=2)
    _hdr(ws, r, 2, "Share in\nGlobal Index", C_NAVY)

    # Cols 3-5 "Weekly Average Price" — merges cols 3,4,5 in row 1
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    _hdr(ws, r, 3, "Weekly Average Price", C_GREY, fg=C_BLACK)

    # Col 6 "Index Value" — merges rows 1+2
    ws.merge_cells(start_row=r, start_column=6, end_row=r+1, end_column=6)
    _hdr(ws, r, 6, "Index Value\n(Year 2000 = 100)", C_BLUE)

    # Cols 7-9 "Weekly Average Price versus" — merges cols 7,8,9 in row 1
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
    _hdr(ws, r, 7, "Weekly Average Price versus", C_PALE_BG, fg=C_BLACK)

    # ── Row 2: sub-header ─────────────────────────────────────────────────────
    r += 1
    # Cols 1, 2, 6 are merged from row 1 — write blank borders only
    for c in (1, 2, 6):
        ws.cell(row=r, column=c).border = _border()

    _hdr(ws, r, 3, "cts/gal",              C_GREY, fg=C_BLACK)
    _hdr(ws, r, 4, "$/bbl",                C_GREY, fg=C_BLACK)
    _hdr(ws, r, 5, "$/t",                  C_GREY, fg=C_BLACK)
    _hdr(ws, r, 7, "prior week's\naverage",  C_PALE_BG, fg=C_BLACK)
    _hdr(ws, r, 8, "prior month's\naverage", C_PALE_BG, fg=C_BLACK)
    _hdr(ws, r, 9, "prior year's\naverage",  C_PALE_BG, fg=C_BLACK)

    # ── Data rows ─────────────────────────────────────────────────────────────
    for di, row in enumerate(rows):
        r += 1
        padded = (row + [""] * 9)[:9]

        region    = padded[0]
        is_total  = "jet fuel" in region.lower() or di == 0
        is_oil    = "oil" in region.lower() or "brent" in region.lower()
        is_crack  = "crack" in region.lower()
        is_special = is_oil or is_crack
        alt_bg    = C_ALT_ROW if (di % 2 == 1 and not is_total and not is_special) else None

        for ci_zero, val in enumerate(padded):
            ci = ci_zero + 1   # 1-based column
            bold    = is_total or is_special
            h_align = "left" if ci <= 2 else "right"
            # Column-specific background
            if ci == 6:
                bg = C_IDX_DATA
            elif ci >= 7:
                bg = C_PALE_BG if not alt_bg else C_PALE_BG  # always pale blue
            else:
                bg = alt_bg

            _cell(ws, r, ci, val, bold=bold, bg=bg, h_align=h_align)

    return r


# ── Table 2 ───────────────────────────────────────────────────────────────────
def _write_table2(ws, rows: list[list[str]], start_row: int) -> int:
    """
    Write Table 2. Returns last written row index.

    rows: each row has exactly 5 elements:
      [0] week ending date
      [1] index value
      [2] weekly avg price $/bbl
      [3] change vs prior week
      [4] weekly avg crack spread $/bbl
    """
    r = start_row

    # ── Header ────────────────────────────────────────────────────────────────
    _hdr(ws, r, 1, "Week ending",                    C_NAVY)
    _hdr(ws, r, 2, "Index Value\n(Year 2000 = 100)", C_BLUE)
    _hdr(ws, r, 3, "Weekly Average Price\n$/bbl",    C_GREY, fg=C_BLACK)
    _hdr(ws, r, 4, "Change vs\nprior week",           C_GREY, fg=C_BLACK)
    _hdr(ws, r, 5, "Weekly Average\nCrack Spread $/bbl", C_GREY, fg=C_BLACK)

    # ── Data rows ─────────────────────────────────────────────────────────────
    for di, row in enumerate(rows):
        r += 1
        padded  = (row + [""] * 5)[:5]
        alt_bg  = C_ALT_ROW if di % 2 == 1 else None

        for ci_zero, val in enumerate(padded):
            ci      = ci_zero + 1
            h_align = "left" if ci == 1 else "right"
            bg      = C_IDX_DATA if ci == 2 else alt_bg
            _cell(ws, r, ci, val, bg=bg, h_align=h_align)

    return r


# ── Column / row sizing ───────────────────────────────────────────────────────
def _size_sheet(ws):
    # Table 1 uses cols 1-9, Table 2 also uses cols 1-5
    widths = {1: 26, 2: 10, 3: 11, 4: 10, 5: 12, 6: 14, 7: 12, 8: 12, 9: 12}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    # Header rows taller
    ws.row_dimensions[1].height = 38
    ws.row_dimensions[2].height = 28


# ── Duplicate detection ───────────────────────────────────────────────────────
def _data_hash(t1, t2) -> str:
    raw = json.dumps([t1, t2], ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode()).hexdigest()[:16]

def _last_hash(wb) -> str | None:
    if not wb.sheetnames:
        return None
    try:
        return wb[wb.sheetnames[-1]].cell(row=1000, column=1).value
    except Exception:
        return None

def _stash_hash(ws, h: str):
    c = ws.cell(row=1000, column=1, value=h)
    c.font = Font(color="FFFFFF", size=1)


# ── Public API ────────────────────────────────────────────────────────────────
def build_or_update(t1_rows: list[list[str]],
                    t2_rows: list[list[str]]) -> tuple[Path, bool, str]:
    """
    Create workbook (first call) or append a sheet (subsequent calls).

    Returns: (workbook_path, is_duplicate, sheet_name)
    """
    WORKBOOK_PATH.parent.mkdir(parents=True, exist_ok=True)

    current_hash = _data_hash(t1_rows, t2_rows)
    sheet_name   = datetime.now().strftime("%Y-%m-%d %H.%M")

    if WORKBOOK_PATH.exists():
        wb = load_workbook(WORKBOOK_PATH)
        if _last_hash(wb) == current_hash:
            last_sheet = wb.sheetnames[-1] if wb.sheetnames else sheet_name
            wb.close()
            # Return the existing sheet name (not current time) so the UI can
            # show "data unchanged since <last_sheet>" with the actual timestamp
            return WORKBOOK_PATH, True, last_sheet
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    ws = wb.create_sheet(title=sheet_name)

    # ── Table 1 ───────────────────────────────────────────────────────────────
    last_t1 = _write_table1(ws, t1_rows, start_row=1)

    # ── Blank spacer ─────────────────────────────────────────────────────────
    spacer = last_t1 + 1
    ws.row_dimensions[spacer].height = 8

    # ── Table 2 ───────────────────────────────────────────────────────────────
    _write_table2(ws, t2_rows, start_row=spacer + 1)

    # ── Formatting ────────────────────────────────────────────────────────────
    _size_sheet(ws)
    # ws.freeze_panes = "A3"  # disabled — no frozen rows

    # ── Hash for duplicate detection ─────────────────────────────────────────
    _stash_hash(ws, current_hash)

    wb.save(WORKBOOK_PATH)
    wb.close()

    return WORKBOOK_PATH, False, sheet_name