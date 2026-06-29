"""
excel_search.py — add an order-independent keyword search sheet to a workbook.
-----------------------------------------------------------------------------
Built for the ground team who work straight off the Excel file (no laptop /
localhost / dashboard). They type keywords in ANY order into one yellow cell
and the matching rows appear instantly below.

How the match works (Excel 365 dynamic arrays):
  * TEXTSPLIT(TRIM(box)," ")  → splits the typed text into keywords
  * SEARCH(keyword, description) is case-insensitive and order-independent
  * a row is a match only when EVERY keyword is found (BYROW + AND)
  * FILTER returns the whole matching rows

So searching "cockpit glass broken" finds "GLASS BROKEN IN COCKPIT".

NOTE on the funny prefixes below (_xlfn / _xlfn._xlws / _xlpm):
openpyxl stores formulas verbatim, but the OOXML file format requires modern
functions to be written with these internal prefixes or Excel shows #NAME?.
Excel displays the clean names (FILTER, TEXTSPLIT, …) to the user.
"""

from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

INPUT_CELL = "B3"          # the yellow "type here" cell
HEADER_ROW = 5             # results header row
DATA_ROW   = 6             # FILTER spills from here down

_THIN = Side(style="thin", color="BBBBBB")


def add_search_sheet(
    book,
    title: str,
    src_sheet: str,
    headers: list[str],
    search_col_name: str,
    n_data_rows: int,
    tab_color: str = "FFC000",
    blurb: str = "",
) -> None:
    """Insert a search sheet (at the front) that searches `src_sheet`.

    book            : openpyxl Workbook
    title           : name for the new search sheet
    src_sheet       : name of the data sheet to search
    headers         : column headers of the data sheet (in order)
    search_col_name : which column to match keywords against
    n_data_rows     : number of data rows (excluding header) in src_sheet
    """
    ncols          = len(headers)
    last_col       = get_column_letter(ncols)
    last_row       = n_data_rows + 1                       # data starts at row 2
    search_idx     = headers.index(search_col_name) + 1
    search_col     = get_column_letter(search_idx)

    ws = book.create_sheet(title, 0)                       # 0 = put it first
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_view.showGridLines = False

    # ── Title + instructions ─────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    t = ws["A1"]
    t.value = "🔍  TASK SEARCH  —  type keywords in ANY order, then press Enter"
    t.font = Font(bold=True, size=14, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="1F4E78")
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    ws.merge_cells(f"A2:{last_col}2")
    sub = ws["A2"]
    sub.value = (
        "Example: typing  cockpit glass broken  also finds  “GLASS BROKEN IN "
        "COCKPIT”.   " + blurb
    ).strip()
    sub.font = Font(italic=True, size=10, color="555555")
    sub.alignment = Alignment(horizontal="left", indent=1)

    # ── Input label + yellow search box ──────────────────────────────────────
    lbl = ws["A3"]
    lbl.value = "Search:"
    lbl.font = Font(bold=True, size=12)
    lbl.alignment = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("B3:F3")
    box = ws[INPUT_CELL]
    box.fill = PatternFill("solid", fgColor="FFF2CC")
    box.font = Font(bold=True, size=12)
    box.alignment = Alignment(vertical="center", indent=1)
    medium = Side(style="medium", color="BF9000")
    for col in range(2, 7):                                # B..F border box
        c = ws.cell(row=3, column=col)
        c.border = Border(left=medium, right=medium, top=medium, bottom=medium)
    ws.row_dimensions[3].height = 24

    # ── Results header row (copied from the data sheet) ───────────────────────
    for j, name in enumerate(headers, start=1):
        c = ws.cell(row=HEADER_ROW, column=j, value=str(name))
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2F5496")
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

    # ── The search formula (spills the matching rows) ─────────────────────────
    src = f"'{src_sheet}'"
    box_ref = f"${INPUT_CELL[0]}${INPUT_CELL[1:]}"          # B3 -> $B$3
    no_match = "No matches — try fewer or different keywords"
    empty_msg = "← Type keywords (any order) in the yellow box above"
    # BYROW and MMULT both rely on SEARCH broadcasting a 1×n_kw keyword array
    # against a m×1 description column into a m×n_kw matrix — this is unreliable
    # in Excel and returns wrong dimensions (only column A shows up in results).
    #
    # REDUCE processes keywords one at a time as scalars:
    #   SEARCH(scalar_kw, m×1_desc) → m×1   (no 2D broadcasting needed)
    #   acc (starts TRUE=1) * m×1 → m×1     (element-wise multiply)
    # Final result: m×1 boolean — TRUE only where ALL keywords were found.
    formula = (
        f'=IF(TRIM({box_ref})="","{empty_msg}",'
        f'_xlfn._xlws.FILTER('
        f'{src}!A2:{last_col}{last_row},'
        f'_xlfn.REDUCE(TRUE,_xlfn.TEXTSPLIT(TRIM({box_ref})," "),'
        f'_xlfn.LAMBDA(_xlpm.a,_xlpm.k,'
        f'_xlpm.a*ISNUMBER(SEARCH(_xlpm.k,'
        f'{src}!{search_col}2:{search_col}{last_row})))),'
        f'"{no_match}"))'
    )
    ws.cell(row=DATA_ROW, column=1, value=formula)

    # ── Cosmetics: column widths + freeze the search area ─────────────────────
    for j in range(1, ncols + 1):
        ws.column_dimensions[get_column_letter(j)].width = 22
    ws.freeze_panes = f"A{DATA_ROW}"
