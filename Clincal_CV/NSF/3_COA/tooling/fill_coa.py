# -*- coding: utf-8 -*-
"""Fill NSF's COA template (PAPPG 24-1 five-table version) from table_data.json.

Layout of the pristine template (one sheet, 'NSF COA Template', no merged cells):
  Table 1  header row 16  data rows 17-21   cols B name | C affiliation | D last active date
  Table 2  header row 27  data rows 28-32   cols A type(R:) | B name | C relationship | D optional | E last active
  Table 3  header row 37  data rows 38-44   cols A type(G:/T:) | B name | C affiliation | D optional
  Table 4  header row 51  data rows 52-56   cols A type(A:/C:) | B name | C affiliation | D optional | E last active
  Table 5  header row 63  data rows 64-68   cols A type(B:/E:) | B name | C affiliation | D journal | E last active
Rows are inserted only inside a table's data area; data validations and row heights below the insertion point
are shifted by hand because openpyxl does not move them.
"""
import copy, json, sys
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SRC, DATA, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
data = json.load(open(DATA, encoding="utf-8"))

wb = load_workbook(SRC)
ws = wb["NSF COA Template"]

# --- locate tables by their header labels (so the script does not depend on hard-coded rows)
def find_header(col, text):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v is not None and str(v).strip() == text:
            return r
    raise KeyError(text)

def table_rows(hdr):
    """Data rows of a table: rows below the header that share its bottom-border style, up to the first row
    whose column B/C cells have no bottom border."""
    rows = []
    r = hdr + 1
    while ws.cell(r, 2).border.bottom.style == "thin" and ws.cell(r, 3).border.bottom.style == "thin":
        rows.append(r)
        r += 1
    return rows

# --- helpers to shift validations / row heights when inserting rows
def shift_ref(ref, at, n):
    c1, r1, c2, r2 = range_boundaries(ref)
    if r1 >= at:
        r1, r2 = r1 + n, r2 + n
    elif r2 >= at:                       # insertion point inside the range -> range grows
        r2 += n
    a = f"{get_column_letter(c1)}{r1}"
    b = f"{get_column_letter(c2)}{r2}"
    return a if a == b else f"{a}:{b}"

def insert_rows_inside(at, n, style_row):
    """Insert n rows before row `at`, shifting validations and row heights, and copying the cell styles of
    `style_row` (a data row of the same table) into the new rows."""
    # row heights below the insertion point
    heights = {}
    for idx in sorted(ws.row_dimensions.keys()):
        if idx >= at:
            rd = ws.row_dimensions[idx]
            heights[idx] = (rd.height, rd.hidden)
    for idx in heights:
        del ws.row_dimensions[idx]
    ws.insert_rows(at, n)
    for idx, (h, hidden) in heights.items():
        rd = ws.row_dimensions[idx + n]
        rd.height, rd.hidden = h, hidden
    # data validations
    for dv in ws.data_validations.dataValidation:
        new = " ".join(shift_ref(str(rng), at, n) for rng in dv.sqref.ranges)
        dv.sqref = new
    # Excel table objects (ListObjects): the table containing the insertion point grows, those below move
    for t in ws.tables.values():
        t.ref = shift_ref(t.ref, at, n)
        if t.autoFilter is not None and t.autoFilter.ref:
            t.autoFilter.ref = shift_ref(t.autoFilter.ref, at, n)
    # styles for the new rows
    src_row = style_row if style_row < at else style_row + n
    for r in range(at, at + n):
        for c in range(1, 6):
            s = ws.cell(src_row, c)
            d = ws.cell(r, c)
            d._style = copy.copy(s._style)

def clear_rows(rows):
    for r in rows:
        for c in range(1, 6):
            ws.cell(r, c).value = None

def write_row(r, values):
    for c, v in enumerate(values, start=1):
        if v not in (None, ""):
            ws.cell(r, c).value = v

# ---------------------------------------------------------------- Tables 1-3 (no insertion needed)
h1 = find_header(2, "Your Name:")
rows1 = table_rows(h1)
clear_rows(rows1)
t1 = data["table1"]
assert len(t1) <= len(rows1)
for r, row in zip(rows1, t1):
    write_row(r, [None, row["name"] or None, row["org"], None])

h2 = find_header(3, "Type of Relationship")
rows2 = table_rows(h2)
clear_rows(rows2)
assert len(data["table2"]) <= len(rows2)
for r, row in zip(rows2, data["table2"]):
    write_row(r, ["R:" if row.get("preclude") else None, row["name"], row["relationship"], row.get("optional"), None])

h3 = find_header(2, "Advisor/Advisee Name:")
rows3 = table_rows(h3)
clear_rows(rows3)
assert len(data["table3"]) <= len(rows3)
for r, row in zip(rows3, data["table3"]):
    write_row(r, [row["type"] + ":", row["name"], row["org"], row["optional"], None])

# ---------------------------------------------------------------- Table 4 (insert rows)
h4 = find_header(3, "Organizational Affiliation")
# there are three 'Organizational Affiliation' headers (tables 3, 4, 5); pick the one whose col E says 'Last Active'
# and whose col A says '4'
for r in range(1, ws.max_row + 1):
    if str(ws.cell(r, 1).value).strip() == "4" and str(ws.cell(r, 3).value).strip() == "Organizational Affiliation":
        h4 = r
        break
rows4 = table_rows(h4)
need = len(data["table4"])
extra = max(0, need - len(rows4))
if extra:
    insert_rows_inside(rows4[1], extra, rows4[0])     # insert after the first data row, inside the validation range
    rows4 = table_rows(h4)
assert len(rows4) >= need, (len(rows4), need)
clear_rows(rows4)
# The template's sample rows 53-55 carry per-cell style quirks (e.g. 'Hyperlink' style on the Optional column
# where the example e-mail sat). Give every data row except the table's last row the style of the first data
# row so the filled table is uniform; the last row keeps its template style (no bottom border, table edge).
for r in rows4[1:]:
    for c in range(1, 6):
        ws.cell(r, c)._style = copy.copy(ws.cell(rows4[0], c)._style)

SHORT = {   # compact flag text for the spreadsheet; full reasons live in COA_source_list.md
    "last known": "[CONFIRM] last known org; current employer unverified",
    "identity": "[CONFIRM] identity ambiguous",
    "move": "[CONFIRM] possible move; verify",
    "listing": "[CONFIRM] not in CV entry; from ASTRO 2026 listing",
    "spelling": "[CONFIRM] CV spelling; likely match",
    "generic": "[CONFIRM] verify affiliation",
}
def short_flag(optional):
    o = optional.lower()
    if "last known" in o or "no newer record" in o or "current employer" in o:
        return SHORT["last known"]
    if "identity" in o or "first name not found" in o or "probably" in o or "could be a different" in o:
        return SHORT["identity"]
    if "not in the cv entry" in o:
        return SHORT["listing"]
    if "cv spelling" in o:
        return SHORT["spelling"]
    if "move" in o or "moffitt" in o or "varian" in o or "georgetown" in o or "bylines" in o or "alumni page" in o:
        return SHORT["move"]
    return SHORT["generic"]

for r, row in zip(rows4, data["table4"]):
    opt = row["optional"]
    if "[CONFIRM]" in opt:
        dept = opt.split("[CONFIRM]")[0].strip(" ;")
        cell = (dept + "; " if dept else "") + short_flag(opt)
    else:
        cell = opt
    write_row(r, ["A:", row["name"], row["org"], cell or None, None])

# ---------------------------------------------------------------- Table 5
h5 = find_header(4, "Journal/Collection")
rows5 = table_rows(h5)
clear_rows(rows5)
assert len(data["table5"]) <= len(rows5)
for r, row in zip(rows5, data["table5"]):
    write_row(r, ["B:", row["name"], row["org"], "Medical Physics", None])

wb.save(OUT)
print("saved", OUT)
print("table rows:", {"t1": rows1, "t2": rows2, "t3": rows3, "t4": (rows4[0], rows4[-1], len(rows4)), "t5": rows5})
print("validations:", [(str(dv.sqref), dv.formula1) for dv in ws.data_validations.dataValidation])
print("tables:", [(t.name, t.ref, t.autoFilter.ref if t.autoFilter else None) for t in ws.tables.values()])
