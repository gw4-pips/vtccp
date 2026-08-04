"""
Update CPIPM-Project-Outline-v1.0.docx → CPIPM-Project-Outline-v1.2.docx

Changes:
  1. Title paragraph fix: "CP Inline" → "VCCS Command Pilot™ Inline"
  2. First body-copy TM: first prose use of "Command Pilot Inline" → "Command Pilot™ Inline"
  3. General terminology: remaining "CP Inline" → "Command Pilot Inline"
  4. Version string: 1.0 → 1.2, date 2026-08-03 → 4 August 2026
  5. TABLE 1 only — proportional height: calc_visual_rows × 0.25" (360 twips per row)
     All other tables — previous H1/H2/H3 tier formula, unchanged.
  6. Vertical-centre all table cells.
  7. Version history table: append v1.2 row.
"""

import math, os, copy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = os.path.join(os.path.dirname(__file__), "..", "..", "CPIPM-Project-Outline-v1.0.docx")
OUT = os.path.join(os.path.dirname(__file__), "..", "references", "CPIPM-Project-Outline-v1.2.docx")

# ── Height constants for non-T1 tables (twips; 1440 = 1 inch) ────────────────
H1 = 360    # 0.25"  — 1 visual row
H2 = 1080   # 0.75"  — 2 visual rows (3×)
H3 = 1440   # 1.00"  — 3+ visual rows (4×)
TWIPS_PER_ROW = 360  # 0.25" — used for Table 1 proportional formula

# Estimated usable characters per line by number of columns.
CHARS_PER_LINE = {1: 90, 2: 75, 3: 60, 4: 45, 5: 35, 6: 25}
CHARS_DEFAULT  = 20

# ── Text replacements (order matters — most-specific first) ──────────────────
# Note: title "CP Inline" is handled by exact-match in fix_paragraph()
REPLACEMENTS = [
    # Exact header/version line
    ("Version 1.0   |   2026-08-03   |   DRAFT", "Version 1.2   |   4 August 2026   |   DRAFT"),
    # Para [1]: "Command Pilot Inline Production Module" — already correct after below rules
    # Uppercase wireframe
    ("CP INLINE v1.0",  "Command Pilot Inline v1.0"),
    ("CP INLINE",       "Command Pilot Inline"),
    # Author
    ("CP / Engineering", "VCCS/PIPS / Engineering"),
    # General CP Inline → Command Pilot Inline (all remaining forms)
    ("CP Inline is",    "Command Pilot Inline is"),
    ("CP Inline reuses","Command Pilot Inline reuses"),
    ("CP Inline v",     "Command Pilot Inline v"),
    ("CP Inline —",     "Command Pilot Inline —"),
    ("CP Inline,",      "Command Pilot Inline,"),
    ("CP Inline.",      "Command Pilot Inline."),
    ("CP Inline ",      "Command Pilot Inline "),
    ("CP Inline",       "Command Pilot Inline"),
]

def replace_text(text):
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    return text

def fix_paragraph(para):
    """Consolidate runs, apply replacements, restore into first run.
       Special case: if the full text is exactly 'CP Inline' → title heading."""
    full = "".join(r.text for r in para.runs)
    if not para.runs:
        return

    # Title paragraph: exact match → VCCS Command Pilot™ Inline
    if full.strip() == "CP Inline":
        para.runs[0].text = "VCCS Command Pilot\u2122 Inline"
        for r in para.runs[1:]:
            r.text = ""
        return

    new = replace_text(full)
    if new == full:
        return
    para.runs[0].text = new
    for r in para.runs[1:]:
        r.text = ""

# ── Row-height helpers ────────────────────────────────────────────────────────
def cell_visual_rows(cell, cpl):
    texts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    if not texts:
        return 0
    total = 0
    for t in texts:
        total += max(1, math.ceil(len(t) / cpl))
    return total

def max_visual_rows_in_row(table_row, n_cols):
    cpl  = CHARS_PER_LINE.get(n_cols, CHARS_DEFAULT)
    seen = set()
    max_vr = 0
    for cell in table_row.cells:
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue
        seen.add(tc_id)
        vr = cell_visual_rows(cell, cpl)
        if vr > max_vr:
            max_vr = vr
    return max_vr

def set_row_height(row, twips):
    tr   = row._tr
    trPr = tr.get_or_add_trPr()
    for el in trPr.findall(qn("w:trHeight")):
        trPr.remove(el)
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"),   str(twips))
    trH.set(qn("w:hRule"), "atLeast")
    trPr.append(trH)

def set_cell_valign(cell, align="center"):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for el in tcPr.findall(qn("w:vAlign")):
        tcPr.remove(el)
    vA = OxmlElement("w:vAlign")
    vA.set(qn("w:val"), align)
    tcPr.append(vA)

# ── Version-history row append ────────────────────────────────────────────────
def append_version_row(table, version, date, author, changes):
    new_tr = copy.deepcopy(table.rows[-1]._tr)
    table._tbl.append(new_tr)
    new_row = table.rows[-1]
    for cell, val in zip(new_row.cells, [version, date, author, changes]):
        for para in cell.paragraphs:
            for run in para.runs:
                run.text = ""
        if cell.paragraphs:
            p = cell.paragraphs[0]
            if p.runs:
                p.runs[0].text = val
            else:
                p.add_run(val)
        set_cell_valign(cell, "center")

# ── Main ──────────────────────────────────────────────────────────────────────
doc = Document(SRC)

# 1. Body paragraphs — fix text, then find first body TM use
tm_first_done = False
for para_idx, para in enumerate(doc.paragraphs):
    fix_paragraph(para)
    # First prose paragraph (past title block at index ≤ 5) that contains
    # "Command Pilot Inline" gets the ™ on its first occurrence.
    if not tm_first_done and para_idx > 5:
        for run in para.runs:
            if "Command Pilot Inline" in run.text:
                run.text = run.text.replace(
                    "Command Pilot Inline", "Command Pilot\u2122 Inline", 1)
                tm_first_done = True
                break

# 2. Tables
tables = doc.tables
for ti, table in enumerate(tables):
    n_cols   = len(table.columns)
    is_t1    = (ti == 0)
    is_last  = (ti == len(tables) - 1)

    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                fix_paragraph(para)
            set_cell_valign(cell, "center")

        vr = max_visual_rows_in_row(row, n_cols)

        if is_t1:
            # Table 1: proportional — each visual row = 0.25"
            twips = max(1, vr) * TWIPS_PER_ROW
        else:
            # All other tables: H1/H2/H3 tier formula
            if vr <= 1:   twips = H1
            elif vr == 2: twips = H2
            else:         twips = H3

        set_row_height(row, twips)

    if is_last:
        append_version_row(
            table,
            "1.2", "4 August 2026", "VCCS/PIPS / Engineering",
            "Table 1 row heights revised to calc-lines \u00d7 0.25\"; "
            "title TM symbol fixed; first body-copy TM added."
        )

# 3. Save
os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print(f"Saved: {OUT}")

# Quick verification — print Table 1 heights for review
print("\nTable 1 heights (v1.2):")
cpl = CHARS_PER_LINE[2]
for ri, row in enumerate(doc.tables[0].rows):
    vr = max_visual_rows_in_row(row, 2)
    tw = max(1, vr) * TWIPS_PER_ROW
    inch = tw / 1440
    print(f"  r{ri:02d}  vr={vr}  {tw}tw  {inch:.3f}\"")
