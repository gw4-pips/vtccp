"""
Update CPIPM-Project-Outline-v1.0.docx → CPIPM-Project-Outline-v1.1.docx

Changes applied:
  1. Proportional row heights — visual rows estimated from text length / column width:
       1 visual row  → 0.25" (360 twips)
       2 visual rows → 0.75" (1080 twips  =  3 × 0.25")
       3+ visual rows→ 1.00" (1440 twips  =  4 × 0.25")
  2. Vertical-centre all table cells.
  3. Terminology: "CP Inline" → "Command Pilot Inline" throughout.
  4. Title version string: "Version 1.0" → "Version 1.1".
  5. Version history table: append v1.1 row.
"""

import math, os
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

SRC = os.path.join(os.path.dirname(__file__), "..", "..", "CPIPM-Project-Outline-v1.0.docx")
OUT = os.path.join(os.path.dirname(__file__), "..", "references", "CPIPM-Project-Outline-v1.1.docx")

# ── Height tiers (twips; 1440 twips = 1 inch) ────────────────────────────────
H1 = 360    # 0.25"  — 1 visual row
H2 = 1080   # 0.75"  — 2 visual rows  (= 3 × 0.25", gives breathing room)
H3 = 1440   # 1.00"  — 3+ visual rows (= 4 × 0.25")

# Estimated usable characters per line by number of columns.
# Based on ~6.2" usable table width, typical 10–11pt body text ≈ 13–14 cpl/inch.
# Widest single cell dominates, so these are conservative (wide-column estimates).
CHARS_PER_LINE = {1: 90, 2: 75, 3: 60, 4: 45, 5: 35, 6: 25}
CHARS_DEFAULT  = 20   # > 6 columns

# ── Text replacements ─────────────────────────────────────────────────────────
REPLACEMENTS = [
    ("CP Inline\n",            "VCCS Command Pilot\u2122 Inline\n"),
    ("CP INLINE v1.0",         "Command Pilot Inline v1.0"),
    ("CP INLINE",              "Command Pilot Inline"),
    ("CP Inline is",           "Command Pilot Inline is"),
    ("CP Inline reuses",       "Command Pilot Inline reuses"),
    ("CP Inline v",            "Command Pilot Inline v"),
    ("CP Inline —",            "Command Pilot Inline —"),
    ("CP Inline,",             "Command Pilot Inline,"),
    ("CP Inline.",             "Command Pilot Inline."),
    ("CP Inline ",             "Command Pilot Inline "),
    ("CP Inline",              "Command Pilot Inline"),
    # Version string in title header
    ("Version 1.0 |",          "Version 1.1 |"),
    # Date format
    ("2026-08-03",             "3 August 2026"),
    # Author
    ("CP / Engineering",       "VCCS/PIPS / Engineering"),
]

def replace_text(text):
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    return text

def fix_paragraph(para):
    """Merge all runs, apply replacements, restore into first run."""
    full = "".join(r.text for r in para.runs)
    new  = replace_text(full)
    if new == full:
        return
    if para.runs:
        para.runs[0].text = new
        for r in para.runs[1:]:
            r.text = ""

# ── Row-height helpers ────────────────────────────────────────────────────────
def cell_visual_rows(cell, cpl):
    """Estimate visual line count for a single cell."""
    texts = [p.text.strip() for p in cell.paragraphs if p.text.strip()]
    if not texts:
        return 0
    total = 0
    for t in texts:
        total += max(1, math.ceil(len(t) / cpl))
    return total

def row_twips(table_row, n_cols):
    """Choose H1/H2/H3 based on the tallest cell in the row."""
    cpl  = CHARS_PER_LINE.get(n_cols, CHARS_DEFAULT)
    seen = set()
    max_vr = 0
    for cell in table_row.cells:
        tc_id = id(cell._tc)
        if tc_id in seen:
            continue            # skip merged-cell duplicates
        seen.add(tc_id)
        vr = cell_visual_rows(cell, cpl)
        if vr > max_vr:
            max_vr = vr
    if max_vr <= 1:
        return H1
    elif max_vr == 2:
        return H2
    else:
        return H3

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
def copy_row_format(src_row, dst_row):
    """Copy trPr XML from src to dst (preserves shading / borders)."""
    src_trPr = src_row._tr.find(qn("w:trPr"))
    if src_trPr is None:
        return
    import copy
    dst_tr   = dst_row._tr
    old_trPr = dst_tr.find(qn("w:trPr"))
    if old_trPr is not None:
        dst_tr.remove(old_trPr)
    dst_tr.insert(0, copy.deepcopy(src_trPr))

def append_version_row(table):
    """Add a v1.1 row to the version history table (Table 15)."""
    import copy
    # Clone the last data row to inherit formatting
    src_row  = table.rows[-1]
    new_tr   = copy.deepcopy(src_row._tr)
    table._tbl.append(new_tr)
    new_row  = table.rows[-1]
    values   = ["1.1", "4 August 2026", "VCCS/PIPS / Engineering",
                "Applied proportional row heights (Option B); "
                "updated terminology CP Inline \u2192 Command Pilot Inline throughout."]
    for cell, val in zip(new_row.cells, values):
        # Clear existing text
        for para in cell.paragraphs:
            for run in para.runs:
                run.text = ""
        # Write new text into first paragraph, first run
        if cell.paragraphs:
            p = cell.paragraphs[0]
            if p.runs:
                p.runs[0].text = val
            else:
                p.add_run(val)
        set_cell_valign(cell, "center")

# ── Main ──────────────────────────────────────────────────────────────────────
doc = Document(SRC)

# 1. Fix body paragraphs
for para in doc.paragraphs:
    fix_paragraph(para)

# 2. Process tables
tables = doc.tables
for ti, table in enumerate(tables):
    n_cols = len(table.columns)
    is_last = (ti == len(tables) - 1)   # version history table

    for row in table.rows:
        # Fix cell text
        for cell in row.cells:
            for para in cell.paragraphs:
                fix_paragraph(para)
            set_cell_valign(cell, "center")
        # Apply proportional row height
        set_row_height(row, row_twips(row, n_cols))

    # Append v1.1 version row to the last table
    if is_last:
        append_version_row(table)

# 3. Save
os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print(f"Saved: {OUT}")
