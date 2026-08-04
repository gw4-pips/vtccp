"""
Update CPIPM-Project-Outline-v1.0.docx:
  1. Apply Option B row heights (432 twips min, atLeast) to every table row
  2. Vertically centre all table cells
  3. Terminology: "CP Inline" → "Command Pilot Inline" throughout
  4. Header/title line 1: "CP Inline" → "VCCS Command Pilot™ Inline"
  5. Date: "2026-08-03" → "3 August 2026"
  6. Version table author: "CP / Engineering" → "VCCS/PIPS / Engineering"
"""
import copy, os, re
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = os.path.join(os.path.dirname(__file__), "..", "..", "CPIPM-Project-Outline-v1.0.docx")
OUT = os.path.join(os.path.dirname(__file__), "..", "references",
                   "CPIPM-Project-Outline-v1.1.docx")

MIN_ROW_HEIGHT = 432   # twips  ≈ 0.30 in  (Option B)

# ── Text replacements (applied to every paragraph in body + tables) ──────────
# Order matters — longest/most-specific first.
REPLACEMENTS = [
    # Title header line 1
    ("CP Inline\n",                         "VCCS Command Pilot\u2122 Inline\n"),
    # Uppercase wireframe status bar
    ("CP INLINE v1.0",                      "Command Pilot Inline v1.0"),
    ("CP INLINE",                           "Command Pilot Inline"),
    # General prose
    ("CP Inline is",                        "Command Pilot Inline is"),
    ("CP Inline reuses",                    "Command Pilot Inline reuses"),
    ("CP Inline v",                         "Command Pilot Inline v"),
    ("CP Inline —",                         "Command Pilot Inline —"),
    ("CP Inline,",                          "Command Pilot Inline,"),
    ("CP Inline.",                          "Command Pilot Inline."),
    ("CP Inline ",                          "Command Pilot Inline "),
    # Remaining bare "CP Inline" at end of string / before punctuation
    ("CP Inline",                           "Command Pilot Inline"),
    # Date
    ("2026-08-03",                          "3 August 2026"),
    # Version table author
    ("CP / Engineering",                    "VCCS/PIPS / Engineering"),
]

def replace_in_text(text):
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    return text

def fix_paragraph_text(para):
    """Consolidate all runs into one, apply replacement, restore run."""
    full = "".join(r.text for r in para.runs)
    new_full = replace_in_text(full)
    if new_full == full:
        return
    # Put new text in first run, blank the rest
    if para.runs:
        para.runs[0].text = new_full
        for r in para.runs[1:]:
            r.text = ""

def set_row_min_height(row, twips):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    # Remove any existing trHeight to avoid duplicates
    for el in trPr.findall(qn("w:trHeight")):
        trPr.remove(el)
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"),   str(twips))
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)

def set_cell_valign(cell, align="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for el in tcPr.findall(qn("w:vAlign")):
        tcPr.remove(el)
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), align)
    tcPr.append(vAlign)

# ── Load ──────────────────────────────────────────────────────────────────────
doc = Document(SRC)

# ── 1. Fix body paragraphs ────────────────────────────────────────────────────
for para in doc.paragraphs:
    fix_paragraph_text(para)

# ── 2. Fix tables: row height + valign + text replacement ────────────────────
for table in doc.tables:
    for row in table.rows:
        set_row_min_height(row, MIN_ROW_HEIGHT)
        for cell in row.cells:
            set_cell_valign(cell, "center")
            for para in cell.paragraphs:
                fix_paragraph_text(para)

# ── 3. Save ───────────────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print(f"Saved: {OUT}")
