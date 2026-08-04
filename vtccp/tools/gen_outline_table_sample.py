"""
Sample: Table 1 from CPIPM Project Outline — shows proposed open row spacing.
Generates a one-page comparison: current tight style vs. proposed open style.
"""
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm, Twips
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = os.path.join(os.path.dirname(__file__), "..", "references",
                   "CPIPM-TableStyle-Sample.docx")

NAVY  = RGBColor(0x1E, 0x3A, 0x5F)
TEAL  = RGBColor(0x0D, 0x94, 0x88)
BLACK = RGBColor(0x11, 0x18, 0x27)
GREY  = RGBColor(0x6B, 0x72, 0x80)

ROWS = [
    ("Target application",  "UDI labeling inline inspection — single station, single scan lane"),
    ("Scanner",             "Cognex DM-475V-LBL on desktop stand adapted for conveyor use"),
    ("Barcode type",        "GS1 DataMatrix only"),
    ("Trigger mode",        "Photosensor hardware item-presence (default); auto-cycle and manual also supported"),
    ("Indicator pole",      "3–5 segment light pole with per-color steady/flash states; audible alarm — TBD Engineering"),
    ("Conveyor control",    "Pusher divert on fail; line-stop behaviour TBD Engineering"),
    ("Timeline",            "Wk 1–3: Build & test  |  Wk 4: PE trial run  |  Wk 5: Harden  |  Wk 6: Go-live"),
    ("Physical product",    "Small folding cartons (~cell phone to paperback size, 1.25–1.75\" thick), "
                            "conveyed flat, narrow edge leading — virus and diagnostic test kits"),
]

def shd_cell(cell, hex_fill):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_fill)
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=108, right=108):
    """Cell margins in twentieths of a point (twips). 108 twips ≈ 0.075 in."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)

def make_table(doc, font_size_pt, cell_margin_twips, space_before_pt, space_after_pt,
               label=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(3)
    r = p.add_run(label)
    r.bold = True; r.font.size = Pt(10)
    r.font.color.rgb = NAVY

    tbl = doc.add_table(rows=1 + len(ROWS), cols=2)
    tbl.style = "Table Grid"

    # Set column widths: 2.0" parameter, 4.5" value
    for row in tbl.rows:
        row.cells[0].width = Inches(2.0)
        row.cells[1].width = Inches(4.5)

    # Header row
    hdr = tbl.rows[0]
    for i, txt in enumerate(["Parameter", "Value"]):
        cell = hdr.cells[i]
        cell.text = txt
        set_cell_margins(cell, top=cell_margin_twips, bottom=cell_margin_twips)
        shd_cell(cell, "1E3A5F")
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(font_size_pt)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            para.paragraph_format.space_before = Pt(space_before_pt)
            para.paragraph_format.space_after  = Pt(space_after_pt)

    # Data rows
    for ri, (param, value) in enumerate(ROWS):
        row = tbl.rows[ri + 1]
        fill = "EBF8F7" if ri % 2 == 0 else "FFFFFF"
        for ci, txt in enumerate([param, value]):
            cell = row.cells[ci]
            cell.text = txt
            set_cell_margins(cell, top=cell_margin_twips, bottom=cell_margin_twips)
            shd_cell(cell, fill)
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(font_size_pt)
                    if ci == 0:
                        run.bold = True
                        run.font.color.rgb = NAVY
                    else:
                        run.font.color.rgb = BLACK
                para.paragraph_format.space_before = Pt(space_before_pt)
                para.paragraph_format.space_after  = Pt(space_after_pt)

    doc.add_paragraph()

# ── Build sample doc ─────────────────────────────────────────────────────────
doc = Document()
for section in doc.sections:
    section.top_margin    = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin   = Inches(1.0)
    section.right_margin  = Inches(1.0)

p = doc.add_paragraph()
r = p.add_run("CPIPM Table Style Sample — Table 1 (Project Summary)")
r.bold = True; r.font.size = Pt(13); r.font.color.rgb = NAVY

p2 = doc.add_paragraph()
r2 = p2.add_run("Two options shown below. Reply with A or B (or suggest tweaks).")
r2.italic = True; r2.font.size = Pt(9); r2.font.color.rgb = GREY

make_table(doc,
    font_size_pt=9.5,
    cell_margin_twips=60,   # ~tight current style
    space_before_pt=1,
    space_after_pt=1,
    label="Option A — current (tight): 9.5 pt, minimal cell padding")

make_table(doc,
    font_size_pt=10,
    cell_margin_twips=120,  # ~0.083 in top/bottom padding per cell
    space_before_pt=2,
    space_after_pt=2,
    label="Option B — proposed (open): 10 pt, increased cell padding")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print(f"Saved: {OUT}")
