from pathlib import Path
import fitz

ROOT = Path("attached_assets")
OUT = Path(".agents/outputs/attached_pdf_renders")
OUT.mkdir(parents=True, exist_ok=True)

for pdf in sorted(ROOT.glob("*.pdf")):
    doc = fitz.open(pdf)
    print(f"{pdf.name}: {len(doc)} page(s)")
    for index, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        output = OUT / f"{pdf.stem}__page-{index + 1}.png"
        pix.save(output)
        print(f"  {output} {page.rect.width:.0f}x{page.rect.height:.0f}")