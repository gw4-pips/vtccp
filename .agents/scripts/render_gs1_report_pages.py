import fitz
from pathlib import Path

source = Path("attached_assets/GS1_GenSpecs_2026_1788710621616.pdf")
output = Path(".agents/outputs/gs1-report-pages")
output.mkdir(parents=True, exist_ok=True)

doc = fitz.open(source)
for page_number in range(452, 456):
    page = doc.load_page(page_number)
    pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    pixmap.save(output / f"page-{page_number + 1}.png")

print(f"Rendered {doc.page_count} page document pages 453-456 to {output}")