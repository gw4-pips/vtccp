from pathlib import Path
import fitz


source = Path(
    "attached_assets/"
    "2026-08-19_00-59-27_vccs_rfid_20260818-205921_1787101334004.pdf"
)
output_dir = Path(".agents/outputs/attached-report")
output_dir.mkdir(parents=True, exist_ok=True)

with fitz.open(source) as document:
    print(f"pages={document.page_count} metadata={document.metadata}")
    for index, page in enumerate(document):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        output = output_dir / f"page-{index + 1}.png"
        pixmap.save(output)
        print(f"{output} size={page.rect}")