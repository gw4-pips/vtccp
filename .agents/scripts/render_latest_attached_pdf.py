from pathlib import Path
import fitz

pdf_path = Path("attached_assets/2026-08-22_21-24-41_vccs_rfid_20260822-212124_1787448450248.pdf")
output_dir = Path(".agents/outputs/latest-attached-vccs-rfid")
output_dir.mkdir(parents=True, exist_ok=True)

with fitz.open(pdf_path) as document:
    for index, page in enumerate(document):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        pixmap.save(output_dir / f"page-{index + 1}.png")
    print(f"rendered {document.page_count} page(s) to {output_dir}")