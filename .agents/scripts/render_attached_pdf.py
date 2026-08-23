from pathlib import Path

import fitz


source = Path("attached_assets/2026-08-22_20-54-05_vccs_rfid_20260822-205359_1787446506448.pdf")
output = Path(".agents/outputs/attached-vccs-rfid")
output.mkdir(parents=True, exist_ok=True)

document = fitz.open(source)
print(f"pages={document.page_count}")
for index, page in enumerate(document):
    image = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    destination = output / f"page-{index + 1}.png"
    image.save(destination)
    print(destination)