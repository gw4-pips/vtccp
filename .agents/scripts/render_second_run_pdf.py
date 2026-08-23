import fitz
from pathlib import Path
src = Path('attached_assets/2026-08-23_08-03-20_vccs_rfid_20260823-080312_1787489139014.pdf')
out = Path('.agents/outputs/second_run_pdf')
out.mkdir(parents=True, exist_ok=True)
doc = fitz.open(src)
print('pages', len(doc))
for i, page in enumerate(doc):
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    path = out / f'page-{i+1}.png'
    pix.save(path)
    print(path, page.rect, 'images', len(page.get_images(full=True)))
    print(page.get_text()[:1000].replace('\n', ' | '))
