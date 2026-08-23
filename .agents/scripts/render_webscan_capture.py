from pathlib import Path
import fitz

pdf_path = Path("attached_assets/Webscan_Report--26-08-22_21_53_29_1787450204733.pdf")
out_dir = Path(".agents/outputs/webscan-capture")
out_dir.mkdir(parents=True, exist_ok=True)

doc = fitz.open(pdf_path)
print(f"pages={doc.page_count}")
for index, page in enumerate(doc):
    pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    output = out_dir / f"page-{index + 1}.png"
    pixmap.save(output)
    print(output)

    for image_index, image in enumerate(page.get_images(full=True), start=1):
        xref = image[0]
        extracted = doc.extract_image(xref)
        image_path = out_dir / f"page-{index + 1}-embedded-{image_index}.{extracted['ext']}"
        image_path.write_bytes(extracted["image"])
        print(image_path)