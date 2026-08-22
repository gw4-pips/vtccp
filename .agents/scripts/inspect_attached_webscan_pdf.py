from pathlib import Path
import fitz

pdf_path = Path("attached_assets/GS1_DataMatrix-26-08-22_10_44_34-_F1_0100696114704288217280328_1787411311901.pdf")
output_dir = Path(".agents/outputs/webscan-pdf")
output_dir.mkdir(parents=True, exist_ok=True)

doc = fitz.open(pdf_path)
print(f"pages={doc.page_count}")
print(f"metadata={doc.metadata}")

for page_number, page in enumerate(doc, start=1):
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    rendered = output_dir / f"page-{page_number}.png"
    pix.save(rendered)
    print(f"rendered={rendered} size={page.rect.width}x{page.rect.height}")
    images = page.get_images(full=True)
    print(f"page={page_number} embedded_images={len(images)}")
    for image_number, image in enumerate(images, start=1):
        xref = image[0]
        extracted = doc.extract_image(xref)
        suffix = extracted.get("ext", "bin")
        image_path = output_dir / f"page-{page_number}-image-{image_number}.{suffix}"
        image_path.write_bytes(extracted["image"])
        print(
            f"  image={image_number} xref={xref} "
            f"size={extracted.get('width')}x{extracted.get('height')} "
            f"ext={suffix} path={image_path}"
        )