from pathlib import Path
import fitz

source = Path("attached_assets/DataMatrix-26-08-22_08_31_50-WEBSCAN_020_CAL._1787402227622.pdf")
output_dir = Path(".agents/outputs/webscan-evidence")
output_dir.mkdir(parents=True, exist_ok=True)

document = fitz.open(source)
print(f"pages={document.page_count}")
print(f"metadata={document.metadata}")

for index, page in enumerate(document, start=1):
    pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    output = output_dir / f"page-{index}.png"
    pixmap.save(output)
    print(f"rendered={output} size={pixmap.width}x{pixmap.height}")

    images = page.get_images(full=True)
    print(f"embedded_images_page_{index}={len(images)}")
    for image_index, image in enumerate(images, start=1):
        image_bytes = document.extract_image(image[0])
        image_output = output_dir / f"page-{index}-image-{image_index}.{image_bytes['ext']}"
        image_output.write_bytes(image_bytes["image"])
        print(f"extracted={image_output} size={image_bytes['width']}x{image_bytes['height']}")