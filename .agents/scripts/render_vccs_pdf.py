"""Render a supplied VCCS PDF page to PNG for visual report inspection."""

from pathlib import Path
import sys

import fitz


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("Usage: render_vccs_pdf.py <input.pdf> <output.png>")

    pdf_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with fitz.open(pdf_path) as document:
        if document.page_count < 1:
            raise ValueError("The PDF has no pages.")
        pixmap = document[0].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        pixmap.save(output_path)


if __name__ == "__main__":
    main()