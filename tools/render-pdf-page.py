#!/usr/bin/env python3
"""Render PDF page(s) to PNG so they can be inspected visually.

Usage: python tools/render-pdf-page.py <file.pdf> <page>[-<page>] [out_dir]

Pages are 1-based. Output: <out_dir or pdf's folder>/<pdfname>-p<N>.png at 150 dpi.
Needed because pdftoppm is not installed on this machine; PyMuPDF is.
"""
import os
import sys

import fitz  # PyMuPDF

# Same guard as find-ghg-pages.py: Windows cp1252 consoles crash on printing
# non-ASCII (e.g. accented report filenames), killing the run mid-way.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    pdf_path, page_spec = sys.argv[1], sys.argv[2]
    out_dir = sys.argv[3] if len(sys.argv) > 3 else os.path.dirname(os.path.abspath(pdf_path))

    if "-" in page_spec:
        first, last = (int(p) for p in page_spec.split("-", 1))
    else:
        first = last = int(page_spec)

    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    doc = fitz.open(pdf_path)
    for n in range(first, last + 1):
        if n < 1 or n > doc.page_count:
            print(f"skip p.{n}: out of range (1-{doc.page_count})")
            continue
        pix = doc[n - 1].get_pixmap(dpi=150)
        out = os.path.join(out_dir, f"{stem}-p{n}.png")
        pix.save(out)
        print(out)


if __name__ == "__main__":
    main()
