# -*- coding: utf-8 -*-
"""Crop a figure out of a paper PDF by anchoring on its caption.

Research reports embed figures (architecture diagrams, result plots) that the
user otherwise has to screenshot by hand. Given the caption text (e.g.
"Figure 2"), this finds the caption's text block — whose width equals the
figure's column width — then unions the vector drawings + raster images sitting
just above it (clipped to that column, so two-column pages don't bleed into the
neighbour) to get the figure's bounding box, and renders that region to a PNG.
Handles both vector plots and raster/diagram figures.

Requires PyMuPDF (`pip install pymupdf`).

CLI:
    python extract_pdf_figure.py paper.pdf 2 "Figure 2" out.png
            (pdf)              (1-based page) (caption)  (output png)

Programmatic:
    from extract_pdf_figure import extract_figure
    extract_figure("paper.pdf", page=2, caption="Figure 2", out_png="fig2.png")

After extracting, OPEN the PNG to eyeball the crop (it's easy to clip too much
or too little) and adjust `max_above` or the caption string if needed.
"""
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("PyMuPDF not installed. Run:  pip install pymupdf")


def extract_figure(pdf, page, caption, out_png, max_above=420, dpi=250):
    """page is 1-based (as printed in the paper). Returns the output path."""
    doc = fitz.open(pdf)
    pg = doc[page - 1]
    hits = pg.search_for(caption)
    if not hits:
        raise ValueError(f"caption {caption!r} not found on page {page}")
    hit = hits[0]
    blocks = pg.get_text("blocks")

    # the caption phrase is only a few chars wide; find the full caption block,
    # whose width is the figure's column width.
    cx0, cx1, cap_top = hit.x0, hit.x1, hit.y0
    for b in blocks:
        if b[0] - 1 <= hit.x0 and hit.x1 <= b[2] + 1 \
           and b[1] - 1 <= hit.y0 and hit.y1 <= b[3] + 1:
            cx0, cx1, cap_top = b[0], b[2], b[1]
            break

    # vertical extent from visual content (drawings + images) within the column
    # and above the caption — text blocks are unreliable because figures contain
    # internal labels.
    visual = [fitz.Rect(d["rect"]) for d in pg.get_drawings()]
    visual += [fitz.Rect(im["bbox"]) for im in pg.get_image_info()]
    top = cap_top
    for r in visual:
        if r.is_empty or r.width < 2 or r.height < 2:
            continue
        if r.y1 <= cap_top + 2 and r.y0 >= cap_top - max_above \
           and r.x1 > cx0 + 2 and r.x0 < cx1 - 2:
            top = min(top, r.y0)
    if top >= cap_top:
        top = max(pg.rect.y0 + 36, cap_top - max_above * 0.85)

    pad = 4
    fig = fitz.Rect(cx0 - pad, top - pad, cx1 + pad, cap_top - 1) & pg.rect
    pix = pg.get_pixmap(clip=fig, matrix=fitz.Matrix(dpi / 72, dpi / 72))
    pix.save(out_png)
    return out_png


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit("Usage: python extract_pdf_figure.py <pdf> <page> <caption> <out.png>")
    pdf, page, caption, out = sys.argv[1], int(sys.argv[2]), sys.argv[3], sys.argv[4]
    path = extract_figure(pdf, page, caption, out)
    print("saved:", path)
