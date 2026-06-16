#!/usr/bin/env python3
"""Dump the full structure + formatting fingerprint of a .docx weekly report.

Usage:
    python inspect_docx.py path/to/report.docx

Walks the document body **in reading order** so you see the real layout — where
images sit, where tables sit, which heading each falls under — not just a flat
paragraph list. For each paragraph it prints style, alignment, indentation,
spacing, automatic numbering, the first run's font, any inline images (with
size), and a text preview. For each table it prints position, dimensions,
alignment, whether it has borders, column widths, and a small cell preview —
and flags "figure tables" (a table used to hold an image + caption). Then it
prints document-level defaults (fonts, spacing).

The point is so a generated report can match EVERYTHING the user's report does —
fonts, spacing, numbering, image placement/size, table layout — by learning it
from the file, not guessing from a rendered view.
"""
import sys

# Reports are Chinese; force UTF-8 stdout so the console codepage (e.g. GBK on
# Windows) doesn't turn the text preview and font names into mojibake.
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
    from docx.table import Table
except ImportError:
    sys.exit(
        "python-docx is not installed. Run:  pip install python-docx\n"
        "Then re-run this script."
    )

EMU_PER_CM = 360000.0


def emu_to_str(val):
    if val is None:
        return "-"
    try:
        return f"{val.pt:.1f}pt"
    except Exception:
        return str(val)


def run_font(run):
    """Return (latin_font, eastasia_font, size, bold, italic) for a run."""
    f = run.font
    latin = f.name
    size = f.size.pt if f.size is not None else None
    bold = f.bold
    italic = f.italic
    eastasia = None
    rpr = run._element.rPr
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            eastasia = rfonts.get(qn("w:eastAsia"))
    return latin, eastasia, size, bold, italic


def para_numbering(p):
    """Return (numId, ilvl) if the paragraph uses automatic numbering, else None.

    Auto numbers (一、二、三 / 1、2、3) live in numPr and do NOT appear in the
    paragraph text, so they're invisible to a plain text read — surface them
    here so generated reports don't silently drop section numbering."""
    ppr = p._p.find(qn("w:pPr"))
    if ppr is None:
        return None
    numpr = ppr.find(qn("w:numPr"))
    if numpr is None:
        return None
    numid = numpr.find(qn("w:numId"))
    ilvl = numpr.find(qn("w:ilvl"))
    nid = numid.get(qn("w:val")) if numid is not None else "?"
    lvl = ilvl.get(qn("w:val")) if ilvl is not None else "0"
    return (nid, lvl)


def para_images(p):
    """Return a list of (width_cm, height_cm) for every inline image in the
    paragraph (w:drawing or legacy w:pict). Empty list if none. This is how we
    learn whether the report uses figures, how big, and — via document-order
    placement — where they go."""
    out = []
    for drawing in p._p.findall(".//" + qn("w:drawing")):
        ext = drawing.find(".//" + qn("wp:extent"))
        if ext is not None:
            try:
                cx = int(ext.get("cx")); cy = int(ext.get("cy"))
                out.append((cx / EMU_PER_CM, cy / EMU_PER_CM))
                continue
            except Exception:
                pass
        out.append((None, None))
    for _ in p._p.findall(".//" + qn("w:pict")):
        out.append((None, None))
    return out


def para_indent(p):
    pf = p.paragraph_format
    return {
        "left": emu_to_str(pf.left_indent),
        "first_line": emu_to_str(pf.first_line_indent),
        "line_spacing": pf.line_spacing,
        "before": emu_to_str(pf.space_before) if pf.space_before else "-",
        "after": emu_to_str(pf.space_after) if pf.space_after else "-",
    }


ALIGN_NAMES = {
    None: "default",
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
}


def _spacing_from_ppr(ppr):
    """Read line/before/after spacing from a w:pPr element, or None."""
    if ppr is None:
        return None
    sp = ppr.find(qn("w:spacing"))
    jc = ppr.find(qn("w:jc"))
    out = {}
    if sp is not None:
        line = sp.get(qn("w:line"))
        rule = sp.get(qn("w:lineRule"))
        before = sp.get(qn("w:before"))
        after = sp.get(qn("w:after"))
        if line is not None:
            if rule in (None, "auto"):
                out["line"] = f"{int(line)/240:.2f}x"
            else:
                out["line"] = f"{int(line)/20:.1f}pt ({rule})"
        if before is not None:
            out["before"] = f"{int(before)/20:.1f}pt"
        if after is not None:
            out["after"] = f"{int(after)/20:.1f}pt"
    if jc is not None:
        out["align"] = jc.get(qn("w:val"))
    return out or None


def spacing_summary(document):
    """Effective spacing defaults: docDefaults pPr + Normal style pPr."""
    out = {"docDefaults": None, "Normal": None}
    styles_el = document.styles.element
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is not None:
        ppr_default = dd.find(qn("w:pPrDefault"))
        if ppr_default is not None:
            out["docDefaults"] = _spacing_from_ppr(ppr_default.find(qn("w:pPr")))
    for s in styles_el.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == "Normal":
            out["Normal"] = _spacing_from_ppr(s.find(qn("w:pPr")))
            break
    return out


def doc_defaults(document):
    """Best-effort read of the document default font from styles.xml."""
    out = {"default_font": None, "default_eastasia": None, "default_size": None}
    try:
        styles_el = document.styles.element
        dflt = styles_el.find(qn("w:docDefaults"))
        if dflt is not None:
            rpr = dflt.find(qn("w:rPrDefault"))
            if rpr is not None:
                r = rpr.find(qn("w:rPr"))
                if r is not None:
                    rfonts = r.find(qn("w:rFonts"))
                    if rfonts is not None:
                        out["default_font"] = rfonts.get(qn("w:ascii"))
                        out["default_eastasia"] = rfonts.get(qn("w:eastAsia"))
                    sz = r.find(qn("w:sz"))
                    if sz is not None:
                        val = sz.get(qn("w:val"))
                        if val:
                            out["default_size"] = f"{int(val)/2:.1f}pt"
    except Exception as e:
        out["error"] = str(e)
    return out


def table_info(t):
    """Describe a table's layout: dimensions, alignment, borders, column widths,
    and whether it looks like a 'figure table' (holds an image + caption)."""
    rows, cols = len(t.rows), len(t.columns)
    # alignment
    align = "default"
    jc = t._tbl.tblPr.find(qn("w:jc")) if t._tbl.tblPr is not None else None
    if jc is not None:
        align = jc.get(qn("w:val"))
    # borders: explicit tblBorders, or inherited via a *Grid* style
    has_borders = False
    if t._tbl.tblPr is not None:
        if t._tbl.tblPr.find(qn("w:tblBorders")) is not None:
            has_borders = True
    style_name = t.style.name if t.style is not None else None
    if style_name and ("Grid" in style_name or "网格" in style_name):
        has_borders = True
    # column widths
    widths = []
    for col in t.columns:
        w = col.width
        widths.append(f"{w.cm:.1f}cm" if w is not None else "-")
    # figure-table? any cell holds an image
    has_image = bool(t._tbl.findall(".//" + qn("w:drawing"))) or \
        bool(t._tbl.findall(".//" + qn("w:pict")))
    # small text preview of each cell (first row or two)
    preview_cells = []
    for r in t.rows[:3]:
        for c in r.cells:
            txt = c.text.strip().replace("\n", " ")
            if txt:
                preview_cells.append(txt[:24])
    return {
        "rows": rows, "cols": cols, "align": align,
        "borders": has_borders, "style": style_name,
        "widths": widths, "has_image": has_image,
        "preview": preview_cells[:6],
    }


def is_heading(p, num):
    """Heuristic: a paragraph that acts as a section/subsection heading — has
    automatic numbering, or a Heading/Title style, or is short+bold."""
    style = p.style.name if p.style is not None else ""
    if num is not None:
        return True
    if style and ("Heading" in style or "Title" in style or "标题" in style):
        return True
    return False


def main():
    if len(sys.argv) != 2:
        sys.exit("Usage: python inspect_docx.py path/to/report.docx")
    path = sys.argv[1]
    document = docx.Document(path)

    print(f"=== Document: {path} ===\n")

    defs = doc_defaults(document)
    print("--- Document defaults ---")
    print(f"  default font (latin) : {defs.get('default_font')}")
    print(f"  default font (东亚)   : {defs.get('default_eastasia')}")
    print(f"  default size         : {defs.get('default_size')}")
    if defs.get("error"):
        print(f"  (note: {defs['error']})")
    print()

    sp = spacing_summary(document)
    print("--- Spacing defaults (line/before/after/align) ---")
    print(f"  docDefaults : {sp['docDefaults']}")
    print(f"  Normal style: {sp['Normal']}")
    print("  (None means inherited from the application default — for python-docx")
    print("   output that default is 1.15 line + 8pt after, which is LOOSER than")
    print("   a typical Chinese report; set it explicitly to match the template.)")
    print()

    # ---- ordered walk: paragraphs AND tables interleaved, tracking headings ----
    body = document.element.body
    n_imgs = 0
    n_tbls = 0
    cur_heading = "(none)"
    pidx = -1

    print("--- Body in reading order (paragraphs + tables + images) ---")
    print("    (img/table lines note which heading they fall under — that's the")
    print("     placement convention to reproduce.)\n")

    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            pidx += 1
            p = Paragraph(child, document)
            text = p.text.strip()
            style = p.style.name if p.style is not None else "?"
            align = ALIGN_NAMES.get(p.alignment, str(p.alignment))
            ind = para_indent(p)
            if p.runs:
                latin, eastasia, size, bold, italic = run_font(p.runs[0])
            else:
                latin = eastasia = size = bold = italic = None
            num = para_numbering(p)
            if is_heading(p, num) and text:
                cur_heading = text[:30]
            imgs = para_images(p)
            preview = (text[:50] + "…") if len(text) > 50 else text
            flags = []
            if bold:
                flags.append("bold")
            if italic:
                flags.append("italic")
            flagstr = ",".join(flags) if flags else "-"
            numstr = f" num(id={num[0]},lvl={num[1]})" if num else ""
            print(
                f"[p{pidx:>3}] style={style!r} align={align} "
                f"indent(L={ind['left']},first={ind['first_line']}) "
                f"ls={ind['line_spacing']} before={ind['before']} "
                f"after={ind['after']}{numstr}"
            )
            print(f"        font: latin={latin} 东亚={eastasia} size={size} {flagstr}")
            if imgs:
                n_imgs += len(imgs)
                for w, h in imgs:
                    dim = f"{w:.1f}×{h:.1f}cm" if w else "size?"
                    print(f"        >>> IMAGE ({dim})  under heading: {cur_heading!r}")
            print(f"        text: {preview!r}")
        elif child.tag == qn("w:tbl"):
            n_tbls += 1
            t = Table(child, document)
            info = table_info(t)
            kind = "FIGURE-TABLE (image+caption)" if info["has_image"] else "data table"
            print(
                f"[tbl  ] {kind}: {info['rows']}r×{info['cols']}c "
                f"align={info['align']} borders={info['borders']} "
                f"style={info['style']!r}"
            )
            print(f"        col widths: {info['widths']}  under heading: {cur_heading!r}")
            if info["preview"]:
                print(f"        cells: {info['preview']}")
            if info["has_image"]:
                n_imgs += 1

    print(f"\n--- Layout summary: {n_imgs} image(s), {n_tbls} table(s) ---")
    print("  To match this report: reproduce each image/table at the same place")
    print("  (under the same heading), same size/borders/alignment, and number")
    print("  figures by order of appearance. Cloning the template is the reliable")
    print("  way to inherit fonts, spacing and numbering for free.")

    print("\n=== end ===")


if __name__ == "__main__":
    main()
