#!/usr/bin/env python3
"""Dump the structure and formatting fingerprint of a .docx weekly report.

Usage:
    python inspect_docx.py path/to/report.docx

Prints, per paragraph: index, style name, alignment, indentation, and the
font (name / East-Asian name / size / bold / italic) of its first run, plus a
text preview. Then prints document-level defaults (default font, East-Asian
font, size). This is meant to give you a reliable picture of how the user's
report is formatted so a generated report can match it, instead of guessing
from a rendered view.
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
except ImportError:
    sys.exit(
        "python-docx is not installed. Run:  pip install python-docx\n"
        "Then re-run this script."
    )


def emu_to_str(val):
    if val is None:
        return "-"
    # python-docx Length is in EMU; .pt and .cm helpers exist
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
    # East Asian font lives in rPr/rFonts@w:eastAsia and isn't exposed directly
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
            # line is in 240ths for auto/atLeast/exact; auto 240 == single
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
    """Effective spacing defaults: docDefaults pPr + Normal style pPr.

    Line spacing and paragraph spacing are usually inherited from here rather
    than set per-paragraph, so matching a template's 'feel' depends on these."""
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

    numbered = sum(1 for p in document.paragraphs if para_numbering(p))
    if numbered:
        print(f"--- Automatic numbering: {numbered} paragraph(s) use numPr ---")
        print("  Headings are likely auto-numbered (一、二、三 / 1、2、3). These")
        print("  numbers are NOT in the text — to reproduce them, reuse the same")
        print("  numId (cloning the template is the easiest way). See num(id=…) below.")
        print()

    print("--- Paragraphs ---")
    for i, p in enumerate(document.paragraphs):
        text = p.text.strip()
        style = p.style.name if p.style is not None else "?"
        align = ALIGN_NAMES.get(p.alignment, str(p.alignment))
        ind = para_indent(p)
        if p.runs:
            latin, eastasia, size, bold, italic = run_font(p.runs[0])
        else:
            latin = eastasia = size = bold = italic = None
        preview = (text[:50] + "…") if len(text) > 50 else text
        flags = []
        if bold:
            flags.append("bold")
        if italic:
            flags.append("italic")
        flagstr = ",".join(flags) if flags else "-"
        num = para_numbering(p)
        numstr = f" num(id={num[0]},lvl={num[1]})" if num else ""
        print(
            f"[{i:>3}] style={style!r} align={align} "
            f"indent(L={ind['left']},first={ind['first_line']}) "
            f"ls={ind['line_spacing']} before={ind['before']} after={ind['after']}"
            f"{numstr}"
        )
        print(
            f"      font: latin={latin} 东亚={eastasia} "
            f"size={size} {flagstr}"
        )
        print(f"      text: {preview!r}")

    # Tables, if any, are common in reports — note their presence.
    if document.tables:
        print(f"\n--- Tables: {len(document.tables)} found ---")
        for ti, t in enumerate(document.tables):
            print(f"  table[{ti}]: {len(t.rows)} rows x {len(t.columns)} cols")

    print("\n=== end ===")


if __name__ == "__main__":
    main()
