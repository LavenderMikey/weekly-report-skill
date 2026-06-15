# -*- coding: utf-8 -*-
"""Convert LaTeX math into native, editable Word equations (OMML).

Many reports contain formulas. Writing LaTeX as plain text is ugly and the user
shouldn't have to convert it by hand, so we turn each LaTeX string into Word's
native equation format (OMML) using Pandoc, which has a reliable LaTeX->OMML
path. The resulting elements can be dropped straight into a python-docx
paragraph, so you keep full control of the surrounding formatting (fonts,
indentation, headings) while the equations render and edit like normal Word
equations.

Requirements: Pandoc on PATH (`pandoc --version`) and lxml. Pandoc is widely
available; if it's missing, install it (https://pandoc.org/installing.html) or
fall back to leaving the LaTeX as text and telling the user.

Typical use when building a report with python-docx:

    from docx_math import latex_to_omath
    omml = latex_to_omath([r"\\frac{QK^\\top}{\\sqrt{d}}+M"])
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p._p.append(omml[0])          # the paragraph now holds a real equation
"""
import os
import zipfile
import subprocess
import tempfile
from copy import deepcopy

from lxml import etree

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def latex_to_omath(latex_list):
    """Convert a list of LaTeX strings to a list of Word OMML elements.

    Surrounding ``$``/``$$`` delimiters are stripped if present, so you can pass
    either bare LaTeX or ``$...$``. Each formula becomes one display equation;
    the returned elements are in the same order as the input and are detached
    deep copies safe to append into any document."""
    cleaned = [s.strip().strip("$").strip() for s in latex_list]
    md = "\n\n".join("$$" + s + "$$" for s in cleaned)
    tmpdir = tempfile.mkdtemp()
    md_path = os.path.join(tmpdir, "f.md")
    docx_path = os.path.join(tmpdir, "f.docx")
    with open(md_path, "w", encoding="utf-8") as fh:
        fh.write(md)
    subprocess.run(["pandoc", md_path, "-o", docx_path], check=True)
    with zipfile.ZipFile(docx_path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    maths = root.findall(".//{%s}oMathPara" % M_NS)
    if len(maths) != len(cleaned):
        # Pandoc occasionally emits a bare oMath without the oMathPara wrapper.
        maths = root.findall(".//{%s}oMath" % M_NS)
    return [deepcopy(m) for m in maths]


if __name__ == "__main__":
    import sys
    # quick self-test: convert one formula and report success
    latex = sys.argv[1] if len(sys.argv) > 1 else r"E = mc^2"
    els = latex_to_omath([latex])
    print(f"Converted {len(els)} formula(s); first tag: {els[0].tag}")
