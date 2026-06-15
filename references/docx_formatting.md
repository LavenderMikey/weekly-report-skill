# python-docx formatting recipes for matching a report's look

Read this when generating a `.docx` so the output's font, size, and indentation
actually match the user's historical report. The non-obvious parts are Chinese
(East Asian) fonts and first-line indentation, which trip people up constantly.

Install if needed: `pip install python-docx`

## 1. Set a Chinese (East Asian) font correctly

`run.font.name = "宋体"` only sets the **Latin** font. Chinese glyphs will fall
back to the document default unless you also set the East Asian font on the
run's `rPr/rFonts@w:eastAsia`. This is the single most common reason a generated
报告 "looks wrong."

```python
from docx.oxml.ns import qn

def set_font(run, latin=None, eastasia=None, size_pt=None, bold=None):
    if size_pt is not None:
        from docx.shared import Pt
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if latin is not None:
        run.font.name = latin
    if eastasia is not None:
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:eastAsia"), eastasia)
        # also set ascii/hAnsi so mixed text is consistent
        if latin is None:
            rfonts.set(qn("w:ascii"), eastasia)
            rfonts.set(qn("w:hAnsi"), eastasia)
```

Use the `latin` / `东亚` values that `inspect_docx.py` reported for the matching
paragraph type (heading vs. body).

## 2. Set the document default font (so you don't style every run)

Cleaner than touching every run: set the Normal style and the doc default,
including East Asian.

```python
from docx.shared import Pt
from docx.oxml.ns import qn

def set_default_font(document, latin, eastasia, size_pt):
    style = document.styles["Normal"]
    style.font.name = latin
    style.font.size = Pt(size_pt)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), eastasia)
```

## 3. First-line indent (首行缩进 2 字符)

Chinese body paragraphs are usually indented by two characters. Two-character
indent ≈ 2 × font size. For 小四 (12pt): `Pt(24)`. Or use character units:

```python
from docx.shared import Pt

p.paragraph_format.first_line_indent = Pt(24)  # ~2 chars at 12pt
# left indent for nested bullets:
p.paragraph_format.left_indent = Pt(21)
```

## 4. Clone the template and replace its body (preferred for stable layouts)

Keeps `styles.xml`, section setup, headers/footers, and default fonts exactly as
the user has them. Copy the file first, then edit content in place.

```python
import shutil, docx
shutil.copy("old_report.docx", "周报_2026-06-15.docx")
doc = docx.Document("周报_2026-06-15.docx")

# Option A: clear all body paragraphs, then re-add with captured styles
for p in list(doc.paragraphs):
    p._element.getparent().remove(p._element)

# then add new paragraphs using the styles you saw, e.g.:
h = doc.add_paragraph("一、本周工作", style="Heading 2")
body = doc.add_paragraph("本周完成了……")
body.paragraph_format.first_line_indent = Pt(24)
doc.save("周报_2026-06-15.docx")
```

If only the *content under fixed headings* changes, an even safer move is to
keep the heading paragraphs and only replace the paragraphs between them, so the
heading styling is untouched.

## 5. Add a heading that matches the template's heading look

If the template uses real Word heading styles, reuse them by name
(`doc.add_paragraph(text, style="Heading 1")`). If it fakes headings with bold
body text (common in Chinese reports), don't use Heading styles — replicate the
bold run instead:

```python
p = doc.add_paragraph()
run = p.add_run("二、学习与思考")
set_font(run, eastasia="黑体", size_pt=12, bold=True)
```

`inspect_docx.py`'s `style=` field tells you which case you're in: a real style
name like `'Heading 2'` vs. `'Normal'` with bold runs.

## 6. Native Word equations from LaTeX (for reports with formulas)

Don't dump LaTeX source as plain text — render real, editable Word equations so
the user doesn't have to convert anything. `scripts/docx_math.py` does the
conversion via Pandoc (LaTeX → OMML):

```python
from docx_math import latex_to_omath
from docx.enum.text import WD_ALIGN_PARAGRAPH

omml = latex_to_omath([
    r"\mathrm{Attention}(Q,K,V)=\mathrm{softmax}\!\left(\frac{QK^\top}{\sqrt d}+M\right)V",
])
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER   # display equations look right centered
p._p.append(omml[0])                      # paragraph now holds a real equation
```

Notes:
- Needs Pandoc on PATH (`pandoc --version`). If it's missing, leave the LaTeX as
  text and tell the user, rather than failing silently.
- Pass either bare LaTeX or `$...$` — the helper strips delimiters.
- Convert many formulas in one call (`latex_to_omath([f1, f2, ...])`) — it's one
  Pandoc invocation and returns them in order.
- Figures (architecture diagrams, result plots) are images, not math — you can't
  regenerate them. Insert a placeholder paragraph like `【此处插入图：…】` and ask
  the user to paste the screenshot.

## 7. Quick sanity check after generating

Re-run `inspect_docx.py` on your output and eyeball that the fonts/sizes/indents
line up with the source. It's the fastest way to catch a missing East Asian font
before handing the file to the user.
