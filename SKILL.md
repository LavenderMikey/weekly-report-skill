---
name: weekly
description: >-
  Generates a new weekly status report (周报) that closely imitates the user's
  own past reports — same section structure, tone, level of detail, and visual
  formatting (font, size, indentation, heading styles) — by reading one or more
  historical reports as a template, then summarizing the user's work and the
  articles they read this week into the matching format. Use this skill whenever
  the user wants to write, draft, generate, or fill in a 周报 / weekly report /
  weekly summary / 工作周报, especially when they provide past reports to copy the
  style from, mention summarizing this week's work, or want this week's reading
  ("看的文章/论文/资料") folded into a learning section. Trigger on phrases like
  "帮我写周报", "根据以前的周报生成这周的", "把这周的工作整理成周报", "周报", "weekly report",
  even if they don't explicitly say the word "skill".
---

# Weekly Report Generator (周报生成)

Produce a new weekly report that looks and reads like the user wrote it
themselves. The whole point is **fidelity to their existing reports**: a manager
should not be able to tell this one was generated. That means matching three
things — the *visual formatting*, the *content/writing style*, and the report's
core nature: **it is summarizing prose.** Never invent accomplishments the
materials don't support.

A weekly report is **condensed, summary language** — not a paste of whatever the
user sent you. Whatever they hand you (a paper, a long doc, rough notes, a code
diff) must be **read, understood, and distilled** into a few summary sentences in
*their* voice. So two kinds of fidelity matter, and form is only one of them:

- **Form** — fonts, spacing, numbering, layout (the easy-to-measure part).
- **Substance** — *how they summarize*: their terminology, how technical vs.
  high-level they pitch it, sentence patterns, how they phrase a takeaway, how
  much they compress. This is why you must understand the **content** of the past
  reports, not just their formatting. A report that nails the fonts but reads in
  a different voice still feels generated.

## Two paths — pick the short one when you can

Most complaints about this skill are "太长了，总在转格式." Usually that's because
the full pipeline (inspect → clone → LaTeX→Word → PDF→图) ran when half of it
wasn't needed. So first decide which path you're on:

- **🟢 Fast path (the common case — repeat week).** The user already has a report
  in the exact target format — almost always **last week's `.docx`, especially
  one this skill generated**. The format is already perfect *inside that file*,
  so **don't re-inspect or re-learn the formatting**: just clone it, swap in this
  week's content, save. Skip Step 1's formatting *dump* (`inspect_docx.py`).
  But still **read last week's prose** to lock in the voice and summarizing
  style — that's just reading text, it's cheap, and it's what keeps the new week
  sounding like the user. This is 80% of runs and should feel near one-shot.
- **🔵 Full path (first time, or no prior report in the right format).** Study one
  historical report once to learn structure + formatting (Step 1), then clone.

**The heavy converters are opt-in, not default.** LaTeX→Word equations (Pandoc)
and PDF→图 cropping (PyMuPDF) only run **when the report actually has math / the
user actually wants paper figures in it.** A plain work周报 touches neither —
don't invoke them reflexively. That alone removes most of the "总在转格式" feeling.

### Steps

1. *(Full path only)* Study a historical report → learn structure + formatting.
2. Ask for this week's materials (work + articles) — one message.
3. Summarize them faithfully into the structure.
4. Emit in the same format — **clone** the prior report; add formulas/figures
   only if the content needs them.
5. Hand back the file + a short per-section rundown.

## Step 1 — Study the historical report(s) *(full path only — skip on repeat weeks)*

**Skip this whole step on the fast path.** If a prior report in the right format
exists (e.g. last week's `.docx`), you don't need to dump and study its
formatting — cloning in Step 4 inherits all of it. Only do Step 1 the **first**
time you see a user's template, or when the cloned output looks off.

The user will provide one or more past reports. They may be `.docx`, `.pdf`, or
plain text/markdown. Your goal is to extract a **template** in your head (and
literally, when it's a `.docx`):

- **Section structure**: What are the recurring headings? (e.g. 本周工作 /
  本周完成 / 下周计划 / 学习与思考). In what order? Is work grouped by project, by
  day, or as a flat bullet list?
- **Granularity & voice — read the actual prose, don't just scan headings.**
  This is the part that's easy to under-do. Long prose or terse bullets? Past
  tense, results-oriented? Quantified ("完成 3 个模块", "测试覆盖率提升至 85%") or
  qualitative? First person or impersonal? And critically — **how do they
  summarize?** When their past report mentions a paper or a piece of work, how
  many sentences do they spend, at what level of abstraction, with what
  terminology, and how do they phrase the takeaway? Internalize that summarizing
  style; it's what you'll reproduce in Step 3. You are learning *content and
  voice here, not only layout.*
- **Visual formatting** (this matters because the user explicitly cares about it
  for an upper-management audience): font family, font size, bold/heading
  styles, indentation, bullet markers, numbering, and — easy to overlook —
  **line spacing and paragraph spacing (space before/after) and alignment**.
  These are usually set on the `Normal` style or document defaults rather than
  per paragraph, so `inspect_docx.py` reports them in a dedicated "Spacing
  defaults" section. Don't skip it: spacing is a big part of why a report
  "looks like mine," and it's the easiest thing to get subtly wrong.

### Reading the different formats

- **`.docx`** — Use the helper script, which dumps structure and formatting so
  you don't have to eyeball it:
  ```bash
  python "<skill-dir>/scripts/inspect_docx.py" "path/to/old_report.docx"
  ```
  It prints, per paragraph: the style name, text preview, font name, size,
  bold/italic, alignment, and indentation — plus the document's default font.
  This is your formatting fingerprint. If the script reports a missing
  dependency, install it first: `pip install python-docx`.
- **`.pdf`** — Read it with the Read tool (it handles PDFs natively). You'll get
  the content and rough layout but *not* exact font metrics; infer formatting
  from visual structure and ask the user if a specific detail matters.
- **plain text / markdown** — Read directly; structure is whatever the text
  shows.

If the user gives **several** historical reports, treat the **most recent** one
as the authoritative template for formatting, and use the others to confirm
which sections are stable vs. occasional.

## Step 2 — Ask for this week's materials

Ask for this week's inputs in **one** message (don't pre-announce that you
studied the template — just ask). Inputs come as both files and pasted text, so
invite both:

> 把这周的材料发我吧（沿用上周的格式）：
> 1. **本周工作**：做了哪些事（文件、文档、或直接贴要点都行）
> 2. **看过的文章/资料**：链接、PDF、或你自己记的笔记要点（没有就说没有）

Read every file they provide (`.docx`/`.pdf` via the methods above; code or docs
with the Read tool). Don't proceed to writing until you have enough to fill the
sections that the template says are required. If something's thin (e.g. they
gave work but no articles, and the template has a learning section), ask rather
than padding it with invented content.

**Identify gaps and offer to predict.** Compare the template's sections against
what the materials actually cover. Some sections are **forward-looking or
inferential** rather than factual — e.g. 下周计划 / next-week plan, 心得思考 /
reflections, 存在的问题 / 下一步打算 — and the user often hasn't written them out.
Don't just leave them blank (that dumps the work back on the user), and don't
silently invent them either. Instead, **explicitly ask whether the user wants
you to predict/draft these from this week's work plus the patterns in their past
reports** — make it an opt-in choice:

> 历史周报里有「下周计划」「心得思考」这几节，但你这周的材料没直接提到。
> 要不要我根据你**本周的工作内容** + **以前周报的规律**，帮你**预测/草拟**这些部分？
> 草稿你可以直接改、让我重写，或者你自己填都行；不需要的话我就留个占位标记。

If they say yes, draft grounded predictions in Step 3. If no, leave a clear
`【待补充：下周计划】`-style placeholder so they know exactly what to fill in.
Use `AskUserQuestion` if you want to offer the choice per-section (预测 / 留空 /
我自己说) rather than all-or-nothing.

## Step 3 — Understand, then summarize in their voice

This is the heart of the skill. A weekly report is **summarizing prose**, so for
every input — a paper, a long doc, rough notes, a diff — the move is the same:
**read it, understand it, then compress it into a few summary sentences that
match the voice and granularity you saw in Step 1.** Never paste source text or a
raw abstract; never just list what you read without distilling it. If you can't
say what something *means* in one or two sentences, you haven't understood it
well enough to summarize it yet — go back and read more closely.

- **Work section**: Group the way their template groups (by project/by theme/
  flat). Lead with outcomes and impact, since the audience is the user's
  manager — but only claims supported by the materials. If the user gave rough
  notes, elevate them into the report's voice; don't downgrade the report to
  their notes' casualness.
- **Learning section** (articles read): This is a **dedicated section** — the
  user wants their reading called out separately, not blended into the work
  items. For each article/paper, actually read enough to **understand it**, then
  give a one-to-three line summary in their style: what it's about, the key idea,
  and the takeaway or how it connects to their work — pitched at the same
  technical level and terminology their past learning sections use. Don't copy
  the abstract; say it the way *they* would. Keep it tighter than the work
  section unless their template does otherwise.
- **Predicted / forward-looking sections** (下周计划, 心得思考, 下一步打算, etc.):
  Only draft these if the user opted in back in Step 2 — otherwise leave the
  `【待补充：…】` placeholder. When you do draft them, **ground every prediction in
  something concrete**: in-progress items and obvious follow-ups from this week's
  actual work, plus the recurring shape these sections take in past reports (what
  kinds of plans/reflections the user usually writes, their granularity and
  phrasing). Then clearly flag them as drafts for the user to confirm or rewrite.
  This is the **one** place inference is allowed — it is forward-looking and
  opt-in, and it does **not** loosen the rule below.

The golden rule: **never fabricate past accomplishments, metrics, or articles.**
Predicting a *plan* the user opted into is fine; inventing *work that didn't
happen* or *numbers that weren't measured* is not. A report that overstates is
worse than one that's modestly accurate, because the user has to stand behind
it. When in doubt, summarize what's there and note the gap.

## Step 4 — Generate in the same format

### When the source is `.docx` (preserve formatting)

**Strongly prefer cloning the template over rebuilding from scratch.** A real
report carries a lot of formatting you can't see from the text: automatic
numbering (一、二、三 / 1、2、3 stored in `numbering.xml`, invisible in the text),
line/paragraph spacing and alignment set on the `Normal` style, page margins,
fonts. Every time you rebuild fresh you risk silently dropping one of these —
in practice numbering and spacing get lost. Cloning inherits all of it for free.

**Clone-and-rebuild-from-prototypes (the reliable recipe):**

1. `shutil.copy` the most recent report to the output path and open it with
   `python-docx`. This keeps `styles.xml`, `numbering.xml`, section/page setup,
   fonts and spacing exactly as the user has them.
2. From the most recent week's section, **harvest one "prototype" paragraph of
   each kind** by deep-copying its `<w:p>` element — e.g. a top-level numbered
   heading, a subsection numbered heading, a plain body paragraph, a plan-list
   item, the title line. Each prototype carries its own `pPr` (including the
   `numPr` that drives numbering) and run `rPr`.
3. Remove all existing `<w:p>` from the body, but keep the trailing
   `<w:sectPr>` (page setup). Insert new paragraphs *before* it
   (`sectPr.addprevious(el)`).
4. Re-emit the report by cloning the right prototype for each line and swapping
   in this week's text (keep the prototype's first run so its formatting and the
   paragraph's `numPr` survive; drop the extra runs/images).

Because cloned headings reuse the same `numId`, Word renumbers them
automatically: reuse the top-level `numId` for every section (综述 + each
paper → 一、二、三), one subsection `numId` per paper (its subsections restart at
1、2、3…), and the plan-list `numId` for next-week items. Body paragraphs and
equation lines should carry no numbering (often `numId=0`, which means "no
number") — clone the body prototype for those. `inspect_docx.py` prints
`num(id=…,lvl=…)` per paragraph and a numbering summary so you can see exactly
which prototype maps to which `numId`. `scripts/docx_report.py` bundles the
reusable building blocks for this recipe: `wipe_body`, `emit`,
`clone_with_text`, `clone_with_omath`, `add_figure_table`, `add_centered_image`
and `strip_unused_media`.

**If you must rebuild fresh** (no usable template, or the user wants a brand-new
layout): set the document default font including the East Asian font
(`w:eastAsia`, which python-docx omits), and beware that python-docx's default
template gives `Normal` loose spacing (1.15 line + ~8–10pt after). Most Chinese
reports are tight — override to match what `inspect_docx.py` reports:
`nf = doc.styles["Normal"].paragraph_format; nf.line_spacing = 1.0;
nf.space_before = Pt(0); nf.space_after = Pt(0)`. Reproducing automatic
numbering from scratch is painful (you must inject `numbering.xml`), which is
exactly why cloning is preferred.

Chinese-font gotcha: in `python-docx`, setting `run.font.name` only sets the
Latin font. To make Chinese text use the intended font you must also set the
East Asian font on the run's rPr. The helper notes whether the template uses a
distinct East Asian font; if so, apply it. See `references/docx_formatting.md`
for the exact snippet and other formatting recipes.

Formulas — **only if the report actually contains math** (skip entirely for a
plain work周报; don't run Pandoc for nothing). When there is math, render it as
**native, editable Word equations**, not as LaTeX source text — the user should
not have to convert anything by hand. Use `scripts/docx_math.py`'s
`latex_to_omath(...)`, which turns LaTeX into Word OMML via Pandoc, then append
each returned element into a paragraph (`paragraph._p.append(element)`). This
keeps your formatting intact while the equations render and edit like normal
Word equations. It needs Pandoc on PATH; if Pandoc isn't available, leave the
LaTeX as text and tell the user.

Figures from source papers — **only if the report should actually show paper
figures** (many weeks don't need any; if unsure, ask once with a yes/no rather
than extracting speculatively). When it does, and the sources are PDFs, you can
pull the figures out automatically instead of leaving placeholders. Use
`scripts/extract_pdf_figure.py` (`extract_figure(pdf, page, caption, out_png)`),
which crops a figure by anchoring on its caption text. **Always open each
produced PNG to verify the crop** — caption-anchored cropping occasionally clips
too much/little, and viewing it is the fast way to catch that before inserting.
Then insert each figure with `docx_report.add_figure_table(...)` (a centered
bordered table: image row + caption row) — or, if the template places figures as
plain centered images without a border, `docx_report.add_centered_image(...)`.
Match whichever the historical report uses. If you can't get the PDFs, fall back
to a `【此处插入图：…】` placeholder and ask the user to paste the screenshot.

**Place figures where the template places them, not bunched at the end.** Check
the historical report's figure positions (`inspect_docx.py` marks paragraphs
containing images): typically each figure sits inline right after the subsection
that discusses it — an architecture/framework diagram under the
"整体框架/方法" subsection, a results/comparison figure under "实验结果" — and
figures are numbered by order of appearance (图1、图2…). Match that flow so the
report reads like the user's, rather than dropping every figure at the section's
tail.

Cloning gotcha when inserting figures: the cloned template's body may contain
**tables and inline images from prior weeks**. Removing paragraphs doesn't
remove `<w:tbl>` elements, and old images linger in `word/media/` and bloat the
file. When wiping the body, remove *all* block children except the trailing
`<w:sectPr>` (use `docx_report.wipe_body`), and after saving call
`docx_report.strip_unused_media(path)` to drop images not referenced by
`document.xml` — otherwise the report balloons to the template's full size
(often 10+ MB) and carries stale content.

Save the result next to the inputs (or where the user asks) with a clear name
like `周报_2026-06-15.docx`.

### When the source is PDF or plain text

Match the structure and voice; produce the output in the format the user wants
(ask if unclear — often a `.docx` or markdown). You can't perfectly clone PDF
typography, so aim for clean, consistent formatting that mirrors the layout.

## Step 5 — Hand it back

Give the user the file path and a brief rundown: which sections you filled, what
you summarized into each, and anything you want them to verify (especially the
next-week plan and any place the materials were thin). Make it easy for them to
spot-check before they send it on.

## Resources

- `scripts/inspect_docx.py` — dumps a `.docx`'s per-paragraph structure and
  formatting (style, font, size, bold, alignment, indentation) + document
  defaults. Run it first on any `.docx` historical report.
- `scripts/docx_report.py` — report-agnostic building blocks for the
  clone-and-rebuild recipe: `wipe_body`, `emit`, `clone_with_text`,
  `clone_with_omath`, `add_figure_table`, `add_centered_image`,
  `strip_unused_media`.
- `scripts/docx_math.py` — `latex_to_omath(latex_list)` converts LaTeX into
  native Word equations (OMML) via Pandoc, for reports that contain formulas.
- `scripts/extract_pdf_figure.py` — `extract_figure(pdf, page, caption, out_png)`
  crops a figure out of a source PDF by anchoring on its caption (PyMuPDF).
  Verify each crop by opening the PNG before inserting it.
- `references/docx_formatting.md` — python-docx recipes: setting Chinese
  (East Asian) fonts, indentation, copying a template, matching styles, and
  inserting equations. Read it when generating a `.docx` so the output
  formatting actually matches.
