# AGENTS.md — 周报生成（Codex / 通用 Agent 入口）

> 这是给 **OpenAI Codex CLI** 及其它读取 `AGENTS.md` 的 agent 准备的入口。
> 它和 [SKILL.md](SKILL.md)（Claude Code 入口）共用同一套 `scripts/` 与
> `references/`，只是触发方式不同：Claude Code 会按 SKILL.md 的 description
> **自动唤起**；Codex 没有自动发现机制，请在对话里说一句"按 AGENTS.md 的周报
> 流程帮我写周报"来引导，或直接把历史周报 + 本周材料丢进来。

当用户想**写 / 生成 / 补全一份周报（周报 / weekly report / 工作周报）**，尤其是
提供了过去的周报想模仿其风格、要总结本周工作、或要把本周读的文章/论文整理进
学习板块时，按下面的流程来。

## 核心原则

生成的周报要让领导**看不出是自动写的**：既要还原内容结构与语气，也要还原视觉
格式（字体、字号、缩进、行间距、自动编号）。**绝不编造未发生的工作、指标或
文章**——宁可如实写得朴素，也不要夸大，因为用户要为这份周报负责。

## 工作流（5 步）

1. **研究历史周报** → 学结构 + 格式。
   - `.docx`：先跑 `python scripts/inspect_docx.py "旧周报.docx"`，它会逐段打印
     样式、字体、字号、对齐、缩进、行间距/段距、以及隐藏的自动编号
     （`num(id=,lvl=)`）。这是你的"格式指纹"。多份周报时，以**最近一份**为
     格式权威模板。
   - `.pdf` / 纯文本：直接读，从可见结构推断格式。
2. **向用户要本周材料**：① 本周工作（文件/文档/要点都行）；② 看过的文章/论文
   （链接、PDF、笔记）。材料不足以填满模板要求的板块时，**追问**，不要拿编造的
   内容凑数。
3. **忠实总结**：工作板块按模板的分组方式（按项目/主题/平铺）组织，突出成果但
   只写材料支持的内容；学习板块**单独成节**，每篇文章 1–3 行讲清是什么 + 启发/
   与工作的关联；有"下周计划"则据在办事项草拟并提示用户确认。
4. **同格式产出**（`.docx` 强烈建议**克隆模板**而非从零重建，见下）。
5. **交付**：给出文件路径 + 各板块填了什么 + 需用户核对处（尤其下周计划和材料
   单薄处）。

## .docx 的关键做法：克隆模板，别从零搭

真实周报带着很多文本看不见的格式：自动编号（存在 `numbering.xml`）、`Normal`
样式上的行距/段距/对齐、页边距、字体。从零重建很容易悄悄丢掉编号或行距。克隆
则全部免费继承。

**克隆—重建配方**（用 `scripts/docx_report.py` 的积木）：

1. `shutil.copy` 最近一份周报到输出路径，再用 `python-docx` 打开。
2. 从最近一周里**各采一个"原型"段落**（深拷贝其 `<w:p>`）：一级编号标题、
   小节编号标题、正文段、计划项、标题行。每个原型自带 `pPr`（含驱动编号的
   `numPr`）和 run 的 `rPr`。
3. `wipe_body(doc)` 清空正文里的所有块级元素（段落和**上几周残留的表格**），只留
   末尾的 `<w:sectPr>`；新内容用 `emit(doc, sectPr, el)` 插在它前面。
4. 用 `clone_with_text(原型, 文本)` / `clone_with_omath(原型, 公式)` 逐行重建。
   复用同一个 `numId` 的标题会被 Word 自动重新编号（综述 + 每篇论文 → 一、二、
   三；每篇论文一个小节 numId → 1、2、3；计划项用计划 numId）。正文/公式行不带
   编号（克隆正文原型，通常 `numId=0`）。
5. `doc.save(...)` 后调用 `strip_unused_media(path)` 删掉 `word/media/` 里没被
   引用的旧图，否则文件会胀到模板原大小（常 10MB+）。

**公式**：论文里的公式渲染成**可编辑的原生 Word 公式**，不要留 LaTeX 源码。用
`scripts/docx_math.py` 的 `latex_to_omath(...)`（经 Pandoc 转 OMML），再
`paragraph._p.append(element)`。需要 PATH 上有 Pandoc；没有就退回保留 LaTeX 文本
并告知用户。

**论文配图**：源材料是 PDF 时，用 `scripts/extract_pdf_figure.py` 的
`extract_figure(pdf, page, caption, out_png)` 按图注裁剪。**每张生成的 PNG 都要
打开核对裁剪**。然后用 `docx_report.add_figure_table(...)`（带框的居中表格：图 +
图注）或 `add_centered_image(...)`（无框居中图）插入，看历史周报用哪种就用哪种。
**按模板的位置放图**：架构图放"整体框架/方法"小节下，结果图放"实验结果"下，按
出现顺序编号（图1、图2…），别一股脑堆在末尾。拿不到 PDF 就放 `【此处插入图：…】`
占位并请用户贴截图。

中文字体坑：`python-docx` 里 `run.font.name` 只设拉丁字体，中文要另设 run 的
`w:eastAsia`。详见 [references/docx_formatting.md](references/docx_formatting.md)。

## 依赖

```bash
pip install python-docx pymupdf
# 公式功能另需 PATH 上有 pandoc
```

更细的脚本说明见 [SKILL.md](SKILL.md) 与 [README.md](README.md)。
