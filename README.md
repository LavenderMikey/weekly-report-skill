# weekly — 周报生成 Skill

一个用来**模仿你过去的周报**自动生成新一周周报的 Agent Skill，**同时兼容 Claude Code 和 OpenAI Codex**。它会先把你的历史周报当模板学习（章节结构、语气、详略，以及字体、字号、缩进、行间距、自动编号等视觉格式），再把你这周的工作和读过的文章/论文忠实地总结进同样的格式里。

核心原则：**绝不编造未发生的工作、指标或文章**——生成的周报要让领导看不出是自动生成的。

## 功能特点

- **格式高保真**：通过 `.docx` 克隆继承原周报的字体、行间距、页面设置，以及隐藏在 `numbering.xml` 里、文本看不到的**自动编号**（一、二、三 / 1、2、3）。
- **原生 Word 公式**：论文公式用 Pandoc 转成可编辑的 Word 公式（OMML），而不是 LaTeX 源码文本，无需手动转换。
- **从论文 PDF 自动抠图**：按图注（caption）定位裁剪论文里的架构图/实验图，以插入表格的形式放进周报。
- **按历史排版规律放图**：架构图放在"整体框架/方法"小节下，结果图放在"实验结果"下，按出现顺序编号（图1、图2…），而不是堆在末尾。
- **学习板块独立**：读过的文章单独成节，不与工作内容混在一起。

## 使用方式

脚本和参考文档是工具无关的纯 Python + Markdown，两边共用；区别只在"入口/触发方式"。

### 🟣 Claude Code

把整个目录放到 skills 目录下（例如 `~/.claude/skills/weekly`），然后对 Claude 说"帮我写周报"、"根据以前的周报生成这周的"之类的话，会**按 [SKILL.md](SKILL.md) 的 description 自动触发**。

### 🟢 OpenAI Codex（及其它读 AGENTS.md 的 agent）

仓库根目录的 [AGENTS.md](AGENTS.md) 就是 Codex 的入口。把它（连同 `scripts/`、`references/`）放进你的项目，然后说一句"按 AGENTS.md 的周报流程帮我写周报"，或直接把历史周报 + 本周材料丢给它。Codex 没有自动发现 / 自动触发机制，所以需要你稍微点一下它去走这个流程。

> 两个入口（SKILL.md / AGENTS.md）描述的是**同一套 5 步工作流**，调用的是**同一批脚本**，不用维护两份逻辑。

## 目录结构

```
weekly-report-skill/
├── SKILL.md                       # Claude Code 入口：触发条件 + 5 步工作流
├── AGENTS.md                      # Codex / 通用 agent 入口（同一套流程）
├── README.md
├── scripts/
│   ├── inspect_docx.py            # 转储 .docx 的逐段结构与格式（字体/字号/对齐/缩进/编号/行间距）
│   ├── docx_report.py             # 克隆重建周报的通用积木：wipe_body / clone_with_text / add_figure_table / strip_unused_media 等
│   ├── docx_math.py               # latex_to_omath：把 LaTeX 转成原生 Word 公式（需 Pandoc）
│   └── extract_pdf_figure.py      # extract_figure：按图注从论文 PDF 裁剪图片（需 PyMuPDF）
└── references/
    └── docx_formatting.md         # python-docx 配方：中文字体、缩进、克隆模板、插入公式
```

## 依赖

- `python-docx`（读写 .docx）
- `pandoc`（LaTeX → Word 公式，可选，仅含公式时需要）
- `pymupdf`（从 PDF 抠图，可选，仅需插图时需要）

```bash
pip install python-docx pymupdf
```
