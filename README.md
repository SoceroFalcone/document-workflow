# document-workflow

适合大学生课程作业、读书报告和论文的写作整理与 Word 排版 Skill。支持新建、校对、润色和重组 DOCX、Markdown、TXT，为 Word 文档应用版式，也可读取 PDF 作为资料。修改现有文件时默认另存副本。

## 默认 Word 结构与版式

普通文章采用「总标题 → 按需分层的正文」。摘要、关键词、目录、引文和参考文献在明确要求完整论文或相应部件时加入。

| 部件 | 默认设置 |
| --- | --- |
| 页面 | A4 纵向，上下边距 2.54 cm，左右边距 3.175 cm |
| 正文 | 中文小四宋体，西文小四 Times New Roman（均为 12 磅）。两端对齐，首行缩进 2 个中文字符 |
| 行距与段距 | 1.5 倍行距，段前、段后均为 0 磅，开启孤行控制 |
| 总标题 | 三号（16 磅）加粗、居中 |
| 一级标题 | 四号（14 磅）加粗、左对齐，编号如「一、」「二、」 |
| 二级标题 | 小四（12 磅）加粗、左对齐，编号如「（一）」「（二）」 |
| 三级标题 | 小四（12 磅）常规字重、左对齐，编号如「1、」「2、」 |
| 已有脚注正文 | 五号（10.5 磅） |

标题编号紧接标题文字。校对、润色和新建时，清理中文标点前后及数字后的多余空格，保留英语、单位、公式和代码中有意义的空格。

### 论文预设

- `academic-paper` 适合需要紧凑字号的学术论文。中文题名小二黑体，一级标题小四黑体，二级标题五号黑体，正文五号。
- `thesis` 提供封面、中英文摘要、关键词、目录标题及参考文献等部件的样式。完整学位论文可依次安排封面、摘要和关键词、目录、正文、参考文献及附录。封面和摘要隐藏页码，目录用大写罗马数字，正文用阿拉伯数字从 1 开始。
- 图题放图下，表题放表上。需要自动目录、多级标题编号、分节页码和跨页表头时，用 Word 的相应功能建立并核对。

排版规则逐项取自本次明确要求、指定模板中选用的规则、项目规范、论文预设、默认版式，优先级依次降低。可按需要提供版式模板或配置。[详细论文版式](references/paper-layout.md)

## 安装到 Codex

### 1. 选择安装范围

**项目级**让 Skill 在当前项目中可用。在项目根目录打开终端并运行下方命令，再用 Codex 打开该项目文件夹。

```sh
git clone https://github.com/SoceroFalcone/document-workflow.git .agents/skills/document-workflow
```

**用户级**适合在多个本地项目中复用。

```sh
git clone https://github.com/SoceroFalcone/document-workflow.git "$HOME/.agents/skills/document-workflow"
```

也可以在 GitHub 仓库页面选择 **Code → Download ZIP**，解压后把包含 `SKILL.md`、`references/` 和 `scripts/` 的文件夹命名为 `document-workflow`，放入对应的 `.agents/skills/` 目录。通过 Git 安装后，可在安装目录运行 `git pull --ff-only` 获取更新。

### 2. 准备 Word 脚本依赖

运行 Word 排版、校验或交叉引用脚本需要 Python 3.10+ 和 `python-docx` 1.2。若已有独立 Python 环境，在其中运行 `python -m pip install "python-docx==1.2.0"`。也可以在项目根目录创建虚拟环境。

Windows PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install "python-docx==1.2.0"
```

macOS 或 Linux

```sh
python3 -m venv .venv
./.venv/bin/python -m pip install "python-docx==1.2.0"
```

如果 Windows 的 Python 启动命令是 `py`，可用 `py -m venv .venv` 创建虚拟环境。

## 在 Codex 中使用

输入 `$document-workflow` 并从列表选择 Skill。若同时安装了两个同名副本，按路径选择当前任务要用的版本。安装后列表仍未显示时，重新打开 Codex。若依赖装在项目的 `.venv`，可在任务中指定用该环境里的 Python。

可直接发送下面任一示例，并把文件名换成自己的稿件。

```text
$document-workflow 写一页关于小组作业分工的课程报告，使用默认排版，依正文需要设置标题，保存为「课程报告.docx」。
```

```text
$document-workflow 只排版「原稿.docx」，保留文字和结构，按「排版模板.docx」的行距另存为「排版稿.docx」。
```

```text
$document-workflow 为「原稿.docx」中的图表和数字文献引文建立可跳转的交叉引用，按正文的首引顺序重排参考文献，并另存为「引用稿.docx」。
```

在 Windows 桌面 Word 中，可按住 `Ctrl` 单击引用编号跳转。移动图表或条目后，按 `Ctrl+A`、`F9` 更新域。如果正文的首引顺序改变，再运行首引排序以重排文末列表。[交叉引用操作](references/word-crossrefs.md)

## 直接使用 Word 脚本

下面以项目级安装为例，在包含 `原稿.docx` 的项目根目录运行。`python` 指已安装依赖的解释器，使用虚拟环境时替换为上一步的环境内 Python 路径。

```sh
python ".agents/skills/document-workflow/scripts/word_layout.py" apply "原稿.docx" --out "排版稿.docx"
python ".agents/skills/document-workflow/scripts/word_layout.py" check "排版稿.docx"
```

应用和核验论文预设时，在两条命令中加上同一个 `--profile academic-paper` 或 `--profile thesis`。

采用数字引文时，将文末条目放在「参考文献」标题下，并以 `[1]`、`[2]` 等编号起行。以下命令另存引用稿，建立跳转、按首引排序并给正文引文加上标。

```sh
python ".agents/skills/document-workflow/scripts/word_crossrefs.py" apply "排版稿.docx" --sort-bibliography --superscript-citations --out "引用稿.docx"
python ".agents/skills/document-workflow/scripts/word_crossrefs.py" check "引用稿.docx" --sort-bibliography --superscript-citations
```

`apply` 需要新的输出文件路径。`check` 会给出检查报告，交付前按报告复核跳过的内容、歧义或断链。版式预设和配置参数见 [Word 排版与校验](references/word-layout.md)。[SKILL.md](SKILL.md) 是 Codex 的工作流程入口。
