# document-workflow

适用于大学生文章写作和 Word 排版的 Skill，可用于课程作业、读书报告和论文。
它可以统一字体、行距、段落缩进和标题编号，也支持按本次要求调整版式。
日常文章沿用标题和正文，论文专用部件按写作任务选用。

## 使用

在 Codex 中，将本仓库保存为 `~/.agents/skills/document-workflow`，用 `$document-workflow` 调用。

> 用 $document-workflow 排版这份课程报告，统一标题和正文格式，保留原文并另存为 Word。

运行 Word 排版与交叉引用脚本需要 Python 3.10+ 和 `python-docx` 1.2，具体用法见 [SKILL.md](SKILL.md)。
