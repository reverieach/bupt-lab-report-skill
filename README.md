# 北邮课程实验报告 Skill

这是一个供北京邮电大学学生使用的**非官方**课程报告 Skill。它能依据课程要求、教师模板和实验材料，整理内容并生成可编辑的 DOCX。版式来自课程报告样例，不代表学校统一规范；教师要求始终优先。

## 安装

在 Codex 中让 `$skill-installer` 从本仓库安装 `skills/bupt-lab-report`，或将该目录完整复制到自己的 `~/.agents/skills/`。例如，可以把仓库链接发给 Codex 并说：

> 请用 `$skill-installer` 安装这个 GitHub 仓库中的 `skills/bupt-lab-report` Skill。

安装后可在新对话中使用 `$bupt-lab-report`，并附上课程要求、实验材料和自己的封面信息。普通网页版 ChatGPT 的本地文件访问和 Skill 安装方式可能不同；仓库链接本身不会自动安装 Skill。

## 独立使用生成脚本

在仓库根目录运行：

```bash
python -m pip install -r requirements.txt
python skills/bupt-lab-report/scripts/build_report.py --input examples/demo-report.json --output demo-report.docx
python skills/bupt-lab-report/scripts/validate_report.py demo-report.docx
```

`examples/demo-report.json` 展示输入格式。正式使用时，请把专业、班级、姓名、学号等占位符改为自己的信息，并在 Word 中更新目录和页码域。详见 [生成规格](skills/bupt-lab-report/references/build-spec.md)。

仓库不包含原始学生报告、个人信息或其预览图；内置 DOCX 只是空白模板。

