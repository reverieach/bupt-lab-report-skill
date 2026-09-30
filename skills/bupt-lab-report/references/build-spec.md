# DOCX 生成 JSON 规格

使用 UTF-8 JSON 调用：

```powershell
python scripts/build_report.py --input report.json --output report.docx
```

最小结构：

```json
{
  "course_name": "人工智能基础",
  "report_type": "课程实验报告",
  "title": "图像分类模型训练与评估",
  "major": "【待填写专业】",
  "class_name": "【待填写班级】",
  "student_name": "【待填写姓名】",
  "student_id": "【待填写学号】",
  "chapters": [
    {
      "title": "一、任务与数据",
      "new_page": true,
      "sections": [
        {
          "title": "1. 任务定义",
          "blocks": [
            {"type": "paragraph", "text": "……"},
            {"type": "bullets", "items": ["……", "……"]}
          ]
        }
      ]
    }
  ]
}
```

## 顶层字段

- `school`：默认“北京邮电大学”。
- `course_name`：课程名称。
- `report_type`：默认“课程实验报告”；可改为课程设计报告、项目报告等。
- `title`：报告或实验题目。
- `major`、`class_name`、`student_name`、`student_id`：默认是待填写占位符，正式报告中必须填写真实信息。
- `instructor`、`date`、`college`：可选；非空时加入封面。
- `cover_fields`：可选数组 `[{"label":"组员","value":"……"}]`，追加任意封面字段。
- `include_toc`：默认 `true`。
- `chapters`：任意顶级章节数组，不限制名称和数量。

## 章节与小节

- chapter：`title`、可选 `new_page`、可选 `blocks`、可选 `sections`。
- section：`title`、可选 `level`（2 或 3）、`blocks`。
- 标题文本由内容决定，脚本不会添加“实验原理”等固定名称。

## block 类型

- `paragraph`：`text`；可选 `align`（left/center/right/justify）、`no_indent`。
- `bullets`：`items` 字符串数组。
- `numbered`：`items` 字符串数组。
- `image`：`path`、可选 `caption`、可选 `width_cm`。相对路径以 JSON 所在目录为基准。
- `table`：可选 `caption`、`headers`、`rows`。
- `code`：`text`、可选 `language`、可选 `caption`。
- `formula`：`text`、可选 `number`；用于简单文本公式。复杂公式应在后处理中使用 Word 公式对象。
- `page_break`：无其他字段。
- `note`：`text`；生成醒目的待补说明，最终交付前应处理。

未知 block 类型会报错，避免静默丢失内容。

## 示例策略

先根据用户材料生成 JSON，再生成 DOCX。不要让 JSON 结构反过来限制内容：确有需要时可扩展脚本或在生成后的 DOCX 中继续编辑。
