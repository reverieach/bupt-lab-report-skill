#!/usr/bin/env python3
"""Build a BUPT-styled DOCX from an arbitrary UTF-8 JSON report spec."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import (
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_TABLE_ALIGNMENT,
)
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


SKILL_DIR = Path(__file__).resolve().parent.parent
DEFAULTS = {
    "school": "北京邮电大学",
    "report_type": "课程实验报告",
    "major": "【待填写专业】",
    "class_name": "【待填写班级】",
    "student_name": "【待填写姓名】",
    "student_id": "【待填写学号】",
}


THEME_FONT_ATTRS = ("asciiTheme", "eastAsiaTheme", "hAnsiTheme", "cstheme")


def configure_rpr_font(r_pr, east_asia: str, *, latin: str | None = None) -> None:
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for attr in THEME_FONT_ATTRS:
        r_fonts.attrib.pop(qn(f"w:{attr}"), None)
    r_fonts.set(qn("w:ascii"), latin or east_asia)
    r_fonts.set(qn("w:hAnsi"), latin or east_asia)
    r_fonts.set(qn("w:eastAsia"), east_asia)
    r_fonts.set(qn("w:hint"), "eastAsia")
    lang = r_pr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        r_pr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")


def set_run_font(run, east_asia: str, *, latin: str | None = None) -> None:
    run.font.name = latin or east_asia
    r_pr = run._element.get_or_add_rPr()
    configure_rpr_font(r_pr, east_asia, latin=latin)


def set_style_font(style, east_asia: str, *, latin: str | None = None) -> None:
    style.font.name = latin or east_asia
    r_pr = style._element.get_or_add_rPr()
    configure_rpr_font(r_pr, east_asia, latin=latin)


def set_size_and_complex_size(run, points: float) -> None:
    run.font.size = Pt(points)
    r_pr = run._element.get_or_add_rPr()
    size_cs = r_pr.find(qn("w:szCs"))
    if size_cs is None:
        size_cs = OxmlElement("w:szCs")
        r_pr.append(size_cs)
    size_cs.set(qn("w:val"), str(int(round(points * 2))))


def set_bold_and_complex_bold(run, value: bool) -> None:
    run.bold = value
    r_pr = run._element.get_or_add_rPr()
    bold_cs = r_pr.find(qn("w:bCs"))
    if bold_cs is None:
        bold_cs = OxmlElement("w:bCs")
        r_pr.append(bold_cs)
    bold_cs.set(qn("w:val"), "1" if value else "0")


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def remove_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        element = borders.find(tag)
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "nil")


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cell_bottom_border(cell, *, size: int = 8) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    bottom = borders.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        borders.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), "000000")


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def set_cell_text(
    cell,
    text: str,
    *,
    bold: bool = False,
    size: float = 10.5,
    east_asia: str = "宋体",
    latin: str | None = None,
    alignment=WD_ALIGN_PARAGRAPH.CENTER,
) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = alignment
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(str(text))
    set_bold_and_complex_bold(run, bold)
    set_run_font(run, east_asia, latin=latin)
    set_size_and_complex_size(run, size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def configure_styles(doc: Document) -> None:
    styles = doc.styles
    defaults = styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        r_pr_default = defaults.find(qn("w:rPrDefault"))
        if r_pr_default is not None:
            default_r_pr = r_pr_default.find(qn("w:rPr"))
            if default_r_pr is not None:
                configure_rpr_font(default_r_pr, "宋体")
    normal = styles["Normal"]
    set_style_font(normal, "宋体")
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.first_line_indent = Cm(0.74)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

    for name, size, align, before, after in (
        ("Heading 1", 16, WD_ALIGN_PARAGRAPH.CENTER, 12, 8),
        ("Heading 2", 14, WD_ALIGN_PARAGRAPH.LEFT, 8, 6),
        ("Heading 3", 12, WD_ALIGN_PARAGRAPH.LEFT, 6, 4),
    ):
        style = styles[name]
        set_style_font(style, "黑体")
        style.font.size = Pt(size)
        style.font.bold = True
        r_pr = style._element.get_or_add_rPr()
        size_cs = r_pr.find(qn("w:szCs"))
        if size_cs is None:
            size_cs = OxmlElement("w:szCs")
            r_pr.append(size_cs)
        size_cs.set(qn("w:val"), str(int(size * 2)))
        if r_pr.find(qn("w:bCs")) is None:
            r_pr.append(OxmlElement("w:bCs"))
        style.font.color.rgb = RGBColor(0, 0, 0)
        pf = style.paragraph_format
        pf.alignment = align
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.keep_with_next = True

    for name in ("图题", "表题"):
        style = styles[name] if name in styles else styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        set_style_font(style, "宋体")
        style.font.size = Pt(10.5)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.first_line_indent = Cm(0)
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.keep_with_next = name == "表题"

    code = styles["代码"] if "代码" in styles else styles.add_style("代码", WD_STYLE_TYPE.PARAGRAPH)
    set_style_font(code, "等线", latin="Consolas")
    code.font.size = Pt(9.5)
    code.font.color.rgb = RGBColor(0, 0, 0)
    code.paragraph_format.first_line_indent = Cm(0)
    code.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    code.paragraph_format.space_before = Pt(4)
    code.paragraph_format.space_after = Pt(6)


def configure_section(section) -> None:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)
    section.header_distance = Cm(1.27)
    section.footer_distance = Cm(1.27)


def set_update_fields(doc: Document) -> None:
    settings = doc.settings._element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")


def add_field(paragraph, instruction: str, placeholder: str = "") -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend((begin, instr, separate))
    if placeholder:
        paragraph.add_run(placeholder)
    paragraph.add_run()._r.append(end)


def add_page_number(section, start: int = 1) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    add_field(p, " PAGE ", str(start))
    sect_pr = section._sectPr
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num)
    pg_num.set(qn("w:start"), str(start))


def add_cover(doc: Document, spec: dict[str, Any]) -> None:
    fields = [
        ("专业", spec["major"]),
        ("班级", spec["class_name"]),
        ("姓名", spec["student_name"]),
        ("学号", spec["student_id"]),
    ]
    for key, label in (("college", "学院"), ("instructor", "教师"), ("date", "日期")):
        if spec.get(key):
            fields.append((label, spec[key]))
    for item in spec.get("cover_fields", []):
        fields.append((str(item.get("label", "字段")), str(item.get("value", ""))))
    if len(fields) > 8:
        raise ValueError("封面字段超过 8 项，无法保证全部位于单页；请合并字段或移至正文")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(spec["school"])
    set_bold_and_complex_bold(r, True)
    set_size_and_complex_size(r, 36)
    set_run_font(r, "宋体")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(24)
    title = f'{spec.get("course_name", "").strip()}{spec["report_type"]}'.strip()
    report_title_size = 36 if len(title) <= 9 else 30 if len(title) <= 14 else 26
    r = p.add_run(title)
    set_bold_and_complex_bold(r, True)
    set_size_and_complex_size(r, report_title_size)
    set_run_font(r, "宋体")

    title_table = doc.add_table(rows=1, cols=2)
    title_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    title_table.autofit = False
    remove_table_borders(title_table)
    title_table.columns[0].width = Cm(3.3)
    title_table.columns[1].width = Cm(11.7)
    title_table.cell(0, 0).width = Cm(3.3)
    title_table.cell(0, 1).width = Cm(11.7)
    title_table.rows[0].height = Cm(1.25)
    title_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    prevent_row_split(title_table.rows[0])
    report_subject = str(spec.get("title", "【待填写题目】"))
    subject_size = 22 if len(report_subject) <= 12 else 18
    set_cell_text(
        title_table.cell(0, 0),
        "题目：",
        bold=True,
        size=22,
        east_asia="宋体",
        alignment=WD_ALIGN_PARAGRAPH.RIGHT,
    )
    set_cell_text(
        title_table.cell(0, 1),
        report_subject,
        bold=True,
        size=subject_size,
        east_asia="宋体",
    )
    set_cell_bottom_border(title_table.cell(0, 1))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(100 if len(fields) <= 4 else 80 if len(fields) <= 6 else 36)
    logo = SKILL_DIR / "assets" / "bupt-logo.png"
    if logo.exists():
        logo_width = 4.2 if len(fields) <= 6 else 3.7
        p.add_run().add_picture(str(logo), width=Cm(logo_width))
    else:
        r = p.add_run("【北邮校徽】")
        set_run_font(r, "宋体")

    info = doc.add_table(rows=len(fields), cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info.autofit = False
    remove_table_borders(info)
    info.columns[0].width = Cm(3.0)
    info.columns[1].width = Cm(7.5)
    info_font_size = 18 if len(fields) <= 6 else 16
    row_height = 0.92 if len(fields) <= 6 else 0.76
    for i, (label, value) in enumerate(fields):
        row = info.rows[i]
        row.height = Cm(row_height)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        prevent_row_split(row)
        info.cell(i, 0).width = Cm(3.0)
        info.cell(i, 1).width = Cm(7.5)
        set_cell_text(
            info.cell(i, 0),
            f"{label}：",
            bold=True,
            size=info_font_size,
            east_asia="华文仿宋",
            alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        )
        set_cell_text(
            info.cell(i, 1),
            value,
            bold=True,
            size=info_font_size,
            east_asia="华文仿宋",
        )
        set_cell_bottom_border(info.cell(i, 1))


def add_toc(doc: Document) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run("目 录")
    set_bold_and_complex_bold(r, True)
    set_size_and_complex_size(r, 16)
    set_run_font(r, "黑体")
    toc = doc.add_paragraph()
    toc.paragraph_format.first_line_indent = Cm(0)
    add_field(toc, ' TOC \\o "1-2" \\h \\z \\u ', "右键更新目录")
    doc.add_page_break()


def add_caption(doc: Document, kind: str, chapter_no: int, item_no: int, text: str) -> None:
    label = "图" if kind == "figure" else "表"
    style = "图题" if kind == "figure" else "表题"
    p = doc.add_paragraph(style=style)
    r = p.add_run(f"{label} {chapter_no}-{item_no} {text}".strip())
    set_run_font(r, "宋体")
    set_size_and_complex_size(r, 10.5)


def add_block(doc: Document, block: dict[str, Any], base_dir: Path, counters: dict[str, int], chapter_no: int) -> None:
    block_type = block.get("type")
    if block_type == "paragraph":
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0) if block.get("no_indent") else Cm(0.74)
        alignments = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                      "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
        p.alignment = alignments.get(block.get("align", "justify"), WD_ALIGN_PARAGRAPH.JUSTIFY)
        r = p.add_run(str(block.get("text", "")))
        set_run_font(r, "宋体")
        set_size_and_complex_size(r, 12)
    elif block_type in ("bullets", "numbered"):
        style = "List Bullet" if block_type == "bullets" else "List Number"
        for item in block.get("items", []):
            p = doc.add_paragraph(style=style)
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(str(item))
            set_run_font(r, "宋体")
            set_size_and_complex_size(r, 12)
    elif block_type == "image":
        path = Path(str(block.get("path", "")))
        if not path.is_absolute():
            path = base_dir / path
        if not path.exists():
            raise FileNotFoundError(f"Image not found: {path}")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(str(path), width=Cm(float(block.get("width_cm", 14.5))))
        if block.get("caption"):
            counters["figure"] += 1
            add_caption(doc, "figure", chapter_no, counters["figure"], str(block["caption"]))
    elif block_type == "table":
        if block.get("caption"):
            counters["table"] += 1
            add_caption(doc, "table", chapter_no, counters["table"], str(block["caption"]))
        headers = [str(x) for x in block.get("headers", [])]
        rows = [[str(x) for x in row] for row in block.get("rows", [])]
        col_count = len(headers) or (len(rows[0]) if rows else 1)
        table = doc.add_table(rows=(1 if headers else 0), cols=col_count)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        if headers:
            set_repeat_table_header(table.rows[0])
            for i, value in enumerate(headers):
                set_cell_text(table.cell(0, i), value, bold=True)
                set_cell_shading(table.cell(0, i), "E7E6E6")
        for values in rows:
            if len(values) != col_count:
                raise ValueError("All table rows must have the same number of cells")
            cells = table.add_row().cells
            for i, value in enumerate(values):
                set_cell_text(cells[i], value)
        doc.add_paragraph().paragraph_format.space_after = Pt(0)
    elif block_type == "code":
        if block.get("caption"):
            caption = doc.add_paragraph(style="表题")
            r = caption.add_run(str(block["caption"]))
            set_run_font(r, "宋体")
            set_size_and_complex_size(r, 10.5)
        p = doc.add_paragraph(style="代码")
        r = p.add_run(str(block.get("text", "")))
        set_run_font(r, "等线", latin="Consolas")
        set_size_and_complex_size(r, 9.5)
        p_pr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "F2F2F2")
        p_pr.append(shd)
    elif block_type == "formula":
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        r = p.add_run(str(block.get("text", "")))
        set_run_font(r, "宋体", latin="Cambria Math")
        if block.get("number"):
            r = p.add_run(f"    ({block['number']})")
            set_run_font(r, "宋体", latin="Cambria Math")
    elif block_type == "page_break":
        doc.add_page_break()
    elif block_type == "note":
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        r = p.add_run(f"【待补充】{block.get('text', '')}")
        set_bold_and_complex_bold(r, True)
        set_run_font(r, "宋体")
        set_size_and_complex_size(r, 12)
        r.font.color.rgb = RGBColor(192, 0, 0)
    else:
        raise ValueError(f"Unknown block type: {block_type!r}")


def add_content(doc: Document, spec: dict[str, Any], base_dir: Path) -> None:
    for chapter_index, chapter in enumerate(spec.get("chapters", []), start=1):
        if chapter_index > 1 and chapter.get("new_page", True):
            doc.add_page_break()
        p = doc.add_paragraph(style="Heading 1")
        r = p.add_run(str(chapter.get("title", f"第{chapter_index}章")))
        set_run_font(r, "黑体")
        set_size_and_complex_size(r, 16)
        set_bold_and_complex_bold(r, True)
        counters = {"figure": 0, "table": 0}
        for block in chapter.get("blocks", []):
            add_block(doc, block, base_dir, counters, chapter_index)
        for section in chapter.get("sections", []):
            level = int(section.get("level", 2))
            if level not in (2, 3):
                raise ValueError("Section level must be 2 or 3")
            p = doc.add_paragraph(style=f"Heading {level}")
            r = p.add_run(str(section.get("title", "章节标题")))
            set_run_font(r, "黑体")
            set_size_and_complex_size(r, 14 if level == 2 else 12)
            set_bold_and_complex_bold(r, True)
            for block in section.get("blocks", []):
                add_block(doc, block, base_dir, counters, chapter_index)


def load_spec(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        raw = json.load(f)
    spec = {**DEFAULTS, **raw}
    spec["course_name"] = spec.get("course_name") or ""
    if not isinstance(spec.get("chapters", []), list):
        raise ValueError("chapters must be a list")
    return spec


def build(input_path: Path, output_path: Path) -> None:
    spec = load_spec(input_path)
    doc = Document()
    configure_styles(doc)
    configure_section(doc.sections[0])
    set_update_fields(doc)
    add_cover(doc, spec)
    content_section = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section(content_section)
    add_page_number(content_section, 1)
    if spec.get("include_toc", True):
        add_toc(doc)
    add_content(doc, spec, input_path.parent)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True, help="UTF-8 JSON report specification")
    parser.add_argument("--output", type=Path, required=True, help="Output DOCX path")
    args = parser.parse_args()
    build(args.input.resolve(), args.output.resolve())
    print(args.output.resolve())


if __name__ == "__main__":
    main()
