#!/usr/bin/env python3
"""Validate key BUPT report layout properties and unresolved placeholders."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree


THEME_FONT_ATTRS = ("asciiTheme", "eastAsiaTheme", "hAnsiTheme", "cstheme")


def font_errors(r_pr, expected_font: str, label: str) -> list[str]:
    errors: list[str] = []
    if r_pr is None:
        return [f"{label}缺少显式字体属性"]
    r_fonts = r_pr.rFonts
    if r_fonts is None or r_fonts.get(qn("w:eastAsia")) != expected_font:
        errors.append(f"{label}东亚字体不是{expected_font}")
    if r_fonts is not None:
        for attr in THEME_FONT_ATTRS:
            if r_fonts.get(qn(f"w:{attr}")) is not None:
                errors.append(f"{label}仍保留主题字体属性 {attr}")
    lang = r_pr.find(qn("w:lang"))
    if lang is None or lang.get(qn("w:eastAsia")) != "zh-CN":
        errors.append(f"{label}缺少 zh-CN 东亚语言标记")
    return errors


def close(actual: float, expected: float, tolerance: float = 0.08) -> bool:
    return abs(actual - expected) <= tolerance


def validate(path: Path) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    doc = Document(path)

    if len(doc.sections) < 2:
        errors.append("封面与目录/正文未分节")

    for i, section in enumerate(doc.sections, 1):
        checks = {
            "纸张宽度": (section.page_width.cm, 21.0),
            "纸张高度": (section.page_height.cm, 29.7),
            "上边距": (section.top_margin.cm, 2.5),
            "下边距": (section.bottom_margin.cm, 2.5),
            "左边距": (section.left_margin.cm, 2.8),
            "右边距": (section.right_margin.cm, 2.6),
        }
        for label, (actual, expected) in checks.items():
            if not close(actual, expected):
                errors.append(f"第 {i} 节{label}为 {actual:.2f} cm，应为 {expected:.2f} cm")

    normal = doc.styles["Normal"]
    if normal.font.size is None or abs(normal.font.size.pt - 12) > 0.1:
        errors.append("正文 Normal 样式不是 12 pt")
    for name, expected in (("Heading 1", 16), ("Heading 2", 14)):
        size = doc.styles[name].font.size
        if size is None or abs(size.pt - expected) > 0.1:
            errors.append(f"{name} 不是 {expected} pt")

    errors.extend(font_errors(normal.element.rPr, "宋体", "Normal 样式"))
    for name in ("Heading 1", "Heading 2"):
        errors.extend(font_errors(doc.styles[name].element.rPr, "黑体", f"{name} 样式"))

    for p_idx, paragraph in enumerate(doc.paragraphs):
        if paragraph.style.name not in ("Heading 1", "Heading 2"):
            continue
        for r_idx, run in enumerate(paragraph.runs):
            if run.text.strip():
                errors.extend(font_errors(run._r.rPr, "黑体", f"标题段落 {p_idx} 运行 {r_idx}"))

    if len(doc.tables) < 2:
        errors.append("封面缺少题目表格或个人信息表格")
    else:
        cover_tables = doc.tables[:2]
        for t_idx, table in enumerate(cover_tables):
            if len(table.columns) < 2:
                errors.append(f"封面表格 {t_idx + 1} 缺少值列")
                continue
            for r_idx, row in enumerate(table.rows):
                tc_pr = row.cells[1]._tc.get_or_add_tcPr()
                borders = tc_pr.find(qn("w:tcBorders"))
                bottom = borders.find(qn("w:bottom")) if borders is not None else None
                if bottom is None or bottom.get(qn("w:val")) != "single":
                    errors.append(f"封面表格 {t_idx + 1} 第 {r_idx + 1} 行值字段缺少固定下划线")
        if len(doc.tables[1].rows) > 8:
            errors.append("封面信息字段超过 8 项，存在跨页风险")

    text = "\n".join(
        [p.text for p in doc.paragraphs]
        + [cell.text for table in doc.tables for row in table.rows for cell in row.cells]
    )
    for token in ("【待填写", "【待补充】", "右键更新目录"):
        if token in text:
            warnings.append(f"存在未处理占位内容：{token}")

    with ZipFile(path) as package:
        document_xml = package.read("word/document.xml")
        settings_xml = package.read("word/settings.xml")
        styles_xml = package.read("word/styles.xml")
        if b"MS Gothic" in document_xml or b"MS Gothic" in styles_xml:
            errors.append("文档包含 MS Gothic 字体引用")
        styles_root = etree.fromstring(styles_xml)
        theme_attrs = styles_root.xpath(
            ".//w:style[@w:styleId='Heading1' or @w:styleId='Heading2']//w:rFonts/@w:asciiTheme"
            " | .//w:style[@w:styleId='Heading1' or @w:styleId='Heading2']//w:rFonts/@w:eastAsiaTheme"
            " | .//w:style[@w:styleId='Heading1' or @w:styleId='Heading2']//w:rFonts/@w:hAnsiTheme"
            " | .//w:style[@w:styleId='Heading1' or @w:styleId='Heading2']//w:rFonts/@w:cstheme",
            namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
        )
        if theme_attrs:
            errors.append("标题样式 XML 仍含主题字体属性")
        if b"TOC" not in document_xml:
            warnings.append("未发现自动目录域")
        if b"updateFields" not in settings_xml:
            warnings.append("未启用打开文档时更新域")
        if b"pgNumType" not in document_xml:
            errors.append("未发现页码从 1 开始的分节设置")
        if not any(name.startswith("word/media/") for name in package.namelist()):
            warnings.append("文档中没有图片，封面校徽可能缺失")

    return errors, warnings


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()
    errors, warnings = validate(args.docx.resolve())
    for warning in warnings:
        print(f"WARNING: {warning}")
    for error in errors:
        print(f"ERROR: {error}")
    if errors:
        raise SystemExit(1)
    print("OK")


if __name__ == "__main__":
    main()
