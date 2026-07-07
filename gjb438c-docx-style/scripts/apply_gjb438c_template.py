#!/usr/bin/env python3
"""Apply the bundled GJB438C DOCX template to structured document content."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn


SKILL_DIR = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = SKILL_DIR / "assets" / "1-2总体技术方案-模板参考.docx"

HEADING_STYLE_BY_LEVEL = {
    1: "Heading 1",
    2: "Heading 2",
    3: "Heading 3",
    4: "Heading 4",
    5: "Heading 5",
}

REQUIRED_H1 = [
    "1 概述",
    "2 使用与技术指标要求",
    "3 总体架构",
    "4 分系统设计",
    "5 系统集成设计",
    "6 关键技术分析",
    "7 效能分析",
    "8 工程组织管理",
    "9 初步工作计划",
]

DEMO_H2 = {
    "1 概述": ["1.1 任务依据", "1.2 编制目的", "1.3 指导原则", "1.4 建设目标", "1.5 主要工作", "1.6 与其他项目关系", "1.7 名词术语", "1.8 缩略语"],
    "2 使用与技术指标要求": ["2.1 使用要求", "2.2 技术指标要求"],
    "3 总体架构": ["3.1 体系架构设计", "3.2 逻辑架构设计", "3.3 系统架构设计", "3.4 功能架构设计", "3.5 数据架构设计", "3.6 部署架构设计", "3.7 数据交换设计", "3.8 关键性能指标设计", "3.9 组织运用模式", "3.10 技术体制"],
    "4 分系统设计": ["4.1 一体化业务分系统"],
    "5 系统集成设计": ["5.1 集成联试"],
    "6 关键技术分析": ["6.1 关键技术一"],
    "7 效能分析": ["7.1 综合效能分析"],
    "8 工程组织管理": ["8.1 组织机构", "8.2 考核验收", "8.3 安全保密"],
    "9 初步工作计划": ["9.1 初步工作计划"],
}


def style_names(document: Document) -> set[str]:
    return {style.name for style in document.styles}


def require_styles(document: Document) -> None:
    required = {
        "Heading 1",
        "Heading 2",
        "Heading 3",
        "Heading 4",
        "Heading 5",
        "145正文",
        "Caption",
        "DB表头",
        "DB表正文",
        "311-目录标题",
        "toc 1",
        "toc 2",
        "toc 3",
        "编号密级",
    }
    missing = sorted(required - style_names(document))
    if missing:
        raise ValueError(f"Template is missing required styles: {', '.join(missing)}")


def clear_document_body(document: Document) -> None:
    """Keep template styles and section properties, then remove all example content."""
    body = document._body._element
    children = list(body)
    for child in children:
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def add_styled_paragraph(document: Document, text: str, style_name: str) -> None:
    paragraph = document.add_paragraph(style=style_name)
    paragraph.add_run(text)


def set_cell_text(cell: Any, text: Any, style_name: str) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.style = style_name
    paragraph.add_run("" if text is None else str(text))


def add_table(document: Document, table_data: dict[str, Any]) -> None:
    caption = table_data.get("caption")
    if caption:
        add_styled_paragraph(document, str(caption), "Caption")

    headers = [str(value) for value in table_data.get("headers", [])]
    rows = table_data.get("rows", [])
    if not headers:
        return

    table = document.add_table(rows=1, cols=len(headers))
    table.autofit = True
    for idx, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[idx], header, "DB表头")

    for row_values in rows:
        row = table.add_row()
        values = list(row_values)
        for idx in range(len(headers)):
            value = values[idx] if idx < len(values) else ""
            set_cell_text(row.cells[idx], value, "DB表正文")


def add_front_matter(document: Document, metadata: dict[str, Any]) -> None:
    title = str(metadata.get("title", "总体技术方案")).strip()
    subtitle = str(metadata.get("subtitle", "")).strip()
    add_styled_paragraph(document, "编号：", "编号密级")
    add_styled_paragraph(document, title, "145正文")
    if subtitle:
        add_styled_paragraph(document, subtitle, "145正文")
    add_styled_paragraph(document, "目录", "311-目录标题")
    add_styled_paragraph(document, "目录域请在 Word 或 WPS 中更新。", "145正文")


def normalize_sections(content: dict[str, Any]) -> list[dict[str, Any]]:
    sections = content.get("sections", [])
    if not isinstance(sections, list):
        raise ValueError("content JSON field 'sections' must be a list")
    return sections


def write_sections(document: Document, sections: list[dict[str, Any]]) -> None:
    for section in sections:
        level = int(section.get("level", 1))
        title = str(section.get("title", "")).strip()
        if not title:
            continue
        style_name = HEADING_STYLE_BY_LEVEL.get(level, "Heading 5")
        add_styled_paragraph(document, title, style_name)

        for paragraph in section.get("paragraphs", []) or []:
            if isinstance(paragraph, dict):
                text = str(paragraph.get("text", ""))
                role_style = str(paragraph.get("style", "145正文"))
                add_styled_paragraph(document, text, role_style)
            else:
                add_styled_paragraph(document, str(paragraph), "145正文")

        for table_data in section.get("tables", []) or []:
            add_table(document, table_data)


def demo_content() -> dict[str, Any]:
    sections: list[dict[str, Any]] = []
    for h1 in REQUIRED_H1:
        sections.append(
            {
                "level": 1,
                "title": h1,
                "paragraphs": ["本章为 GJB438C 模板格式样例内容，实际编写时应替换为项目真实资料。"],
            }
        )
        for h2 in DEMO_H2[h1]:
            section: dict[str, Any] = {
                "level": 2,
                "title": h2,
                "paragraphs": ["本节作为二级标题格式控制单元，正文统一采用模板中的 145正文 样式。"],
            }
            if h2.endswith("名词术语"):
                section["tables"] = [
                    {
                        "caption": "表 1 名词术语表",
                        "headers": ["编号", "名称", "说明"],
                        "rows": [["1", "示例术语", "用于验证表格样式。"]],
                    }
                ]
            elif h2.endswith("缩略语"):
                section["tables"] = [
                    {
                        "caption": "表 2 缩略语表",
                        "headers": ["序号", "简写", "全称", "解释说明"],
                        "rows": [["1", "GJB", "国家军用标准", "用于验证缩略语表格式。"]],
                    }
                ]
            sections.append(section)
    return {"metadata": {"title": "GJB438C 模板样例"}, "sections": sections}


def load_content(path: Path | None, use_demo: bool) -> dict[str, Any]:
    if use_demo:
        return demo_content()
    if not path:
        raise ValueError("Either --content-json or --demo is required")
    return json.loads(path.read_text(encoding="utf-8"))


def build_document(template: Path, content: dict[str, Any], output: Path) -> None:
    document = Document(template)
    require_styles(document)
    clear_document_body(document)
    add_front_matter(document, content.get("metadata", {}) or {})
    write_sections(document, normalize_sections(content))
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content-json", type=Path, help="Structured JSON content file.")
    parser.add_argument("--demo", action="store_true", help="Generate a minimal demo document.")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Template DOCX path.")
    parser.add_argument("--output", type=Path, required=True, help="Output DOCX path.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    content = load_content(args.content_json, args.demo)
    build_document(args.template, content, args.output)
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
