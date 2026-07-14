#!/usr/bin/env python3
"""Apply the bundled GJB438C DOCX template to structured document content.

Fidelity rules (measured from the bundled template):
- Headings use the template's multilevel auto-numbering; heading text must not
  contain typed number prefixes (they are stripped automatically).
- Table header cells use 145表头, body cells use 145表正文.
- Figures are a 145图样式 placeholder paragraph followed by a Caption paragraph.
- The template cover table, TOC field, and section break are preserved; only
  the body content (from the first Heading 1 onward) is replaced.
- The generated file is checked after save to ensure the template page-number
  fields, footer relationships, and two-section structure survived.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml.ns import qn

from audit_gjb438c_docx import AuditResult, audit_page_numbering


SKILL_DIR = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = SKILL_DIR / "assets" / "1-2总体技术方案-模板参考.docx"

HEADING_STYLE_BY_LEVEL = {
    1: "Heading 1",
    2: "Heading 2",
    3: "Heading 3",
    4: "Heading 4",
    5: "Heading 5",
}

TABLE_HEADER_STYLE = "145表头"
TABLE_BODY_STYLE = "145表正文"
FIGURE_STYLE = "145图样式"
GENERATED_TABLE_STYLE = "Table Grid1"

HEADING_NUMBER_RE = re.compile(r"^\d+(?:\.\d+)*[\s　]*")

REQUIRED_H1 = [
    "概述",
    "使用与技术指标要求",
    "总体架构",
    "分系统设计",
    "系统集成设计",
    "关键技术分析",
    "效能分析",
    "工程组织管理",
    "初步工作计划",
]

OVERALL_DOCUMENT_TYPE = "overall-technical-solution"
GJB_DOCUMENT_TYPES = {
    "SDP": ("软件开发计划", "A"),
    "SIP": ("软件安装计划", "B"),
    "STrP": ("软件移交计划", "C"),
    "STP": ("软件测试计划", "D"),
    "OCD": ("运行方案说明", "E"),
    "SSS": ("系统/子系统规格说明", "F"),
    "IRS": ("接口需求规格说明", "G"),
    "SSDD": ("系统/子系统设计说明", "H"),
    "IDD": ("接口设计说明", "I"),
    "SRS": ("软件需求规格说明", "J"),
    "SDD": ("软件设计说明", "K"),
    "DBDD": ("数据库设计说明", "L"),
    "STD": ("软件测试说明", "M"),
    "STR": ("软件测试报告", "N"),
    "SPS": ("软件产品规格说明", "O"),
    "SVD": ("软件版本说明", "P"),
    "SUM": ("软件用户手册", "Q"),
    "CPM": ("计算机编程手册", "R"),
    "FSM": ("固件保障手册", "S"),
    "SDSR": ("软件研制总结报告", "T"),
}

DEMO_H2 = {
    "概述": ["任务依据", "编制目的", "指导原则", "建设目标", "主要工作", "与其他项目关系", "名词术语", "缩略语"],
    "使用与技术指标要求": ["使用要求", "技术指标要求"],
    "总体架构": ["体系架构设计", "逻辑架构设计", "系统架构设计", "功能架构设计", "数据架构设计", "部署架构设计", "数据交换设计", "关键性能指标设计", "组织运用模式", "技术体制"],
    "分系统设计": ["一体化业务分系统"],
    "系统集成设计": ["集成联试"],
    "关键技术分析": ["关键技术一"],
    "效能分析": ["综合效能分析"],
    "工程组织管理": ["组织机构", "考核验收", "安全保密"],
    "初步工作计划": ["初步工作计划"],
}


def strip_heading_number(title: str) -> str:
    """Headings are auto-numbered by the template's multilevel list."""
    return HEADING_NUMBER_RE.sub("", title).strip()


def normalize_document_type(value: str | None) -> str:
    if not value:
        return OVERALL_DOCUMENT_TYPE
    cleaned = value.strip()
    if cleaned.lower() in {OVERALL_DOCUMENT_TYPE, "overall", "总体技术方案"}:
        return OVERALL_DOCUMENT_TYPE
    for code in GJB_DOCUMENT_TYPES:
        if cleaned.lower() == code.lower():
            return code
    supported = ", ".join([OVERALL_DOCUMENT_TYPE, *GJB_DOCUMENT_TYPES])
    raise ValueError(f"Unsupported document_type {value!r}. Supported: {supported}")


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
        "145图样式",
        "145表头",
        "145表正文",
        "Caption",
        "311-目录标题",
        "toc 1",
        "toc 2",
        "toc 3",
        "编号密级",
        "文件名称",
        "单位名称",
    }
    missing = sorted(required - style_names(document))
    if missing:
        raise ValueError(f"Template is missing required styles: {', '.join(missing)}")


def style_id_to_name_map(document: Document) -> dict[str, str]:
    return {style.style_id: style.name for style in document.styles if style.style_id}


def paragraph_style_name(p_element: Any, id_map: dict[str, str]) -> str:
    """Resolve the style name of a raw <w:p> element via its styleId."""
    ppr = p_element.find(qn("w:pPr"))
    if ppr is None:
        return "Normal"
    pstyle = ppr.find(qn("w:pStyle"))
    if pstyle is None:
        return "Normal"
    return id_map.get(pstyle.get(qn("w:val")), "Normal")


def clear_body_content(document: Document) -> None:
    """Remove body content from the first Heading 1 paragraph onward.

    Keeps the cover table, TOC title, TOC field, and the front-matter section
    break. The trailing body-level sectPr is always preserved.
    """
    body = document._body._element
    children = list(body)
    id_map = style_id_to_name_map(document)
    first_heading_index = None
    for index, child in enumerate(children):
        if child.tag != qn("w:p"):
            continue
        if paragraph_style_name(child, id_map) == "Heading 1":
            first_heading_index = index
            break

    if first_heading_index is None:
        # Fallback: template without body headings; clear everything but sectPr.
        first_heading_index = 0

    for child in children[first_heading_index:]:
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


def replace_paragraph_text(paragraph: Any, text: str) -> None:
    """Replace visible text while keeping the paragraph style and run format."""
    runs = paragraph.runs
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text)


def replace_first_nonempty_cell_paragraphs(cell: Any, texts: list[str]) -> None:
    """Replace the texts of the first N non-empty paragraphs in a cell."""
    remaining = list(texts)
    for paragraph in cell.paragraphs:
        if not remaining:
            break
        if paragraph.text.strip():
            replace_paragraph_text(paragraph, remaining.pop(0))


def update_cover_table(document: Document, metadata: dict[str, Any]) -> None:
    """Fill the template cover table in place, preserving cover styles."""
    if not document.tables:
        return
    cover = document.tables[0]

    number = str(metadata.get("number", "")).strip()
    identifier = str(metadata.get("identifier", "")).strip()
    title = str(metadata.get("title", "")).strip()
    unit = str(metadata.get("unit", "")).strip()
    date = str(metadata.get("date", "")).strip()

    for row in cover.rows:
        seen: set[int] = set()
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            cell_text = cell.text.strip()
            if number and "编号" in cell_text and len(cell_text) < 30:
                replace_first_nonempty_cell_paragraphs(cell, [f"编号：{number}"])
            elif identifier and "SJZT" in cell_text:
                replace_first_nonempty_cell_paragraphs(cell, [identifier])
            elif title and "总体技术方案" in cell_text:
                replace_first_nonempty_cell_paragraphs(cell, [title])
            elif "技术总师组" in cell_text and (unit or date):
                texts = [t for t in [unit, date] if t]
                replace_first_nonempty_cell_paragraphs(cell, texts)


def enable_update_fields(document: Document) -> None:
    """Ask Word/WPS to refresh the TOC field and page numbers on open."""
    settings = document.settings.element
    if settings.find(qn("w:updateFields")) is None:
        element = settings.makeelement(qn("w:updateFields"), {qn("w:val"): "true"})
        settings.append(element)


def require_page_number_structure(path: Path) -> None:
    """Fail fast if generated output loses template page-number structure."""
    result = AuditResult()
    audit_page_numbering(path, result)
    if result.errors:
        raise ValueError("Generated DOCX failed page-number structure audit: " + "; ".join(result.errors))


def add_table(document: Document, table_data: dict[str, Any]) -> None:
    caption = table_data.get("caption")
    if caption:
        add_styled_paragraph(document, str(caption), "Caption")

    headers = [str(value) for value in table_data.get("headers", [])]
    rows = table_data.get("rows", [])
    if not headers:
        return

    table = document.add_table(rows=1, cols=len(headers))
    try:
        table.style = GENERATED_TABLE_STYLE
    except KeyError:
        pass
    table.autofit = True
    for idx, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[idx], header, TABLE_HEADER_STYLE)

    for row_values in rows:
        row = table.add_row()
        values = list(row_values)
        for idx in range(len(headers)):
            value = values[idx] if idx < len(values) else ""
            set_cell_text(row.cells[idx], value, TABLE_BODY_STYLE)


def add_figure(document: Document, figure_data: dict[str, Any]) -> None:
    placeholder = str(figure_data.get("placeholder", "")).strip()
    add_styled_paragraph(document, placeholder, FIGURE_STYLE)
    caption = figure_data.get("caption")
    if caption:
        add_styled_paragraph(document, str(caption), "Caption")


def normalize_sections(content: dict[str, Any]) -> list[dict[str, Any]]:
    sections = content.get("sections", [])
    if not isinstance(sections, list):
        raise ValueError("content JSON field 'sections' must be a list")
    return sections


def write_sections(document: Document, sections: list[dict[str, Any]]) -> None:
    for section in sections:
        level = int(section.get("level", 1))
        title = strip_heading_number(str(section.get("title", "")))
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

        for figure_data in section.get("figures", []) or []:
            add_figure(document, figure_data)

        for table_data in section.get("tables", []) or []:
            add_table(document, table_data)


def overall_demo_content() -> dict[str, Any]:
    sections: list[dict[str, Any]] = []
    for chapter_index, h1 in enumerate(REQUIRED_H1, start=1):
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
            if h2 == "名词术语":
                section["tables"] = [
                    {
                        "caption": f"表{chapter_index}-1 名词术语",
                        "headers": ["编号", "名称", "说明"],
                        "rows": [["1", "示例术语", "用于验证表格样式。"]],
                    }
                ]
            elif h2 == "缩略语":
                section["tables"] = [
                    {
                        "caption": f"表{chapter_index}-2 缩略语",
                        "headers": ["序号", "简写", "全称", "解释说明"],
                        "rows": [["1", "GJB", "国家军用标准", "用于验证缩略语表格式。"]],
                    }
                ]
            elif h2 == "体系架构设计":
                section["figures"] = [{"caption": f"图 {chapter_index}-1 体系架构示意图"}]
            sections.append(section)
    return {"metadata": {"document_type": OVERALL_DOCUMENT_TYPE, "title": "总体技术方案"}, "sections": sections}


def standard_gjb_demo_content(document_type: str) -> dict[str, Any]:
    name, appendix = GJB_DOCUMENT_TYPES[document_type]
    sections = [
        {
            "level": 1,
            "title": "范围",
            "paragraphs": [
                f"本章说明{name}的适用范围、编制对象和使用边界。该样例用于验证标准 GJB 文档类型不会被强制套用总体技术方案九章结构。"
            ],
        },
        {
            "level": 1,
            "title": "引用文件",
            "paragraphs": ["本章列出编制本文件所依据的标准、合同、任务书和上级文件。"],
            "tables": [
                {
                    "caption": "表2-1 引用文件清单",
                    "headers": ["序号", "文件名称", "说明"],
                    "rows": [["1", "GJB 438C-2021", f"对应附录 {appendix} 的{name}格式。"]],
                }
            ],
        },
        {
            "level": 1,
            "title": f"{name}正文",
            "paragraphs": [
                "本章承载该文档类型的主体内容。真实生成时应按 GJB 438C-2021 对应附录逐条展开。"
            ],
        },
        {
            "level": 2,
            "title": "裁剪说明",
            "paragraphs": ["如项目生命周期或合同约定不适用某些条款，应标注本条无内容并说明理由。"],
        },
        {
            "level": 1,
            "title": "附录",
            "paragraphs": ["本章用于承载可独立维护的图表、清单、数据或补充说明。"],
        },
    ]
    return {
        "metadata": {
            "document_type": document_type,
            "title": name,
        },
        "sections": sections,
    }


def demo_content(document_type: str) -> dict[str, Any]:
    if document_type == OVERALL_DOCUMENT_TYPE:
        return overall_demo_content()
    return standard_gjb_demo_content(document_type)


def load_content(path: Path | None, use_demo: bool, document_type_override: str | None) -> dict[str, Any]:
    document_type = normalize_document_type(document_type_override)
    if use_demo:
        return demo_content(document_type)
    if not path:
        raise ValueError("Either --content-json or --demo is required")
    content = json.loads(path.read_text(encoding="utf-8"))
    metadata = content.setdefault("metadata", {})
    if document_type_override:
        metadata["document_type"] = document_type
    else:
        metadata["document_type"] = normalize_document_type(str(metadata.get("document_type", OVERALL_DOCUMENT_TYPE)))
    return content


def build_document(template: Path, content: dict[str, Any], output: Path) -> None:
    document = Document(template)
    require_styles(document)
    clear_body_content(document)
    update_cover_table(document, content.get("metadata", {}) or {})
    write_sections(document, normalize_sections(content))
    enable_update_fields(document)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    require_page_number_structure(output)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content-json", type=Path, help="Structured JSON content file.")
    parser.add_argument("--demo", action="store_true", help="Generate a minimal demo document.")
    parser.add_argument("--document-type", help="Document type: overall-technical-solution or one of SDP/SIP/STrP/STP/OCD/SSS/IRS/SSDD/IDD/SRS/SDD/DBDD/STD/STR/SPS/SVD/SUM/CPM/FSM/SDSR.")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Template DOCX path.")
    parser.add_argument("--output", type=Path, required=True, help="Output DOCX path.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    content = load_content(args.content_json, args.demo, args.document_type)
    build_document(args.template, content, args.output)
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
