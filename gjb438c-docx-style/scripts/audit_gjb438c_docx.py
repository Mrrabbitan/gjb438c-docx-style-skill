#!/usr/bin/env python3
"""Audit a DOCX or the skill bundle for GJB438C template compliance."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

import yaml
from docx import Document


REQUIRED_STYLES = {
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

REQUIRED_H1_NORMALIZED = [
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

STRICT_H2 = {
    "1": {"1.1 任务依据", "1.2 编制目的", "1.3 指导原则", "1.4 建设目标", "1.5 主要工作", "1.6 与其他项目关系", "1.7 名词术语", "1.8 缩略语"},
    "2": {"2.1 使用要求", "2.2 技术指标要求"},
    "3": {"3.1 体系架构设计", "3.2 逻辑架构设计", "3.3 系统架构设计", "3.4 功能架构设计", "3.5 数据架构设计", "3.6 部署架构设计", "3.7 数据交换设计", "3.8 关键性能指标设计", "3.9 组织运用模式", "3.10 技术体制"},
    "8": {"8.1 组织机构", "8.2 考核验收", "8.3 安全保密"},
}

STRICT_H2_NORMALIZED = {
    primary: {re.sub(r"^\d+\.\d+\s+", "", title) for title in titles}
    for primary, titles in STRICT_H2.items()
}

LEGACY_TERMS = ["ZBZQ", "XX数据中台", "联合XX数据资源体系", "XX数据融合治理"]
FORBIDDEN_LOCAL_PATHS = ["/" + "Users" + "/" + "anzp" + "/", "\\" + "Users" + "\\" + "anzp" + "\\"]

H1_RE = re.compile(r"^\d+\s+\S")
H2_RE = re.compile(r"^\d+\.\d+\s+\S")
H3_RE = re.compile(r"^\d+\.\d+\.\d+\s+\S")


def strip_heading_number(text: str) -> str:
    return re.sub(r"^\d+(?:\.\d+)*\s+", "", text).strip()


@dataclass
class AuditResult:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors

    def add_error(self, message: str) -> None:
        self.errors.append(message)

    def add_warning(self, message: str) -> None:
        self.warnings.append(message)


def paragraph_texts(document: Document) -> Iterable[str]:
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            yield text
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        yield text


def audit_docx(path: Path, strict_secondary: bool, strict_tables: bool) -> AuditResult:
    result = AuditResult()
    if not path.exists():
        result.add_error(f"DOCX does not exist: {path}")
        return result
    if not zipfile.is_zipfile(path):
        result.add_error(f"DOCX is not a valid zip package: {path}")
        return result

    document = Document(path)
    styles = {style.name for style in document.styles}
    missing_styles = sorted(REQUIRED_STYLES - styles)
    if missing_styles:
        result.add_error(f"Missing required styles: {', '.join(missing_styles)}")

    full_text = "\n".join(paragraph_texts(document))
    for term in LEGACY_TERMS:
        if term in full_text:
            result.add_error(f"Legacy template term remains in document: {term}")
    for token in FORBIDDEN_LOCAL_PATHS:
        if token in full_text:
            result.add_error(f"Local environment path remains in document: {token}")

    h1_seen: list[str] = []
    current_primary: str | None = None
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        style_name = paragraph.style.name if paragraph.style else ""
        if style_name in {"toc 1", "toc 2", "toc 3"}:
            continue
        if style_name == "Heading 1":
            normalized = strip_heading_number(text)
            h1_seen.append(normalized)
            if normalized in REQUIRED_H1_NORMALIZED:
                current_primary = str(REQUIRED_H1_NORMALIZED.index(normalized) + 1)
            else:
                result.add_error(f"First-level heading is outside template framework: {text}")
            continue
        if style_name == "Heading 2":
            normalized = strip_heading_number(text)
            if strict_secondary and current_primary in STRICT_H2_NORMALIZED:
                allowed = STRICT_H2_NORMALIZED[current_primary]
                if normalized not in allowed:
                    result.add_error(f"Second-level heading is outside template framework: {text}")
            continue
        if H3_RE.match(text) and style_name != "Heading 3":
            result.add_error(f"Third-level heading uses {style_name!r}, expected 'Heading 3': {text}")
        elif H2_RE.match(text):
            if style_name != "Heading 2":
                result.add_error(f"Second-level heading uses {style_name!r}, expected 'Heading 2': {text}")
            if strict_secondary:
                primary = text.split(".", 1)[0]
                allowed = STRICT_H2.get(primary)
                if allowed is not None and text not in allowed:
                    result.add_error(f"Second-level heading is outside template framework: {text}")
        elif H1_RE.match(text):
            if text in REQUIRED_H1:
                result.add_error(f"First-level heading uses {style_name!r}, expected 'Heading 1': {text}")

    if h1_seen != REQUIRED_H1_NORMALIZED:
        result.add_error(f"First-level chapter sequence mismatch. Expected {REQUIRED_H1_NORMALIZED}, got {h1_seen}")

    for table_index, table in enumerate(document.tables, start=1):
        for row_index, row in enumerate(table.rows):
            expected = "DB表头" if row_index == 0 else "DB表正文"
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text.strip()
                    if not text:
                        continue
                    style_name = paragraph.style.name if paragraph.style else ""
                    if style_name != expected:
                        message = f"Table {table_index} row {row_index + 1} cell paragraph uses {style_name!r}, expected {expected!r}: {text[:40]}"
                        if strict_tables:
                            result.add_error(message)
                        else:
                            result.add_warning(message)

    return result


def parse_frontmatter(skill_md: Path) -> dict[str, str]:
    text = skill_md.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        raise ValueError("SKILL.md must start with YAML frontmatter")
    end = text.find("\n---", 4)
    if end < 0:
        raise ValueError("SKILL.md frontmatter is not closed")
    metadata = yaml.safe_load(text[4:end]) or {}
    if not isinstance(metadata, dict):
        raise ValueError("SKILL.md frontmatter must be a mapping")
    return metadata


def iter_text_files(root: Path) -> Iterable[Path]:
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in {".md", ".py", ".yml", ".yaml", ".toml", ".txt"}:
            yield path


def validate_skill(skill_dir: Path) -> AuditResult:
    result = AuditResult()
    required_paths = [
        skill_dir / "SKILL.md",
        skill_dir / "agents" / "openai.yaml",
        skill_dir / "assets" / "1-2总体技术方案-模板参考.docx",
        skill_dir / "references" / "template-structure.md",
        skill_dir / "references" / "format-lock.md",
        skill_dir / "references" / "writing-prompts.md",
        skill_dir / "scripts" / "apply_gjb438c_template.py",
        skill_dir / "scripts" / "audit_gjb438c_docx.py",
    ]
    for path in required_paths:
        if not path.exists():
            result.add_error(f"Missing required skill file: {path}")

    if (skill_dir / "README.md").exists():
        result.add_error("Skill body must not contain README.md; keep repository docs at repo root.")

    skill_md = skill_dir / "SKILL.md"
    if skill_md.exists():
        try:
            metadata = parse_frontmatter(skill_md)
            if metadata.get("name") != skill_dir.name:
                result.add_error(f"Skill name {metadata.get('name')!r} does not match folder name {skill_dir.name!r}")
            if not metadata.get("description"):
                result.add_error("SKILL.md description is required")
        except Exception as exc:
            result.add_error(f"Invalid SKILL.md frontmatter: {exc}")

    template = skill_dir / "assets" / "1-2总体技术方案-模板参考.docx"
    if template.exists() and not zipfile.is_zipfile(template):
        result.add_error(f"Template asset is not a valid DOCX package: {template}")

    for path in iter_text_files(skill_dir):
        text = path.read_text(encoding="utf-8", errors="ignore")
        for token in FORBIDDEN_LOCAL_PATHS:
            if token in text:
                result.add_error(f"Local environment path appears in {path}: {token}")

    return result


def print_result(result: AuditResult, as_json: bool) -> None:
    if as_json:
        print(json.dumps({"ok": result.ok, "errors": result.errors, "warnings": result.warnings}, ensure_ascii=False, indent=2))
        return
    if result.errors:
        print("ERRORS:")
        for item in result.errors:
            print(f"- {item}")
    if result.warnings:
        print("WARNINGS:")
        for item in result.warnings:
            print(f"- {item}")
    if result.ok:
        print("OK")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", nargs="?", type=Path, help="DOCX file to audit.")
    parser.add_argument("--skill-dir", type=Path, help="Skill directory to validate.")
    parser.add_argument("--validate-skill", action="store_true", help="Validate the skill bundle structure.")
    parser.add_argument("--strict-secondary", action="store_true", help="Require fixed second-level headings where the template defines them.")
    parser.add_argument("--strict-tables", action="store_true", help="Treat table style mismatches as errors.")
    parser.add_argument("--json", action="store_true", help="Print machine-readable audit output.")
    return parser.parse_args()


def merge_results(results: list[AuditResult]) -> AuditResult:
    merged = AuditResult()
    for result in results:
        merged.errors.extend(result.errors)
        merged.warnings.extend(result.warnings)
    return merged


def main() -> int:
    args = parse_args()
    results: list[AuditResult] = []
    if args.validate_skill:
        if not args.skill_dir:
            raise SystemExit("--validate-skill requires --skill-dir")
        results.append(validate_skill(args.skill_dir))
    if args.docx:
        results.append(audit_docx(args.docx, args.strict_secondary, args.strict_tables))
    if not results:
        raise SystemExit("Provide a DOCX path and/or --validate-skill --skill-dir")
    result = merge_results(results)
    print_result(result, args.json)
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
