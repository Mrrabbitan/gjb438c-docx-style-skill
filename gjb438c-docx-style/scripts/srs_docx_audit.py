"""Evidence-oriented Appendix J checks, independent of a particular Word layout.

These checks establish machine-checkable invariants, not semantic or visual
review. A supplied source model is checked independently and compared to the
saved document; missing external evidence is never inferred from trace tables.
"""
from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from pathlib import Path

from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

CHAPTERS = ["范围", "引用文档", "需求", "合格性规定", "需求可追踪性", "注释"]
CLAUSES = ["要求的状态和方式", "CSCI能力需求", "CSCI外部接口需求", "CSCI内部接口需求",
           "CSCI内部数据需求", "适应性需求", "保密性需求", "安全性需求", "CSCI环境适应性需求",
           "其他质量特性", "计算机资源需求", "设计和实现约束", "人员相关需求", "训练相关需求",
           "软件保障需求", "包装需求", "其他需求", "需求的优先顺序和关键性"]
RESOURCES = ["计算机硬件需求", "计算机硬件资源使用需求", "计算机软件需求", "计算机通信需求"]
EXPECTED = {str(i + 1): title for i, title in enumerate(CHAPTERS)}
EXPECTED.update({f"1.{i+1}": t for i, t in enumerate(["标识", "系统概述", "文档概述"])})
EXPECTED.update({f"3.{i+1}": t for i, t in enumerate(CLAUSES)})
EXPECTED.update({f"3.11.{i+1}": t for i, t in enumerate(RESOURCES)})
EXPECTED["3.3.1"] = "接口标识和接口图"
METHODS = {"演示", "测试", "分析", "审查", "特殊合格性方法", "特殊的合格性检验方法"}
TBD_RE = re.compile(r"\bTBD[-_][A-Za-z0-9_-]+", re.I)

# Independent display-contract mapping for model/OOXML reconciliation. Do not
# import the generator's mapping: a missing writer field must fail this audit.
REGISTER_LABELS = {
    "id": ("标识",), "name": ("名称",), "definition": ("定义方式", "定义"),
    "entry": ("进入条件",), "exit": ("退出条件",), "mode": ("方式",),
    "description": ("说明",), "transitions": ("状态转换",), "purpose": ("用途",),
    "parent_id": ("上级能力编号",), "type": ("类型",), "owner": ("责任方",),
    "classification": ("分级",), "lifecycle": ("生命周期",), "integrity": ("完整性",),
    "retention": ("保留要求",), "backup_recovery": ("备份恢复",), "migration": ("迁移",),
    "fields": ("数据元素",), "constraints": ("约束",), "csci_endpoint": ("CSCI端",),
    "external_endpoint": ("外部端",), "direction": ("方向",), "data": ("输入输出/数据",),
    "protocol": ("协议",), "authentication": ("鉴权",), "exception": ("异常处理",),
    "precondition": ("前置条件",), "evidence": ("计划证据", "验证证据"),
    "status": ("状态",), "timing": ("时序",), "capacity": ("容量",), "version": ("版本",),
    "security": ("保密性",), "kind": ("接口类别", "来源类别"), "irs_reference": ("受控IRS引用",),
    "statement": ("陈述",), "reference_id": ("引用文件编号",), "location": ("定位",),
    "allocation": ("分配范围",), "disposition": ("处置",), "reason": ("原因",),
    "number": ("文档编号",), "title": ("标题",), "organization": ("编写单位",),
    "revision": ("修订版",), "date": ("日期",), "source": ("获取来源",),
    "normative": ("规范性引用",), "issue": ("未决事项",), "impact": ("影响范围",),
    "closure_condition": ("关闭条件",), "closure_evidence": ("关闭证据",),
    "affected_ids": ("受影响对象", "影响对象"), "due": ("期限",),
    "overview": ("能力概述",), "order": ("同级顺序",), "physical_clause": ("物理章节",),
    "development_status": ("研制状态",), "development_basis": ("研制状态依据",),
    "confirmation": ("来源确认状态",), "tbd_ids": ("关联未决事项",),
}
REGISTER_IDS = {
    "states": ("标识", "状态编号"), "capabilities": ("标识", "能力编号"),
    "interfaces": ("接口编号",), "data": ("标识", "数据编号"),
    "references": ("文档编号",), "source_requirements": ("来源编号",), "tbd": ("TBD编号",),
}
REGISTER_CLAUSES = {"states": "3.1", "capabilities": "3.2", "data": "3.5", "references": "2", "source_requirements": "5", "tbd": "6"}


def compact(text):
    return re.sub(r"\s+", "", str(text or ""))


def tokens(text):
    return {v.strip() for v in re.split(r"[、,，;；\n]+", str(text or "")) if v.strip()}


def model_value_matches(expected, observed, key=""):
    """Compare composite values structurally, not by a global text search."""
    if observed is None:
        return False
    if isinstance(expected, (dict, list)):
        try:
            return json.loads(observed) == expected
        except (ValueError, TypeError):
            return isinstance(expected, list) and key == "affected_ids" and set(expected) == tokens(observed)
    if isinstance(expected, bool):
        return compact(observed).lower() in ({"是", "true", "1"} if expected else {"否", "false", "0"})
    return compact(expected) == compact(observed)


def relation_target(row):
    if row.get("external_target"):
        target = row["external_target"]
        return "IRS引用：" + target["reference_id"] + " / " + target["requirement_id"] + " / " + target["location"]
    return row.get("requirement_id", "")


def heading_level(paragraph):
    style = paragraph.style
    seen = set()
    direct = paragraph._p.find("./" + qn("w:pPr") + "/" + qn("w:outlineLvl"))
    if direct is not None:
        return int(direct.get(qn("w:val"), "9")) + 1
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        match = re.fullmatch(r"Heading ([1-9])", style.name or "", re.I)
        if match:
            return int(match.group(1))
        lvl = style.element.find("./" + qn("w:pPr") + "/" + qn("w:outlineLvl"))
        if lvl is not None:
            return int(lvl.get(qn("w:val"), "9")) + 1
        style = style.base_style
    return None


def effective_numbering(paragraph, document):
    """Resolve direct and inherited numPr, including partially overridden values."""
    props = []
    direct = paragraph._p.find(qn("w:pPr"))
    if direct is not None:
        props.append(direct)
    style, seen = paragraph.style, set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        ppr = style.element.find(qn("w:pPr"))
        if ppr is not None:
            props.append(ppr)
        style = style.base_style
    values = {}
    for ppr in props:
        numpr = ppr.find(qn("w:numPr"))
        if numpr is not None:
            for tag in ("numId", "ilvl"):
                node = numpr.find(qn("w:" + tag))
                if node is not None and tag not in values:
                    values[tag] = node.get(qn("w:val"))
    if not values.get("numId") or values["numId"] == "0":
        return None, "No effective automatic numbering"
    try:
        root = document.part.numbering_part.element
        nums = root.xpath('./w:num[@w:numId="%s"]' % int(values["numId"]))
        if not nums:
            return None, "Numbering ID has no definition"
        abstract_id = nums[0].find(qn("w:abstractNumId")).get(qn("w:val"))
        abstract = root.xpath('./w:abstractNum[@w:abstractNumId="%s"]' % int(abstract_id))
        level = int(values.get("ilvl", "0"))
        levels = [n for n in abstract[0] if n.tag == qn("w:lvl") and n.get(qn("w:ilvl")) == str(level)] if abstract else []
        if not levels:
            return None, "Numbering level has no definition"
        pattern = levels[0].find(qn("w:lvlText"))
        if pattern is None or not pattern.get(qn("w:val")):
            return None, "Numbering level has no displayed number"
        return (int(values["numId"]), level, pattern.get(qn("w:val"))), None
    except (AttributeError, IndexError, TypeError, ValueError, KeyError):
        return None, "Invalid numbering definition"


def document_blocks(document):
    """Retain paragraph/table order; table positions cannot be inferred separately."""
    counters = [0] * 9
    current, headings = "", {}
    for index, element in enumerate(document.element.body, start=1):
        if element.tag == qn("w:p"):
            obj = Paragraph(element, document)
            if obj.style and (obj.style.name or "").lower().startswith("toc "):
                continue
            level = heading_level(obj)
            if level and level <= 9 and obj.text.strip():
                counters[level-1] += 1
                counters[level:] = [0] * (9-level)
                logical = ".".join(str(n) for n in counters[:level])
                typed = re.match(r"^(\d+(?:\.\d+)*)[\s　]+(.+)$", obj.text.strip())
                title = typed.group(2) if typed else obj.text.strip()
                current = logical
                headings[current] = dict(title=title, level=level, paragraph=obj,
                                         typed=typed.group(1) if typed else None,
                                         location=f"word/document.xml:block[{index}]")
            yield "paragraph", obj, current, f"word/document.xml:block[{index}]", headings
        elif element.tag == qn("w:tbl"):
            yield "table", Table(element, document), current, f"word/document.xml:block[{index}]", headings


def table_records(table):
    rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
    if not rows:
        return [], set()
    if rows[0] == ["字段", "内容"]:
        record = {row[0]: row[1] for row in rows[1:] if len(row) >= 2 and row[0]}
        return [record], set(record)
    headers = rows[0]
    return [dict(zip(headers, row)) for row in rows[1:] if any(row)], set(headers)


def selected_profile(document, requested):
    if requested != "auto":
        return requested
    # A named style alone is not evidence of a selected profile. The generator
    # stamps the package; legacy generated SRS uses the distinctive full style set.
    styles = {s.name for s in document.styles}
    if {"SRS正文", "SRS表头", "SRS表正文", "SRS题注"}.issubset(styles):
        return "SRS-P09"
    return "standard"


def audit_srs(path: Path, document, result, strict=False, template_profile="auto", content=None):
    profile = selected_profile(document, template_profile)
    result.metrics["template_profile"] = profile
    result.checks.update(semantic_review="not_run", visual_review="not_run", source_coverage="not_run")

    def issue(rule, message, clause="", object_id="", location="", source="standard", warning=False):
        hard = rule.startswith("SRS.MODEL_") or rule in {
            "SRS.REQUIREMENT_ID", "SRS.INTERFACE_ID", "SRS.TBD_ID", "SRS.CLAUSE_PARENT",
            "SRS.NUMBERING_LEVEL", "SRS.DOUBLE_NUMBER", "SRS.TYPED_NUMBER",
            "SRS.TAILOR_CONFLICT", "SRS.QUALIFICATION_CONFLICT", "SRS.REVERSE_CONFLICT",
            "SRS.UNKNOWN_RELATION", "SRS.DUPLICATE_RELATION", "SRS.INTERFACE_GROUP"}
        result.add_issue(message, severity="warning" if warning or (not strict and not hard) else "error",
                         rule_id=rule, source=source, clause=clause, object_id=object_id,
                         location=location, status="needs_review" if source == "semantic" else "failed")

    model = None
    if content is not None:
        from srs_model import normalize_srs_content, validate_srs_content
        model = normalize_srs_content(content)
        issues = validate_srs_content(content, mode="review" if strict else "draft")
        for finding in issues:
            result.add_issue(**{k: v for k, v in finding.items() if k in {
                "message", "severity", "rule_id", "source", "clause", "object_id", "location", "status"}})
        result.checks["source_coverage"] = "failed" if any(
            i.get("severity") == "error" for i in issues) else (
                "passed" if "source_requirements" in model and not any(
                    i.get("rule_id", "").startswith(("SRS-SOURCE", "SRS-FORWARD", "SRS-REVERSE", "SRS-TRACE")) for i in issues)
                else "not_run")
    if result.checks["source_coverage"] == "not_run":
        issue("SRS.SOURCE_BASELINE", "Complete allocated-source coverage is not established; provide the independent source model with --content-json", "J.5", source="review")
    from srs_authoring import writing_profile as select_writing_profile
    training = select_writing_profile(model or {}, result.metrics.get("writing_profile")) == "training-review"
    external_labels = {relation_target(row) for key in ("qualification", "forwardTrace", "reverseTrace")
                       for row in (model or {}).get(key, []) if row.get("external_target")}

    blocks = list(document_blocks(document))
    headings = blocks[-1][-1] if blocks else {}
    direct_text = defaultdict(list)
    formal, formal_locations, formal_sections = {}, {}, {}
    qualifications, reverse, forward = defaultdict(list), defaultdict(list), []
    interfaces, references, tbds = {}, {}, {}
    saved_registers = {group: defaultdict(list) for group in REGISTER_IDS}
    tables = []
    chapter1 = [h["title"] for h in headings.values() if h["level"] == 1]
    if chapter1[:6] != CHAPTERS:
        issue("SRS.CHAPTER_ORDER", f"Appendix J chapter order mismatch: {chapter1}", "J")
    for kind, obj, clause, location, _ in blocks:
        if kind == "paragraph":
            if not heading_level(obj):
                direct_text[clause].append(obj.text)
            continue
        tables.append((obj, clause, location))
        records, fields = table_records(obj)
        for row_number, record in enumerate(records, start=2):
            loc = f"{location}:row[{row_number}]"
            direct_text[clause].append(" ".join(record.values()))
            register_group = None
            if "需求编号" not in fields and "SRS需求编号" not in fields:
                for label, group in (("TBD编号", "tbd"), ("接口编号", "interfaces"), ("文档编号", "references"), ("来源编号", "source_requirements")):
                    if label in fields:
                        register_group = group
                        break
                if register_group is None:
                    for group in ("states", "capabilities", "data"):
                        expected_section = REGISTER_CLAUSES[group]
                        if (clause == expected_section or clause.startswith(expected_section + ".")) and any(label in fields for label in REGISTER_IDS[group]):
                            register_group = group
                            break
            if register_group:
                identifier = next((record[label] for label in REGISTER_IDS[register_group] if label in record), "")
                saved_registers[register_group][identifier].append((record, clause, loc))
            if {"需求编号", "需求陈述"}.issubset(fields):
                rid = record.get("需求编号", "")
                if not rid or rid in formal:
                    issue("SRS.REQUIREMENT_ID", f"Empty or duplicate formal requirement ID: {rid!r}", "J.3", rid, loc)
                if rid in formal:
                    continue
                formal[rid], formal_locations[rid], formal_sections[rid] = record, loc, clause
                if not clause.startswith("3."):
                    issue("SRS.REQUIREMENT_LOCATION", f"Requirement {rid} is outside Chapter 3", "J.3", rid, loc)
                statement = record.get("需求陈述", "")
                if not statement:
                    issue("SRS.EMPTY_STATEMENT", f"Requirement {rid} has no statement", "J.3", rid, loc)
                normative = re.findall(r"(?<![响适对相顺感效])应(?:当|该)?|必须|不得|须", statement)
                if not normative or len(normative) > 1 or re.search(r"友好|高效|尽量|适当|必要时|快速地", statement):
                    issue("SRS.SEMANTIC_CANDIDATE", f"Review obligation, boundaries and atomicity in {rid}: {statement[:180]}", "J.3", rid, loc, "semantic", True)
            elif {"需求编号", "合格性方法"}.issubset(fields):
                rid = record.get("需求编号", "")
                qualifications[rid].append(record)
                if clause != "4" and not clause.startswith("4."):
                    issue("SRS.QUALIFICATION_LOCATION", f"Qualification matrix for {rid} is outside Chapter 4", "J.4", rid, loc)
            elif {"SRS需求编号", "来源编号或派生依据"}.issubset(fields):
                rid = record.get("SRS需求编号", "")
                reverse[rid].append(record)
                trace_chapter = "6" if training and model else "5"
                allowed_chapters = {trace_chapter, "6"} if training and not model else {trace_chapter}
                if not any(clause == c or clause.startswith(c + ".") for c in allowed_chapters):
                    issue("SRS.TRACE_LOCATION", "Reverse source trace matrix is outside its profile chapter " + trace_chapter, "J.5", rid, loc)
            elif {"来源编号", "SRS需求编号", "处置"}.issubset(fields):
                forward.append(record)
                trace_chapter = "6" if training and model else "5"
                allowed_chapters = {trace_chapter, "6"} if training and not model else {trace_chapter}
                if not any(clause == c or clause.startswith(c + ".") for c in allowed_chapters):
                    issue("SRS.TRACE_LOCATION", "Forward source trace matrix is outside its profile chapter " + trace_chapter, "J.5", location=loc)
            elif "接口编号" in fields:
                iid = record.get("接口编号", "")
                if not iid or iid in interfaces:
                    issue("SRS.INTERFACE_ID", f"Empty or duplicate interface ID: {iid}", "J.3.3", iid, loc)
                interfaces[iid] = (record, clause, loc)
                if not clause.startswith(("3.3.", "3.4")):
                    issue("SRS.INTERFACE_LOCATION", f"Interface {iid} is outside interface clauses", "J.3.3", iid, loc)
                for field in ("CSCI端", "外部端"):
                    if field in fields and not record.get(field):
                        issue("SRS.INTERFACE_PARTIES", f"Interface {iid} has no {field}", "J.3.3.1", iid, loc)
            elif {"文档编号", "标题", "修订版"}.issubset(fields):
                references[record.get("文档编号", "")] = record
                if clause != "2" and not clause.startswith("2."):
                    issue("SRS.REFERENCE_LOCATION", "Reference register is outside Chapter 2", "J.2", location=loc)
                for field in ("文档编号", "标题", "编写单位", "修订版", "日期"):
                    if not record.get(field):
                        issue("SRS.REFERENCE_METADATA", f"Reference is missing {field}", "J.2", location=loc)
            elif "TBD编号" in fields:
                tid = record.get("TBD编号", "")
                if not tid or tid in tbds:
                    issue("SRS.TBD_ID", f"Empty or duplicate TBD ID: {tid}", object_id=tid, location=loc, source="review")
                tbds[tid] = record
                for field in ("未决事项", "责任方", "关闭条件", "状态"):
                    if not record.get(field):
                        issue("SRS.TBD_METADATA", f"TBD {tid} is missing {field}", object_id=tid, location=loc, source="review")
                if not any(record.get(field) for field in ("影响范围", "受影响对象", "影响对象")):
                    issue("SRS.TBD_METADATA", f"TBD {tid} is missing impact scope or affected objects", object_id=tid, location=loc, source="review")
                if record.get("状态", "").lower() not in {"closed", "resolved", "已关闭", "已解决"}:
                    issue("SRS.OPEN_TBD", f"TBD {tid} remains open", object_id=tid, location=loc, source="review")

    def tailored(clause):
        text = " ".join(direct_text.get(clause, []))
        return bool(re.search(r"本[章条]无内容", text) and re.search(r"(?:原因|理由)\s*[:：]\s*\S+", text))

    for clause, expected in EXPECTED.items():
        if clause not in headings:
            ancestors = [".".join(clause.split(".")[:i]) for i in range(1, len(clause.split(".")))]
            if not any(tailored(parent) for parent in ancestors):
                issue("SRS.MISSING_CLAUSE", f"Missing clause {clause} {expected}; no explicit ancestor tailoring", "J." + clause)
        elif compact(headings[clause]["title"]) != compact(expected):
            issue("SRS.CLAUSE_PARENT", f"Clause {clause} must be {expected}, found {headings[clause]['title']}", "J." + clause, location=headings[clause]["location"])
    expected_by_title = {compact(title): clause for clause, title in EXPECTED.items()}
    for clause, heading in headings.items():
        location = heading["location"]
        expected_clause = expected_by_title.get(compact(heading["title"]))
        if expected_clause and expected_clause != clause:
            issue("SRS.CLAUSE_PARENT", f"{heading['title']} belongs to {expected_clause}, found at {clause}", "J." + expected_clause, location=location)
        numbering, error = effective_numbering(heading["paragraph"], document)
        if error and profile == "SRS-P09":
            issue("PROFILE.NUMBERING", error, clause, location=location, source="template")
        if numbering and numbering[1] != heading["level"]-1:
            issue("SRS.NUMBERING_LEVEL", f"Effective numbering level {numbering[1]} disagrees with heading level {heading['level']}", "4.3.7.2", location=location)
        if heading["typed"] and numbering:
            issue("SRS.DOUBLE_NUMBER", "Heading has both a typed prefix and automatic numbering", "4.3.7.2", location=location)
        if heading["typed"] and heading["typed"] != clause:
            issue("SRS.TYPED_NUMBER", f"Typed number {heading['typed']} does not match hierarchy {clause}", "4.3.7.2", location=location)
        text = " ".join(direct_text.get(clause, []))
        if "本条无内容" in text or "本章无内容" in text:
            if not tailored(clause):
                issue("SRS.TAILOR_REASON", "No-content statement lacks a concrete reason", "4.4", location=location)
            if any(c == clause or c.startswith(clause + ".") for c in formal_sections.values()):
                issue("SRS.TAILOR_CONFLICT", "Tailored clause contains formal requirements", "4.4", location=location)
        if re.search(r"3\.[23]\.X|\bVX\.X\b", heading["title"]):
            issue("SRS.HEADING_PLACEHOLDER", "Unresolved heading placeholder", "4.3.7.2", location=location)

    if not formal and not external_labels:
        issue("SRS.EMPTY_REQUIREMENTS", "No identifiable formal requirements were found", "J.3")
    if not references and not tailored("2"):
        issue("SRS.REFERENCE_REVIEW", "No structured reference register found; verify Chapter 2 manually", "J.2", source="semantic", warning=True)
    allowed_methods = set(METHODS)
    if model:
        for extension in model.get("qualification_method_extensions", []):
            allowed_methods.add(extension.get("name", extension.get("id", "")))
    for rid, record in formal.items():
        rows = qualifications.get(rid, [])
        row_methods = set().union(*(tokens(row.get("合格性方法")) for row in rows)) if rows else set()
        inline_methods = tokens(record.get("合格性方法"))
        if not row_methods and not inline_methods:
            issue("SRS.QUALIFICATION_MISSING", f"No qualification method is assigned to {rid}", "J.4", rid, formal_locations[rid])
        if row_methods and inline_methods and row_methods != inline_methods:
            issue("SRS.QUALIFICATION_CONFLICT", f"Inline and matrix qualification methods disagree for {rid}", "J.4", rid)
        if (row_methods | inline_methods) - allowed_methods:
            issue("SRS.METHOD_REVIEW", f"Check controlled definition for qualification method(s) of {rid}: {sorted((row_methods | inline_methods)-allowed_methods)}", "J.4", rid, source="semantic", warning=True)
        inline_sources = tokens(record.get("来源编号"))
        trace_sources = set().union(*(tokens(row.get("来源编号或派生依据")) for row in reverse[rid])) if reverse[rid] else set()
        if not inline_sources and not trace_sources:
            issue("SRS.REVERSE_MISSING", f"No source or derivation basis is recorded for {rid}", "J.5.a", rid)
        if inline_sources and trace_sources and inline_sources != trace_sources:
            issue("SRS.REVERSE_CONFLICT", f"Inline and reverse-trace sources disagree for {rid}", "J.5.a", rid)
    for kind, relations in (("qualification", qualifications), ("reverse trace", reverse)):
        for rid, rows in relations.items():
            if rid not in formal and rid not in external_labels:
                issue("SRS.UNKNOWN_RELATION", f"{kind} refers to unknown local requirement {rid}", "J.4" if kind == "qualification" else "J.5", rid)
            serialized = [tuple(sorted(row.items())) for row in rows]
            if len(serialized) != len(set(serialized)):
                issue("SRS.DUPLICATE_RELATION", f"Repeated identical {kind} relationship for {rid}", object_id=rid, source="review")
    for row in forward:
        target = row.get("SRS需求编号", "")
        # External IRS targets are validated against the controlled model below.
        if target not in formal and not model:
            issue("SRS.EXTERNAL_TARGET", f"Trace target {target} requires its controlled IRS/source model for validation", "J.5.b", target, source="review")

    for clause, parts in direct_text.items():
        for text in parts:
            ids = set(TBD_RE.findall(text))
            missing = ids - set(tbds)
            if missing:
                issue("SRS.TBD_UNREGISTERED", f"Unregistered TBD markers: {sorted(missing)}", clause, source="review")
            if "待确认" in text and not ids:
                # Signatures may legitimately remain blank; unresolved technical
                # facts must be identified, even if a register has a header row.
                if not (clause == "" and any(word in text for word in ("编制", "审核", "批准", "标审"))):
                    issue("SRS.UNRESOLVED_FACT", f"Unidentified unresolved fact: {text[:130]}", clause, source="review")

    if model:
        from srs_authoring import capability_outline
        try:
            outline = capability_outline(model)
        except ValueError:
            outline = []  # The model validator already reports malformed trees.
        capability_clauses = {r["node"]["id"]: r["physical_clause"] for r in outline}
        expected_reqs = {r["id"]: r for r in model.get("requirements", []) if isinstance(r, dict) and "id" in r}
        if set(expected_reqs) != set(formal):
            issue("SRS.MODEL_IDS", f"Saved DOCX/model requirement IDs differ: missing={sorted(set(expected_reqs)-set(formal))}, extra={sorted(set(formal)-set(expected_reqs))}", "J.3", source="implementation")
        for rid in set(formal) & set(expected_reqs):
            req = expected_reqs[rid]
            if compact(req.get("statement")) != compact(formal[rid].get("需求陈述")):
                issue("SRS.MODEL_STATEMENT", f"Saved statement differs from model for {rid}", "J.3", rid, formal_locations[rid], "implementation")
            for key, label in (("name", "名称"), ("applicability", "适用条件"), ("input", "输入"),
                               ("output", "预期输出"), ("boundary", "性能或边界"), ("exception", "异常与恢复"),
                               ("priority", "优先级"), ("criticality", "关键性"), ("evidence", "计划证据"),
                               ("capability_id", "能力编号"), ("capability", "能力名称"),
                               ("interface_id", "接口编号"), ("clause", "所属条款")):
                if key in req and compact(req[key]) != compact(formal[rid].get(label)):
                    issue("SRS.MODEL_ATTRIBUTE", f"Saved {key} differs from model for {rid}", "J.3", rid, formal_locations[rid], "implementation")
            if "derivation" in req and not model_value_matches(req["derivation"], formal[rid].get("派生依据")):
                issue("SRS.MODEL_ATTRIBUTE", f"Saved derivation basis/rationale differs from model for {rid}", "J.5.a", rid, formal_locations[rid], "implementation")
            for key, label in (("source_ids", "来源编号"), ("qualification_methods", "合格性方法"),
                               ("state_ids", "适用状态"), ("data_ids", "数据编号")):
                if key in req and set(req[key]) != tokens(formal[rid].get(label)):
                    issue("SRS.MODEL_ATTRIBUTE", f"Saved {key} differs from model for {rid}", "J.3", rid, formal_locations[rid], "implementation")
            for key, label in (("contract_trace", "合同映射处置"), ("exception_applicability", "异常适用性"), ("exception_basis", "异常适用依据"), ("physical_clause", "物理章节")):
                if key in req and not model_value_matches(req[key], formal[rid].get(label), key):
                    issue("SRS.MODEL_ATTRIBUTE", f"Saved {key} differs from model for {rid}", "J.3", rid, formal_locations[rid], "implementation")
            expected_clause = req.get("physical_clause") or capability_clauses.get(req.get("capability_id")) or str(req.get("clause", ""))
            actual_clause = formal_sections[rid]
            if actual_clause != expected_clause and not actual_clause.startswith(expected_clause + "."):
                issue("SRS.MODEL_LOCATION", f"Requirement {rid} belongs to {expected_clause}, found at {actual_clause}", "J.3", rid, formal_locations[rid], "implementation")
            if req.get("capability_id") in capability_clauses and actual_clause != capability_clauses[req["capability_id"]]:
                issue("SRS.MODEL_CAPABILITY_LOCATION", f"Requirement {rid} must be directly under {capability_clauses[req['capability_id']]}, found {actual_clause}", "J.3.2", rid, formal_locations[rid], "implementation")
            iid = req.get("interface_id")
            if iid and iid in interfaces and interfaces[iid][1] != actual_clause:
                issue("SRS.INTERFACE_GROUP", f"Requirement {rid} is not under its interface {iid}", "J.3.3", rid, formal_locations[rid], "implementation")
        for group, id_labels in REGISTER_IDS.items():
            id_key = "number" if group == "references" else "id"
            expected_records = {row[id_key]: row for row in model.get(group, [])}
            observed_records = saved_registers[group]
            if set(expected_records) != set(observed_records):
                issue("SRS.MODEL_REGISTER_IDS", f"Saved {group} IDs differ: missing={sorted(set(expected_records)-set(observed_records))}, extra={sorted(set(observed_records)-set(expected_records))}", source="implementation")
            for identifier in set(expected_records) & set(observed_records):
                observed = observed_records[identifier]
                if len(observed) != 1:
                    issue("SRS.MODEL_REGISTER_DUPLICATE", f"Saved {group} has {len(observed)} records for {identifier}", object_id=identifier, source="implementation")
                    continue
                actual, actual_clause, actual_location = observed[0]
                expected = expected_records[identifier]
                expected_clause = ("3.4" if expected.get("kind") == "internal" else "3.3") if group == "interfaces" else REGISTER_CLAUSES[group]
                if training and group == "source_requirements" and expected.get("kind") != "contract":
                    expected_clause = "6"
                if actual_clause != expected_clause and not actual_clause.startswith(expected_clause + "."):
                    issue("SRS.MODEL_REGISTER_LOCATION", f"Saved {group} record {identifier} belongs under {expected_clause}, found {actual_clause}", object_id=identifier, location=actual_location, source="implementation")
                if group == "capabilities" and identifier in capability_clauses and actual_clause != capability_clauses[identifier]:
                    issue("SRS.MODEL_CAPABILITY_LOCATION", f"Capability {identifier} must be at {capability_clauses[identifier]}, found {actual_clause}", "J.3.2", identifier, actual_location, "implementation")
                for key, value in expected.items():
                    labels = id_labels if key == id_key else REGISTER_LABELS.get(key, (key,))
                    label = next((name for name in labels if name in actual), None)
                    if label is None or not model_value_matches(value, actual[label], key):
                        issue("SRS.MODEL_REGISTER_VALUE", f"Saved {group}.{key} differs or is absent for {identifier}", object_id=identifier, location=actual_location, source="implementation")
        if select_writing_profile(model) == "training-review" and outline:
            expected_lists = {"3.2": {r["node"]["id"] for r in outline if r["level"] == 3}}
            for entry in outline:
                expected_lists[entry["physical_clause"]] = {r["node"]["id"] for r in outline if r["node"].get("parent_id") == entry["node"]["id"]} | {r["id"] for r in entry["requirements"]}
            actual_lists = defaultdict(list)
            for kind, obj, clause, location, _ in blocks:
                if kind == "table" and obj.rows:
                    headers = [c.text.strip() for c in obj.rows[0].cells]
                    if headers[:5] == ["序号", "名称", "唯一标识", "需求描述", "研制状态"]:
                        actual_lists[clause].append([row.cells[2].text.strip() for row in obj.rows[1:]])
            for clause, identifiers in expected_lists.items():
                actual = actual_lists.get(clause, [])
                if len(actual) != 1 or set(actual[0]) != identifiers or len(actual[0]) != len(identifiers):
                    issue("SRS.MODEL_FUNCTION_LIST", "功能清单缺失、重复或与该层直属能力/需求不一致", clause, source="implementation")
        expected_images = Counter()
        actual_images = Counter()
        paragraph_text = "\n".join(p.text for p in document.paragraphs)
        for kind, obj, clause, location, _ in blocks:
            if kind == "paragraph":
                actual_images[clause] += len(obj._p.xpath(".//w:drawing"))
        for figure in model.get("figures", []):
            locations = [capability_clauses[c] for c in figure.get("capability_ids", []) if c in capability_clauses]
            if figure.get("clause"):
                locations.append(figure["clause"])
            locations = list(dict.fromkeys(locations))
            for clause in locations:
                if figure.get("path"):
                    expected_images[clause] += 1
            if compact(figure["title"] + "（" + figure["id"] + "）") not in compact(paragraph_text):
                issue("SRS.MODEL_FIGURE", "Saved figure caption/ID missing: " + figure["id"], object_id=figure["id"], source="implementation")
        for clause, count in expected_images.items():
            if actual_images[clause] < count:
                issue("SRS.MODEL_FIGURE", f"Section {clause} contains {actual_images[clause]} drawings, expected at least {count}", clause, source="implementation")
        # Compare actual saved trace/qualification records to the independently
        # validated model, never regenerate a supposed source baseline here.
        actual_forward = {(r.get("来源编号", ""), r.get("SRS需求编号", "")) for r in forward}
        expected_forward = {(r.get("source", ""), relation_target(r)) for r in model.get("forwardTrace", [])}
        if actual_forward != expected_forward:
            issue("SRS.MODEL_FORWARD", f"Saved forward mappings differ: missing={sorted(expected_forward-actual_forward)}, extra={sorted(actual_forward-expected_forward)}", "J.5.b", source="implementation")
        expected_reverse = {(relation_target(r), s) for r in model.get("reverseTrace", [])
                            for s in (r.get("source_ids") or [r.get("source") or r.get("derivation_basis", "")])}
        # A derivation basis is prose/controlled decision citation, and may
        # legitimately contain semicolons. Do not split it as a list of IDs.
        derivation_bases = defaultdict(dict)
        for row in model.get("reverseTrace", []):
            if row.get("derivation_basis") and not row.get("source_ids") and not row.get("source"):
                derivation_bases[relation_target(row)][compact(row["derivation_basis"])] = row["derivation_basis"]
        actual_reverse = set()
        for rid, rows in reverse.items():
            for row in rows:
                text = row.get("来源编号或派生依据", "")
                basis = derivation_bases[rid].get(compact(text))
                for source in ([basis] if basis is not None else tokens(text)):
                    actual_reverse.add((rid, source))
        if expected_reverse != actual_reverse:
            issue("SRS.MODEL_REVERSE", "Saved reverse mappings differ from the validated model", "J.5.a", source="implementation")
        if training:
            from srs_authoring import contract_trace_tables
            for expected_table in contract_trace_tables(model):
                headers = expected_table["headers"]
                matching_tables = [(table, clause) for table, clause, _ in tables
                                   if table.rows and [c.text.strip() for c in table.rows[0].cells] == headers]
                expected_rows = Counter(tuple(compact(value) for value in row) for row in expected_table["rows"])
                actual_rows = Counter(tuple(compact(cell.text) for cell in row.cells) for table, _ in matching_tables for row in table.rows[1:])
                if len(matching_tables) != 1 or any(clause != "5" and not clause.startswith("5.") for _, clause in matching_tables) or actual_rows != expected_rows:
                    issue("SRS.MODEL_CONTRACT_TRACE", "合同正式追踪表缺失、位置错误或与独立合同来源/映射处置不一致", "J.5", source="implementation")
        expected_q_targets = {relation_target(r) for r in model.get("qualification", [])}
        if expected_q_targets != set(qualifications):
            issue("SRS.MODEL_QUALIFICATION", "Saved qualification targets differ from the validated model", "J.4", source="implementation")
        expected_qualification = Counter((relation_target(row), tuple(sorted(set(row.get("methods", [])))), compact(row.get("condition")), compact(row.get("evidence")), compact(row.get("pass_criterion"))) for row in model.get("qualification", []))
        actual_qualification = Counter((rid, tuple(sorted(tokens(row.get("合格性方法")))), compact(row.get("验证条件")), compact(row.get("计划证据")), compact(row.get("通过准则"))) for rid, rows in qualifications.items() for row in rows)
        if expected_qualification != actual_qualification:
            issue("SRS.MODEL_QUALIFICATION", "Saved qualification rows differ from the model: target, methods, condition, evidence and criterion must remain paired", "J.4", source="implementation")

    if profile == "SRS-P09":
        from audit_gjb438c_docx import (table_has_repeat_header, row_prevents_split,
            table_has_visible_borders, table_has_fixed_layout, table_geometry_issue)
        for table, clause, loc in tables:
            if not table_has_repeat_header(table):
                issue("PROFILE.TABLE_HEADER", "Table header is not repeated", clause, location=loc, source="template")
            if not table_has_visible_borders(table) or not table_has_fixed_layout(table):
                issue("PROFILE.TABLE_LAYOUT", "P-09 table needs defined borders and fixed layout", clause, location=loc, source="template")
            if clause and table_geometry_issue(table):
                issue("PROFILE.TABLE_GEOMETRY", table_geometry_issue(table), clause, location=loc, source="template")
            for i, row in enumerate(table.rows[1:], start=2):
                if not row_prevents_split(row):
                    issue("PROFILE.TABLE_SPLIT", "Logical record can split across pages", clause, location=f"{loc}:row[{i}]", source="template")
    result.metrics.update(formal_requirement_count=len(formal), interface_count=len(interfaces),
                          reference_count=len(references), tbd_count=len(tbds),
                          srs_h1_count=len(chapter1), visual_review_required=True)
    result.checks["machine"] = "passed" if result.ok else "failed"
