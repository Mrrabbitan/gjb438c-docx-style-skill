"""Versioned SRS authoring-model validation, independent of Word layout.

These are project automation conventions. Each finding identifies whether its
basis is the input schema, the standard's content obligations, or a review rule.
No normalization step invents project facts, approvals, or allocated sources.
"""
from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path
import re
from typing import Any

from jsonschema import Draft202012Validator

SCHEMA_PATH = Path(__file__).resolve().parents[1] / "references" / "srs-content-schema.json"
SRS_METHODS = frozenset({"演示", "测试", "分析", "审查", "特殊合格性方法"})
SRS_CLAUSES = tuple([f"3.{n}" for n in range(1, 19)] + [f"3.11.{n}" for n in range(1, 5)])
CLAUSE_TITLES = {
    "3.1": "要求的状态和方式", "3.2": "CSCI能力需求", "3.3": "CSCI外部接口需求",
    "3.4": "CSCI内部接口需求", "3.5": "CSCI内部数据需求", "3.6": "适应性需求",
    "3.7": "保密性需求", "3.8": "安全性需求", "3.9": "CSCI环境适应性需求",
    "3.10": "其他质量特性", "3.11": "计算机资源需求", "3.11.1": "计算机硬件需求",
    "3.11.2": "计算机硬件资源使用需求", "3.11.3": "计算机软件需求", "3.11.4": "计算机通信需求",
    "3.12": "设计和实现约束", "3.13": "人员相关需求", "3.14": "训练相关需求",
    "3.15": "软件保障需求", "3.16": "包装需求", "3.17": "其他需求", "3.18": "需求的优先顺序和关键性",
}


def _methods(value: Any) -> Any:
    if isinstance(value, str):
        return [v.strip() for v in re.split(r"[、,，;；/]", value) if v.strip()]
    return value


def normalize_srs_content(content: Any) -> Any:
    """Copy and migrate documented v1 aliases without inferring missing facts.

    Migration is deliberately limited to unversioned/v1 content. A v2 document
    with malformed types must fail schema validation rather than be repaired.
    """
    value = deepcopy(content)
    if not isinstance(value, dict):
        return value
    legacy = value.get("schema_version") in (None, "1.0", "1")
    if not legacy:
        return value
    value["schema_version"] = "2.0"
    for req in value.get("requirements", []) if isinstance(value.get("requirements", []), list) else []:
        if not isinstance(req, dict):
            continue
        if req.get("source") and "source_ids" not in req:
            req["source_ids"] = [req["source"]]
        if "qualification_methods" in req:
            req["qualification_methods"] = _methods(req["qualification_methods"])
        # A legacy capability is a supplied project name, not a generated fact.
        # Keep it as a grouping label; do not invent a controlled capability ID.
    for key in ("qualification", "reverseTrace"):
        for row in value.get(key, []) if isinstance(value.get(key, []), list) else []:
            if isinstance(row, dict) and "methods" in row:
                row["methods"] = _methods(row["methods"])
    return value


def _pointer(parts: Any) -> str:
    return "/" + "/".join(str(p).replace("~", "~0").replace("/", "~1") for p in parts)


def external_target_key(target: dict[str, str]) -> str:
    """Stable internal relation key; never confuse it with a local requirement ID."""
    return "IRS:" + json.dumps([target["reference_id"], target["requirement_id"], target["location"]], ensure_ascii=False, separators=(",", ":"))


def validate_srs_content(content: Any, mode: str = "draft", writing_profile: str | None = None) -> list[dict[str, Any]]:
    """Return located findings. Draft tolerates incompleteness, never bad types.

    Callers may pass legacy content; it is normalized on a private copy. Passing
    this validation does not certify semantic correctness or rendered layout.
    """
    if mode not in {"draft", "review"}:
        raise ValueError("SRS mode must be 'draft' or 'review'")
    model = normalize_srs_content(content)
    issues: list[dict[str, Any]] = []

    def add(rule: str, message: str, location: str = "/", *, object_id: str = "", clause: str = "", source: str = "project-model", hard: bool = False, advisory: bool = False) -> None:
        issues.append({"rule_id": rule, "severity": "error" if hard or (mode == "review" and not advisory) else "warning", "source": source, "clause": clause, "object_id": object_id, "location": location, "message": message})

    try:
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        Draft202012Validator.check_schema(schema)
    except Exception as exc:
        add("SRS-SCHEMA-LOAD", f"无法载入或验证 SRS Schema：{exc}", hard=True)
        return issues
    errors = sorted(Draft202012Validator(schema).iter_errors(model), key=lambda e: str(list(e.absolute_path)))
    for error in errors:
        add("SRS-SCHEMA", error.message, _pointer(error.absolute_path), source="schema", hard=True)
    if errors:
        return issues
    # Non-empty strings are controlled identifiers; whitespace cannot be an ID.
    def registry(key: str, id_key: str = "id") -> dict[str, dict[str, Any]]:
        result: dict[str, dict[str, Any]] = {}
        for i, row in enumerate(model.get(key, [])):
            identifier = row.get(id_key, "")
            if not identifier.strip():
                add("SRS-ID-EMPTY", f"{key} 标识不能为空白", f"/{key}/{i}/{id_key}", hard=True)
            elif identifier != identifier.strip():
                add("SRS-ID-WHITESPACE", "受控标识不应带首尾空白", f"/{key}/{i}/{id_key}", object_id=identifier, hard=True)
            if identifier in result:
                add("SRS-ID-DUPLICATE", f"{key} 存在重复标识 {identifier}", f"/{key}/{i}/{id_key}", object_id=identifier, hard=True)
            result[identifier] = row
        return result

    requirements = registry("requirements")
    references = registry("references", "number")
    sources = registry("source_requirements")
    states = registry("states")
    capabilities = registry("capabilities")
    interfaces = registry("interfaces")
    data = registry("data")
    tbds = registry("tbd")
    tailoring = registry("tailoring", "clause")
    extensions = registry("qualification_method_extensions", "name")
    allowed_methods = set(SRS_METHODS) | set(extensions)
    for name, row in extensions.items():
        if name in SRS_METHODS:
            add("SRS-METHOD-REDEFINE", "扩展不得重定义基本合格性方法", "/qualification_method_extensions", object_id=name, hard=True)
        if not row.get("definition", "").strip() or not row.get("approval", "").strip():
            add("SRS-METHOD-EXTENSION", "项目扩展方法需要定义和受控批准依据", "/qualification_method_extensions", object_id=name)

    if model.get("sections"):
        add("SRS-LOWLEVEL-REVIEW", "sections 为低层装配输入；其内容不能由模型校验认证。评审模式必须使用受控模型生成并独立审核最终 DOCX", "/sections")
    if not requirements and not any(row.get("external_target") for row in model.get("forwardTrace", [])):
        add("SRS-REQUIREMENTS-EMPTY", "尚未提供正式需求；只能生成草稿结构", "/requirements", source="GJB-J.3")
    metadata = model["metadata"]
    if not metadata["title"].strip():
        add("SRS-TITLE-EMPTY", "文档标题不能为空白", "/metadata/title", hard=True)
    for key in ("csci_id", "version", "system_overview", "document_overview"):
        if not metadata.get(key, "").strip():
            add("SRS-SCOPE-MISSING", f"尚未提供 {key}", f"/metadata/{key}", clause="1", source="GJB-J.1")
    if not references:
        add("SRS-REFERENCES-EMPTY", "尚未提供引用文件登记", "/references", clause="2", source="GJB-J.2")
    for i, ref in enumerate(model.get("references", [])):
        for key in ("organization", "revision", "date"):
            if not ref.get(key, "").strip():
                add("SRS-REFERENCE-METADATA", f"引用文件缺少 {key}", f"/references/{i}/{key}", object_id=ref["number"], clause="2", source="GJB-J.2")

    def methods_check(methods: list[str], location: str, rid: str) -> None:
        for method in methods:
            if method not in allowed_methods:
                add("SRS-METHOD-UNKNOWN", f"未受控的合格性方法：{method}", location, object_id=rid, clause="4", hard=True)
        if not methods:
            add("SRS-METHOD-MISSING", "正式需求尚未分配合格性方法", location, object_id=rid, clause="4", source="GJB-J.4")

    def target_check(identifier: str, target: dict[str, Any], location: str, owner: str, rule: str = "SRS-LINK-UNKNOWN") -> None:
        if identifier not in target:
            add(rule, f"引用的标识不存在：{identifier}", location, object_id=owner, hard=True)

    for i, state in enumerate(model.get("states", [])):
        if not state.get("definition", "").strip():
            add("SRS-STATE-DEFINITION", "状态/方式尚缺定义", f"/states/{i}", object_id=state["id"], clause="3.1")
        for transition in state.get("transitions", []):
            target_check(transition["to"], states, f"/states/{i}/transitions", state["id"])
    for i, cap in enumerate(model.get("capabilities", [])):
        if cap.get("parent_id"):
            target_check(cap["parent_id"], capabilities, f"/capabilities/{i}/parent_id", cap["id"])
        seen = {cap["id"]}
        parent = cap.get("parent_id")
        while parent in capabilities:
            if parent in seen:
                add("SRS-CAPABILITY-CYCLE", "能力层级存在环", f"/capabilities/{i}", object_id=cap["id"], hard=True)
                break
            seen.add(parent)
            parent = capabilities[parent].get("parent_id")

    for i, interface in enumerate(model.get("interfaces", [])):
        iid = interface["id"]
        clause = "3.4" if interface.get("kind") == "internal" else "3.3"
        if interface.get("definition") == "irs_reference":
            controlled = interface.get("irs_reference")
            if not controlled:
                add("SRS-IRS-MISSING", "受控 IRS 方式需要文件编号和条款位置", f"/interfaces/{i}/irs_reference", object_id=iid, clause=clause)
            else:
                target_check(controlled["reference_id"], references, f"/interfaces/{i}/irs_reference/reference_id", iid)
        else:
            for key in ("csci_endpoint", "external_endpoint", "direction", "data", "protocol"):
                if not interface.get(key):
                    add("SRS-INTERFACE-INCOMPLETE", f"接口缺少 {key} 或明确适用性说明", f"/interfaces/{i}/{key}", object_id=iid, clause=clause, source="GJB-J.3.3")
    external_inline = [row for row in interfaces.values() if row.get("kind", "external") == "external" and row.get("definition", "inline") == "inline"]
    if external_inline and not model.get("interface_diagram"):
        add("SRS-INTERFACE-DIAGRAM", "外部接口标识与关系尚缺接口图或其受控引用", "/interface_diagram", clause="3.3.1", source="GJB-J.3.3.1")
    if isinstance(model.get("interface_diagram"), dict) and model["interface_diagram"].get("reference_id"):
        target_check(model["interface_diagram"]["reference_id"], references, "/interface_diagram/reference_id", "接口关系图")

    policy = model.get("priority_policy", {})
    expected_pairs: set[tuple[str, str]] = set()
    for i, req in enumerate(model["requirements"]):
        rid, clause = req["id"], req["clause"]
        loc = f"/requirements/{i}"
        if not req["statement"].strip():
            add("SRS-STATEMENT-EMPTY", "需求陈述不能为空白", loc + "/statement", object_id=rid, hard=True)
        elif not re.search(r"(?<![响适对供])应(?![用急答力聘酬变接付援])|必须|不得|须|\bshall\b|\bmust\b", req["statement"], re.I):
            add("SRS-QUALITY-NORMATIVE", "请人工核查该陈述是否明确表达可判定的规范义务", loc + "/statement", object_id=rid, clause=clause, source="review-quality", advisory=True)
        methods_check(req.get("qualification_methods", []), loc + "/qualification_methods", rid)
        if not req.get("applicability", "").strip() and not req.get("state_ids") and states:
            add("SRS-APPLICABILITY", "存在状态/方式时应明确本需求适用范围", loc, object_id=rid, clause=clause, source="GJB-J.3.1")
        if policy.get("mode") == "ranked" and tailoring.get("3.18", {}).get("status") != "not_applicable":
            for key in ("priority", "criticality"):
                if not req.get(key, "").strip():
                    add("SRS-PRIORITY", f"项目选择逐项分配策略，但本记录缺少 {key}", loc + "/" + key, object_id=rid, clause="3.18", source="project-model")
        for key, target in (("state_ids", states), ("data_ids", data)):
            for identifier in req.get(key, []):
                target_check(identifier, target, loc + "/" + key, rid)
        if req.get("capability_id"):
            target_check(req["capability_id"], capabilities, loc + "/capability_id", rid)
        if req.get("interface_id"):
            target_check(req["interface_id"], interfaces, loc + "/interface_id", rid)
            interface = interfaces.get(req["interface_id"], {})
            expected_clause = "3.4" if interface.get("kind") == "internal" else "3.3"
            if clause != expected_clause:
                add("SRS-INTERFACE-CLAUSE", f"接口需求应归属 {expected_clause}", loc + "/clause", object_id=rid, hard=True)
        elif clause in {"3.3", "3.4"}:
            add("SRS-INTERFACE-LINK", "接口需求尚未关联具体接口 ID", loc, object_id=rid, clause=clause)
        source_ids = req.get("source_ids", [])
        if req.get("source") and source_ids and req["source"] not in source_ids:
            add("SRS-SOURCE-ALIAS-CONFLICT", "旧 source 字段与 source_ids 矛盾", loc, object_id=rid, hard=True)
        if not source_ids and req.get("source"):
            source_ids = [req["source"]]
        if not source_ids and not req.get("derivation"):
            add("SRS-SOURCE-MISSING", "需求缺少上层来源或明确派生依据", loc, object_id=rid, clause="5", source="GJB-J.5")
        for identifier in source_ids:
            # Legacy data may have source labels but no source inventory. Draft
            # reports that missing inventory; review blocks it. A populated
            # inventory with an unknown identifier is a definite link error.
            if identifier not in sources:
                add("SRS-SOURCE-UNKNOWN", f"来源未登记在独立来源台账：{identifier}", loc + "/source_ids", object_id=rid, clause="5", hard=bool(sources))
            elif sources[identifier]["allocation"] != "allocated":
                add("SRS-SOURCE-ALLOCATION", "需求引用了未分配给本 CSCI 的来源", loc + "/source_ids", object_id=rid, clause="5", hard=True)
            expected_pairs.add((identifier, rid))

    external_targets: dict[str, dict[str, str]] = {}
    registered_irs = {row.get("irs_reference", {}).get("reference_id") for row in interfaces.values() if row.get("definition") == "irs_reference"}
    for key in ("forwardTrace", "reverseTrace", "qualification"):
        for i, row in enumerate(model.get(key, [])):
            if "external_target" not in row:
                continue
            target = row["external_target"]
            target_id = external_target_key(target)
            external_targets[target_id] = target
            target_check(target["reference_id"], references, f"/{key}/{i}/external_target/reference_id", target["requirement_id"])
            if target["reference_id"] not in registered_irs:
                add("SRS-IRS-UNREGISTERED", "外部需求目标的 IRS 未在接口登记中作为受控引用", f"/{key}/{i}/external_target", object_id=target["requirement_id"], hard=True)
    external_pairs = {(row["source"], external_target_key(row["external_target"])) for row in model.get("forwardTrace", []) if "external_target" in row}
    expected_pairs.update(external_pairs)

    def row_target(row: dict[str, Any]) -> str:
        return external_target_key(row["external_target"]) if "external_target" in row else row["requirement_id"]

    all_targets = dict(requirements)
    all_targets.update(external_targets)
    for i, source in enumerate(model.get("source_requirements", [])):
        sid = source["id"]
        if source.get("reference_id"):
            target_check(source["reference_id"], references, f"/source_requirements/{i}/reference_id", sid)
        else:
            add("SRS-SOURCE-LOCATION", "来源台账缺少受控文件编号", f"/source_requirements/{i}", object_id=sid, clause="5")
        if not source.get("location", "").strip():
            add("SRS-SOURCE-LOCATION", "来源台账缺少条款/页码定位", f"/source_requirements/{i}", object_id=sid, clause="5")
        if source["allocation"] == "allocated":
            covered = any(pair[0] == sid for pair in expected_pairs)
            if not covered:
                if source.get("disposition") in {"not_applicable", "deferred"} and source.get("reason", "").strip():
                    if source["disposition"] == "deferred":
                        add("SRS-SOURCE-DEFERRED", "已分配来源仍待落实", f"/source_requirements/{i}", object_id=sid, clause="5")
                else:
                    add("SRS-SOURCE-UNCOVERED", "已分配来源没有 SRS 需求或有依据的处置", f"/source_requirements/{i}", object_id=sid, clause="5", source="GJB-J.5")
            elif source.get("disposition") in {"not_applicable", "deferred"}:
                add("SRS-SOURCE-CONFLICT", "已落实来源与其处置状态矛盾", f"/source_requirements/{i}", object_id=sid, clause="5", hard=True)

    qualification: dict[str, set[str]] = {}
    qualification_rows: set[str] = set()
    for i, row in enumerate(model.get("qualification", [])):
        rid = row_target(row)
        target_check(rid, all_targets, f"/qualification/{i}", rid)
        methods_check(row["methods"], f"/qualification/{i}/methods", rid)
        signature = json.dumps(dict(row, methods=sorted(set(row["methods"]))), sort_keys=True, ensure_ascii=False)
        if signature in qualification_rows:
            add("SRS-QUALIFICATION-DUPLICATE", "合格性计划存在完全相同的重复记录", f"/qualification/{i}", object_id=rid, hard=True)
        qualification_rows.add(signature)
        qualification.setdefault(rid, set()).update(row["methods"])
        for key in ("condition", "evidence", "pass_criterion"):
            if not row.get(key, "").strip():
                add("SRS-QUALIFICATION-INCOMPLETE", f"合格性计划缺少 {key}", f"/qualification/{i}/{key}", object_id=rid, clause="4", source="review-quality")
    for rid, req in requirements.items():
        if rid not in qualification:
            add("SRS-QUALIFICATION-MISSING", "需求未纳入合格性规定", "/qualification", object_id=rid, clause="4", source="GJB-J.4")
        elif qualification[rid] != set(req.get("qualification_methods", [])):
            add("SRS-QUALIFICATION-CONFLICT", "合格性规定与需求记录的方法不一致", "/qualification", object_id=rid, clause="4", hard=True)
    for rid in {pair[1] for pair in external_pairs} - set(qualification):
        add("SRS-QUALIFICATION-MISSING", "受控 IRS 外部需求目标未纳入合格性规定", "/qualification", object_id=rid, clause="4", source="GJB-J.4")

    actual_pairs: set[tuple[str, str]] = set()
    for i, row in enumerate(model.get("forwardTrace", [])):
        pair = (row["source"], row_target(row))
        target_check(pair[1], all_targets, f"/forwardTrace/{i}", pair[1])
        if sources:
            target_check(pair[0], sources, f"/forwardTrace/{i}/source", pair[0])
        elif pair[0]:
            add("SRS-SOURCE-UNKNOWN", "正向追踪来源缺少独立来源台账", f"/forwardTrace/{i}/source", object_id=pair[0], clause="5")
        if pair in actual_pairs:
            add("SRS-TRACE-DUPLICATE", "重复正向追踪关系", f"/forwardTrace/{i}", object_id=pair[1], hard=True)
        actual_pairs.add(pair)
        if pair not in expected_pairs:
            add("SRS-TRACE-CONFLICT", "正向追踪与需求来源声明不一致", f"/forwardTrace/{i}", object_id=pair[1], clause="5", hard=True)
    for sid, rid in sorted(expected_pairs - actual_pairs):
        add("SRS-FORWARD-MISSING", f"正向追踪缺少 {sid} → {rid}", "/forwardTrace", object_id=rid, clause="5", source="GJB-J.5")
    reverse_pairs: set[tuple[str, str]] = set()
    reverse_ids: set[str] = set()
    for i, row in enumerate(model.get("reverseTrace", [])):
        rid = row_target(row)
        target_check(rid, all_targets, f"/reverseTrace/{i}", rid)
        reverse_ids.add(rid)
        row_sources = row.get("source_ids", []) or ([row["source"]] if row.get("source") else [])
        for sid in row_sources:
            pair = (sid, rid)
            if pair in reverse_pairs:
                add("SRS-TRACE-DUPLICATE", "重复反向追踪关系", f"/reverseTrace/{i}", object_id=rid, hard=True)
            reverse_pairs.add(pair)
            if pair not in expected_pairs:
                add("SRS-TRACE-CONFLICT", "反向追踪与需求来源声明不一致", f"/reverseTrace/{i}", object_id=rid, clause="5", hard=True)
        if rid in requirements and row.get("methods") is not None and set(row["methods"]) != set(requirements[rid].get("qualification_methods", [])):
            add("SRS-TRACE-METHOD-CONFLICT", "反向追踪与需求方法不一致", f"/reverseTrace/{i}", object_id=rid, hard=True)
        if not row_sources and not row.get("derivation_basis"):
            add("SRS-REVERSE-SOURCE", "反向追踪缺少来源或派生依据", f"/reverseTrace/{i}", object_id=rid, clause="5")
    for sid, rid in sorted(expected_pairs - reverse_pairs):
        add("SRS-REVERSE-MISSING", f"反向追踪缺少 {rid} → {sid}", "/reverseTrace", object_id=rid, clause="5", source="GJB-J.5")
    for rid, req in requirements.items():
        if req.get("derivation") and rid not in reverse_ids:
            add("SRS-REVERSE-MISSING", "派生需求缺少反向追踪说明", "/reverseTrace", object_id=rid, clause="5")

    content_clauses = {row["clause"] for row in requirements.values()}
    if states: content_clauses.add("3.1")
    if capabilities: content_clauses.add("3.2")
    if data: content_clauses.add("3.5")
    for row in interfaces.values(): content_clauses.add("3.4" if row.get("kind") == "internal" else "3.3")
    if policy: content_clauses.add("3.18")
    if requirements and all(row.get("priority") and row.get("criticality") for row in requirements.values()) and tailoring.get("3.18", {}).get("status") != "not_applicable": content_clauses.add("3.18")
    if any(key.startswith("3.11.") for key in content_clauses): content_clauses.add("3.11")
    for clause in SRS_CLAUSES:
        tailored = tailoring.get(clause, {})
        if clause.startswith("3.11.") and tailoring.get("3.11", {}).get("status") == "not_applicable":
            if clause in content_clauses or tailored.get("status") in {"applicable", "tbd"}:
                add("SRS-TAILORING-CONFLICT", "资源子条内容/状态与父条不适用声明矛盾", "/tailoring", clause=clause, hard=True)
            continue
        if tailored.get("status") == "not_applicable":
            if clause in content_clauses:
                add("SRS-TAILORING-CONFLICT", "条款已有内容却声明不适用", "/tailoring", clause=clause, hard=True)
            if not tailored.get("reason", "").strip() or not tailored.get("basis", "").strip():
                add("SRS-TAILORING-BASIS", "不适用条款需要具体原因和裁剪依据", "/tailoring", clause=clause)
        elif clause not in content_clauses:
            add("SRS-CLAUSE-UNRESOLVED", "条款尚无内容，也未提供有依据的不适用说明", "/tailoring", clause=clause, source="GJB-J")

    known_ids = set(requirements) | set(sources) | set(states) | set(capabilities) | set(interfaces) | set(data) | set(references)
    for i, row in enumerate(model.get("tbd", [])):
        tid = row["id"]
        for key in ("owner", "closure_condition", "status"):
            if not row.get(key, "").strip():
                add("SRS-TBD-INCOMPLETE", f"未决事项缺少 {key}", f"/tbd/{i}/{key}", object_id=tid)
        if not row.get("affected_ids") and not row.get("impact", "").strip():
            add("SRS-TBD-IMPACT", "未决事项缺少受影响对象或范围", f"/tbd/{i}", object_id=tid)
        for identifier in row.get("affected_ids", []):
            target_check(identifier, {key: True for key in known_ids}, f"/tbd/{i}/affected_ids", tid)
        if row.get("status", "").lower() in {"closed", "resolved", "已关闭", "已解决"}:
            if not row.get("closure_evidence", "").strip():
                add("SRS-TBD-CLOSURE", "已关闭的未决事项缺少关闭证据", f"/tbd/{i}", object_id=tid)
        else:
            add("SRS-TBD-OPEN", "未决事项尚未关闭", f"/tbd/{i}", object_id=tid)
    # Resolve a marker within its owning record. A sibling basis or explicit
    # tbd_ids may control that record; unrelated root-level notes never do.
    # Referencing a closed TBD does not resolve still-unresolved prose.
    tbd_token = re.compile(r"(?<![A-Za-z0-9_])TBD[-_][A-Za-z0-9_-]+", re.I)
    open_tbds = {tid for tid, row in tbds.items() if row.get("status", "").lower() not in {"closed", "resolved", "已关闭", "已解决"}}

    def record_links(record: dict) -> set[str]:
        links = set(record.get("tbd_ids", []))
        for item in record.values():
            if isinstance(item, str):
                links.update(tbd_token.findall(item))
        return links

    def source_links(source_id: str) -> set[str]:
        source = sources.get(source_id, {})
        # A source's file-identity uncertainty belongs to its explicitly linked
        # reference, never to an unrelated global "all documents" note.
        return record_links(source) | record_links(references.get(source.get("reference_id"), {}))

    def scan(value: Any, path: list[Any], owner: str = "", context: frozenset = frozenset()) -> None:
        if isinstance(value, dict):
            record = len(path) == 2 and isinstance(path[1], int)
            owner = value.get("id", value.get("number", value.get("requirement_id", owner)))
            links = set() if record else set(context)
            if record and path[0] == "source_requirements":
                links.update(record_links(references.get(value.get("reference_id"), {})))
            if record and path[0] in {"forwardTrace", "reverseTrace"}:
                source_ids = value.get("source_ids", []) or ([value["source"]] if value.get("source") else [])
                for source_id in source_ids:
                    links.update(source_links(source_id))
            if path:
                for key, child in value.items():
                    if isinstance(child, str):
                        links.update(tbd_token.findall(child))
                    elif key == "tbd_ids" and isinstance(child, list):
                        links.update(child)
                links.update(tid for tid, row in tbds.items() if owner and owner in row.get("affected_ids", []))
            for key, child in value.items():
                if not path and key == "tbd":
                    continue
                if key == "tbd_ids" and isinstance(child, list):
                    for tid in child:
                        if tid not in tbds:
                            add("SRS-TBD-UNKNOWN", "显式TBD引用未登记：" + tid, _pointer(path + [key]), object_id=owner, hard=True)
                scan(child, path + [key], owner, frozenset(links))
        elif isinstance(value, list):
            for i, child in enumerate(value):
                scan(child, path + [i], owner, context)
        elif isinstance(value, str):
            tokens = set(tbd_token.findall(value))
            for tid in tokens - set(tbds):
                add("SRS-TBD-UNKNOWN", "未决标识未登记：" + tid, _pointer(path), object_id=owner, hard=True)
            if re.search(r"待确认|待补充|待定|\bTBD(?:\b|[-：:])", value, re.I):
                if not ((tokens | set(context)) & open_tbds):
                    add("SRS-TBD-UNLINKED", "未决标记没有关联本对象的开放TBD记录", _pointer(path), object_id=owner)
    scan(model, [])
    from srs_authoring import validate_authoring, writing_profile as select_profile
    issues.extend(validate_authoring(model, mode, select_profile(model, writing_profile)))
    return issues


class SrsValidationError(ValueError):
    """Validation failure with located findings for JSON/CLI consumers."""

    def __init__(self, issues: list[dict[str, Any]]) -> None:
        self.issues = issues
        errors = [item for item in issues if item["severity"] == "error"]
        super().__init__("SRS model validation failed: " + "; ".join(f"{item['rule_id']} {item['location']}: {item['message']}" for item in errors))


def require_valid_srs_content(content: Any, mode: str = "draft", writing_profile: str | None = None) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Convenience API for generators; does not modify caller-owned content."""
    model = normalize_srs_content(content)
    issues = validate_srs_content(model, mode, writing_profile)
    errors = [item for item in issues if item["severity"] == "error"]
    if errors:
        raise SrsValidationError(issues)
    return model, issues
