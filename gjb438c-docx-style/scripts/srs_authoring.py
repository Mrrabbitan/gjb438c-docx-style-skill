"""Training writing conventions, independent of GJB content and review mode.

This module never infers contract approval, development status, or diagrams.
The training profile is a selected writing convention, not a GJB requirement.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Any

WRITING_PROFILES = ("training-review", "gjb-standard")
DEVELOPMENT_LABELS = {"new": "新研", "modified": "改造（升级）", "reused": "沿用", "tbd": "待确认"}
BUSINESS_FLOW_TYPES = frozenset({"activity", "swimlane"})
REQUIREMENT_FLOW_TYPES = BUSINESS_FLOW_TYPES | {"sequence"}


def writing_profile(content: dict, override: str | None = None) -> str:
    value = override or content.get("writing_profile", "training-review")
    if value not in WRITING_PROFILES:
        raise ValueError("writing_profile must be training-review or gjb-standard")
    return value


def capability_outline(content: dict) -> list[dict[str, Any]]:
    """Return preorder rows with node, level, physical_clause and requirements.

    Stable sibling order is (supplied order, input position). No IDs or source
    facts are synthesized. Invalid trees raise rather than silently omit nodes.
    """
    nodes = content.get("capabilities", [])
    index = {}
    children = defaultdict(list)
    for position, node in enumerate(nodes):
        if node["id"] in index:
            raise ValueError("duplicate capability ID: " + node["id"])
        index[node["id"]] = node
        children[node.get("parent_id") or None].append((node.get("order", position), position, node))
    for node in nodes:
        if node.get("parent_id") and node["parent_id"] not in index:
            raise ValueError("unknown capability parent: " + node["parent_id"])
    rows, visiting, visited = [], set(), set()

    def walk(parent, prefix, level):
        for number, (_, _, node) in enumerate(sorted(children[parent]), 1):
            if node["id"] in visiting:
                raise ValueError("capability cycle: " + node["id"])
            if level > 9:
                raise ValueError("capability depth exceeds Word Heading 9: " + node["id"])
            visiting.add(node["id"])
            clause = prefix + "." + str(number)
            rows.append({"node": node, "level": level, "physical_clause": clause,
                         "requirements": [r for r in content.get("requirements", []) if r.get("capability_id") == node["id"] and r.get("clause") == "3.2"]})
            walk(node["id"], clause, level + 1)
            visiting.remove(node["id"])
            visited.add(node["id"])
    walk(None, "3.2", 3)
    if len(visited) != len(nodes):
        raise ValueError("capability cycle or unreachable node: " + ", ".join(sorted(set(index) - visited)))
    return rows


def validate_authoring(content: dict, mode: str, profile: str) -> list[dict]:
    """Schema-valid models only. Missing facts warn in draft, fail in review."""
    findings = []

    def add(rule, message, location="/", object_id="", clause="3.2", hard=False):
        findings.append({"rule_id": "SRS-AUTHORING-" + rule, "message": message,
                         "severity": "error" if hard or mode == "review" else "warning",
                         "source": "project-model" if hard else "training-writing-profile", "location": location,
                         "object_id": object_id, "clause": clause})
    try:
        outline = capability_outline(content)
    except ValueError as exc:
        add("TREE", str(exc), "/capabilities", hard=True)
        return findings
    reqs = {r["id"]: r for r in content.get("requirements", [])}
    caps = {r["node"]["id"]: r for r in outline}
    refs = {r["number"]: r for r in content.get("references", [])}
    sources = {r["id"]: r for r in content.get("source_requirements", [])}
    tbds = {r["id"]: r for r in content.get("tbd", [])}
    figures = content.get("figures", [])
    ids = set()
    for i, fig in enumerate(figures):
        loc = f"/figures/{i}"
        if fig["id"] in ids:
            add("FIGURE-ID", "图标识重复", loc, fig["id"], hard=True)
        ids.add(fig["id"])
        for key, registry in (("capability_ids", caps), ("requirement_ids", reqs)):
            for identifier in fig.get(key, []):
                if identifier not in registry:
                    add("FIGURE-LINK", "图引用不存在的对象：" + identifier, loc + "/" + key, fig["id"], hard=True)
        if fig.get("reference_id") and fig["reference_id"] not in refs:
            add("FIGURE-REFERENCE", "图的受控引用文件不存在", loc, fig["id"], hard=True)
        if not fig.get("path") and not (fig.get("reference_id") and fig.get("location")):
            add("FIGURE-CONTENT", "图尚缺图片路径或受控文件及定位", loc, fig["id"])
        if not fig.get("clause") and not fig.get("capability_ids"):
            add("FIGURE-PLACEMENT", "图未指定正文条款或能力节点", loc, fig["id"])
    for row in outline:
        cap = row["node"]
        if cap.get("physical_clause") and cap["physical_clause"] != row["physical_clause"]:
            add("PHYSICAL-CLAUSE", "能力物理章节与父子顺序不一致；应为" + row["physical_clause"], "/capabilities", cap["id"], hard=True)
        for req in row["requirements"]:
            if req.get("physical_clause") and req["physical_clause"] != row["physical_clause"]:
                add("PHYSICAL-CLAUSE", "需求物理章节与所属能力不一致；应为" + row["physical_clause"], "/requirements", req["id"], hard=True)
    if profile != "training-review":
        return findings
    active = any(r.get("clause") == "3.2" for r in reqs.values()) or bool(caps)
    if active:
        if not content.get("capability_overview", "").strip():
            add("OVERVIEW", "培训写作配置尚缺3.2能力总览", "/capability_overview")
        roots = [row for row in outline if row["level"] == 3]
        if content.get("expected_capability_roots") is not None and len(roots) != content["expected_capability_roots"]:
            add("ROOT-COUNT", "能力入口数量与项目显式配置不一致", "/expected_capability_roots")
        for kinds in ({"composition"}, {"use_case"}, BUSINESS_FLOW_TYPES):
            if not any(f.get("clause") == "3.2" and f["type"] in kinds for f in figures):
                add("SYSTEM-DIAGRAM", "3.2尚缺图类型：" + "/".join(sorted(kinds)), "/figures")
    for row in outline:
        cap, items = row["node"], row["requirements"]
        cid, clause = cap["id"], row["physical_clause"]
        loc = "/capabilities/" + str(content["capabilities"].index(cap))
        if not (cap.get("overview") or cap.get("purpose") or cap.get("description", "")).strip():
            add("NODE-OVERVIEW", "能力节点缺少范围与职责概述", loc, cid, clause)
        if not items and not any(n["node"].get("parent_id") == cid for n in outline):
            add("EMPTY-LEAF", "叶能力没有正式需求，不能作为已完成功能点", loc, cid, clause)
        if not cap.get("development_status") or not cap.get("development_basis", "").strip():
            add("DEVELOPMENT", "节点缺研制状态或其事实依据；未知用tbd并关联未决项", loc, cid, clause)
        matching = [f for f in figures if cid in f.get("capability_ids", [])]
        for kinds in ({"composition"}, BUSINESS_FLOW_TYPES):
            if not any(f["type"] in kinds for f in matching):
                add("NODE-DIAGRAM", "能力节点缺少直接关联的" + "/".join(sorted(kinds)) + "图；系统总图不能替代模块图", loc, cid, clause)
        for req in items:
            if not any(f["type"] in REQUIREMENT_FLOW_TYPES and req["id"] in f.get("requirement_ids", []) for f in matching):
                add("REQUIREMENT-DIAGRAM", "正式需求未关联本模块活动图、泳道图或时序图的明确覆盖范围", "/figures", req["id"], clause)
    for i, req in enumerate(content.get("requirements", [])):
        rid, loc = req["id"], f"/requirements/{i}"
        if req["clause"] == "3.2":
            if not req.get("capability_id"):
                add("ORPHAN", "能力需求尚未分配到父子树节点", loc, rid)
            if not req.get("exception", "").strip() and not (req.get("exception_applicability") == "not_applicable" and req.get("exception_basis", "").strip()):
                add("EXCEPTION", "功能点缺异常/非许可行为，或有依据的不适用说明", loc, rid)
        trace = req.get("contract_trace")
        if not trace:
            add("CONTRACT", "尚未登记合同映射、派生或待定处置；一般来源追踪不能替代合同核对", loc, rid, "5")
            continue
        status = trace["status"]
        for sid in trace.get("source_ids", []):
            if sid not in sources:
                add("CONTRACT-LINK", "合同关联来源不存在：" + sid, loc, rid, "5", hard=True)
            elif sid not in req.get("source_ids", []):
                add("CONTRACT-LINK", "合同关联未纳入该需求的正式来源关系：" + sid, loc, rid, "5", hard=True)
        for tid in trace.get("tbd_ids", []):
            if tid not in tbds:
                add("CONTRACT-TBD", "合同处置引用未登记TBD：" + tid, loc, rid, "5", hard=True)
        if status in {"mapped", "confirmed", "derived"}:
            if not trace.get("source_ids") or any(sources.get(s, {}).get("kind") != "contract" for s in trace.get("source_ids", [])):
                add("CONTRACT-KIND", "已核映射或合同派生必须指向明确登记为contract的原文父来源", loc, rid, "5")
            if any(sources.get(s, {}).get("confirmation") == "pending" for s in trace.get("source_ids", [])):
                add("CONTRACT-BASELINE", "条文映射已登记，但合同原件/采纳状态仍待确认，不能宣称批准合同基线", loc, rid, "5")
        elif status == "pending":
            add("CONTRACT-PENDING", "合同依据受控待确认；本需求尚未获得合同范围确认", loc, rid, "5")
            if not trace.get("basis", "").strip() or not trace.get("tbd_ids"):
                add("CONTRACT-DISPOSITION", "待定合同关联需要保留理由与具体TBD", loc, rid, "5")
        if status in {"derived", "not_applicable"} and not trace.get("basis", "").strip():
            add("CONTRACT-DISPOSITION", "派生或不适用处置需要具体依据", loc, rid, "5")
    return findings
