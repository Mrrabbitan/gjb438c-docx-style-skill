"""Regression cases for source coverage and controlled SRS model validation."""
from copy import deepcopy
from pathlib import Path
import sys
import unittest

SCRIPTS = Path(__file__).resolve().parents[1] / "gjb438c-docx-style" / "scripts"
sys.path.insert(0, str(SCRIPTS))
from srs_model import normalize_srs_content, validate_srs_content


def valid_model():
    missing = ["3.1", "3.3", "3.4", "3.5", "3.6", "3.7", "3.8", "3.9", "3.10", "3.11", "3.12", "3.13", "3.14", "3.15", "3.16", "3.17"]
    return {
        "schema_version": "2.0",
        "metadata": {"document_type": "SRS", "title": "校验样例", "version": "1.0", "classification": "公开", "csci_id": "CSCI-1", "system_overview": "本示例为一个离线记录查询配置项。", "document_overview": "本文规定记录查询需求及其验收依据。"},
        "references": [{"number": "SYS-DOC", "title": "示例系统规格", "organization": "示例单位", "revision": "1.0", "date": "2026-09-15"}],
        "requirements": [{"id": "REQ-任意#1", "clause": "3.2", "statement": "软件应返回编号匹配的记录。", "source_ids": ["SYS-1"], "qualification_methods": ["测试"]}],
        "source_requirements": [{"id": "SYS-1", "statement": "查询指定记录", "reference_id": "SYS-DOC", "location": "3.2.1", "allocation": "allocated", "disposition": "implemented"}],
        "qualification": [{"requirement_id": "REQ-任意#1", "methods": ["测试"], "condition": "含记录 R1 的固定数据集", "evidence": "TEST-1", "pass_criterion": "输入 R1 时返回且只返回该记录。"}],
        "forwardTrace": [{"source": "SYS-1", "requirement_id": "REQ-任意#1", "classification": "能力", "location": "3.2.1", "disposition": "落实"}],
        "reverseTrace": [{"requirement_id": "REQ-任意#1", "source_ids": ["SYS-1"], "category": "能力", "methods": ["测试"], "evidence": "TEST-1"}],
        "priority_policy": {"mode": "equal", "statement": "所有需求具有同等优先级和关键性。"},
        "tailoring": [{"clause": clause, "status": "not_applicable", "reason": "本示例配置项边界不包含该类需求。", "basis": "SYS-DOC 1.2 配置项分配边界"} for clause in missing],
        "tbd": [],
    }


def external_model():
    model = valid_model()
    model["references"].append({"number": "IRS-DOC", "title": "受控接口需求", "organization": "接口单位", "revision": "B", "date": "2026-09-15"})
    model["requirements"] = []
    model["interfaces"] = [{"id": "IF-A", "name": "外部接口", "definition": "irs_reference", "irs_reference": {"reference_id": "IRS-DOC", "location": "3.2"}}]
    model["tailoring"] = [item for item in model["tailoring"] if item["clause"] != "3.3"]
    model["tailoring"].append({"clause": "3.2", "status": "not_applicable", "reason": "示例需求全部由受控 IRS 定义。", "basis": "SYS-DOC 1.2"})
    target = {"reference_id": "IRS-DOC", "requirement_id": "EXT-需求-7", "location": "3.2.1"}
    for key in ("qualification", "forwardTrace", "reverseTrace"):
        model[key][0].pop("requirement_id")
        model[key][0]["external_target"] = deepcopy(target)
    return model


class SrsModelTests(unittest.TestCase):
    def errors(self, model, mode="review"):
        return [row for row in validate_srs_content(model, mode) if row["severity"] == "error"]

    def test_review_accepts_controlled_free_ids_and_optional_io(self):
        self.assertEqual(self.errors(valid_model()), [])

    def test_schema_rejects_v2_wrong_type_and_unknown_properties(self):
        model = valid_model()
        model["requirements"][0]["qualification_methods"] = "测试"
        model["dataa"] = []
        issues = self.errors(model, "draft")
        self.assertTrue(issues)
        self.assertTrue(all(row["rule_id"] == "SRS-SCHEMA" for row in issues))
        self.assertTrue(any("qualification_methods" in row["location"] for row in issues))

    def test_legacy_normalization_is_copy_and_does_not_invent_sources(self):
        model = valid_model()
        model.pop("schema_version")
        model.pop("source_requirements")
        req = model["requirements"][0]
        req["source"] = req.pop("source_ids")[0]
        req["qualification_methods"] = "测试"
        original = deepcopy(model)
        converted = normalize_srs_content(model)
        self.assertEqual(model, original)
        self.assertEqual(converted["schema_version"], "2.0")
        self.assertEqual(converted["requirements"][0]["source_ids"], ["SYS-1"])
        self.assertNotIn("source_requirements", converted)
        self.assertTrue(any(row["rule_id"] == "SRS-SOURCE-UNKNOWN" for row in self.errors(converted)))

    def test_empty_model_is_draft_only(self):
        model = {"schema_version": "2.0", "metadata": {"document_type": "SRS", "title": "草稿"}, "requirements": []}
        self.assertEqual(self.errors(model, "draft"), [])
        self.assertTrue(any(row["rule_id"] == "SRS-REQUIREMENTS-EMPTY" for row in self.errors(model)))

    def test_independent_allocated_source_cannot_disappear(self):
        model = valid_model()
        model["source_requirements"].append(dict(model["source_requirements"][0], id="SYS-2"))
        self.assertTrue(any(row["rule_id"] == "SRS-SOURCE-UNCOVERED" and row["object_id"] == "SYS-2" for row in self.errors(model)))

    def test_trace_is_not_automatically_filled(self):
        model = valid_model()
        model["forwardTrace"] = []
        self.assertTrue(any(row["rule_id"] == "SRS-FORWARD-MISSING" for row in self.errors(model)))

    def test_resource_parent_tailoring_covers_subtree(self):
        model = valid_model()
        self.assertFalse(any(row.get("clause", "").startswith("3.11") for row in self.errors(model)))
        model["requirements"][0]["clause"] = "3.11.1"
        self.assertTrue(any(row["rule_id"] == "SRS-TAILORING-CONFLICT" for row in self.errors(model)))

    def test_priority_clause_can_be_explicitly_not_applicable(self):
        model = valid_model()
        model.pop("priority_policy")
        model["tailoring"].append({"clause": "3.18", "status": "not_applicable", "reason": "项目未规定优先顺序和关键性。", "basis": "SYS-DOC 1.2"})
        self.assertEqual(self.errors(model), [])

    def test_same_method_different_conditions_are_valid(self):
        model = valid_model()
        model["qualification"].append(dict(model["qualification"][0], condition="空数据集", evidence="TEST-2", pass_criterion="返回空结果。"))
        self.assertEqual(self.errors(model), [])
        model["qualification"].append(deepcopy(model["qualification"][0]))
        self.assertTrue(any(row["rule_id"] == "SRS-QUALIFICATION-DUPLICATE" for row in self.errors(model)))

    def test_empty_tbd_register_does_not_close_unresolved_fact(self):
        model = valid_model()
        model["requirements"][0]["statement"] = "软件应返回待确认范围内的记录。"
        self.assertTrue(any(row["rule_id"] == "SRS-TBD-UNLINKED" for row in self.errors(model)))
        model["tbd"] = [{"id": "TBD-1", "issue": "确认记录范围", "affected_ids": ["REQ-任意#1"], "owner": "需求负责人", "closure_condition": "批准范围表", "status": "open"}]
        codes = {row["rule_id"] for row in self.errors(model)}
        self.assertNotIn("SRS-TBD-UNLINKED", codes)
        self.assertIn("SRS-TBD-OPEN", codes)

    def test_state_requirements_and_unknown_state_links(self):
        model = valid_model()
        model["requirements"][0].update(clause="3.1", state_ids=["RUN"])
        model["states"] = [{"id": "RUN", "name": "运行", "definition": "正常查询状态"}]
        model["tailoring"] = [item for item in model["tailoring"] if item["clause"] != "3.1"]
        model["tailoring"].append({"clause": "3.2", "status": "not_applicable", "reason": "状态示例不定义其他能力。", "basis": "SYS-DOC 1.2"})
        self.assertEqual(self.errors(model), [])
        model["requirements"][0]["state_ids"] = ["MISSING"]
        self.assertTrue(any(row["rule_id"] == "SRS-LINK-UNKNOWN" for row in self.errors(model)))

    def test_controlled_method_extension(self):
        model = valid_model()
        model["qualification_method_extensions"] = [{"name": "仿真验证", "definition": "采用受控仿真条件判定结果", "approval": "SYS-DOC 4.2"}]
        model["requirements"][0]["qualification_methods"] = ["仿真验证"]
        model["qualification"][0]["methods"] = ["仿真验证"]
        model["reverseTrace"][0]["methods"] = ["仿真验证"]
        self.assertEqual(self.errors(model), [])
        model["qualification_method_extensions"] = []
        self.assertTrue(any(row["rule_id"] == "SRS-METHOD-UNKNOWN" for row in self.errors(model)))

    def test_external_irs_requirement_is_controlled_without_fake_local_id(self):
        model = external_model()
        self.assertEqual(self.errors(model), [])
        self.assertEqual(model["requirements"], [])
        model["interfaces"] = []
        self.assertTrue(any(row["rule_id"] == "SRS-IRS-UNREGISTERED" for row in self.errors(model)))

    def test_quality_heuristics_are_review_advice(self):
        model = valid_model()
        model["requirements"][0]["statement"] = "响应正常。"
        issues = validate_srs_content(model, "review")
        self.assertTrue(any(row["rule_id"] == "SRS-QUALITY-NORMATIVE" and row["severity"] == "warning" for row in issues))
        model["requirements"][0]["statement"] = "系统应在必要时（负载超过 80% 时）拒绝请求。"
        self.assertEqual(self.errors(model), [])


if __name__ == "__main__":
    unittest.main()
