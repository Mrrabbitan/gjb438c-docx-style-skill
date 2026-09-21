#!/usr/bin/env python3
"""Build a public, explicitly fictional source model for SRS regression checks."""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def fixture():
    sources = [
        ("UP-01", "启动后处于就绪状态，归档处理中拒绝重复归档命令。"),
        ("UP-02", "每次合法归档命令产生唯一归档记录。"),
        ("UP-03", "通过输入文件接口接收记录，拒绝空记录编号。"),
        ("UP-04", "通过结果文件接口向接收端输出归档结果。"),
        ("UP-05", "保存归档记录编号和处理结果，编号不得重复。"),
    ]
    definitions = [
        ("REQ-状态-01", "3.1", "归档CSCI应在启动完成后进入就绪状态。", "UP-01", "启动完成", "观察到就绪状态且可以接收一条归档指令", {"state_ids": ["READY"]}),
        ("REQ-状态-01B", "3.1", "归档CSCI应在接受合法输入后进入处理状态，直至当前结果输出完成后返回就绪状态。", "UP-01", "READY状态接受合法记录A，等待当前处理完成", "观察到READY→BUSY→READY，BUSY覆盖记录建立到结果输出完成的过程", {"state_ids": ["READY", "BUSY"]}),
        ("REQ-状态-01C", "3.1", "归档CSCI在处理状态收到当前记录编号的重复归档命令时，应拒绝该命令且不产生额外归档记录或结果文件。", "UP-01", "合法记录A已接受，当前处理尚未完成时再次提交A", "重复命令未生成额外记录或结果文件，原命令仍完成一次", {"state_ids": ["BUSY"]}),
        ("REQ-归档-02", "3.2", "归档CSCI应为每条合法归档命令建立一条唯一归档记录。", "UP-02", "就绪状态，提交一条合法命令", "恰好新增一条具有唯一编号的归档记录", {"capability_id": "CAP-ARCHIVE", "state_ids": ["READY"], "data_ids": ["DATA-RECORD"]}),
        ("REQ-输入-03", "3.3", "归档CSCI应拒绝记录编号为空的输入文件，并返回错误码EMPTY_ID。", "UP-03", "输入文件中记录编号为空", "无新增归档记录，返回EMPTY_ID", {"interface_id": "IF-INGEST", "state_ids": ["READY"]}),
        ("REQ-输入-03A", "3.3", "归档CSCI应在就绪状态接受IF-INGEST约定格式且编号非空、尚不存在于归档集合的输入文件，将该文件作为一条合法归档命令交由CAP-ARCHIVE处理。", "UP-03", "READY状态输入合法JSON文件，record_id为新的非空编号A，payload为字符串", "文件被接受并进入BUSY，CAP-ARCHIVE收到相同record_id与payload，产生一条编号A的记录", {"interface_id": "IF-INGEST", "state_ids": ["READY"]}),
        ("REQ-输出-04", "3.3", "归档CSCI应按IF-EXPORT约定格式输出结果文件，record_id保持输入值；成功归档时status为SUCCESS，空编号拒绝时status为EMPTY_ID。", "UP-04", "分别输入合法新记录和空编号记录，结果目录可写", "两次输出均可解析为约定JSON对象；前者编号与输入相同且status为SUCCESS，后者编号为空字符串且status为EMPTY_ID", {"interface_id": "IF-EXPORT", "state_ids": ["READY", "BUSY"]}),
        ("REQ-数据-05", "3.5", "归档CSCI应以记录编号唯一标识每一条归档记录。", "UP-05", "已存在编号A，再次提交相同编号", "第二次提交未新增具有相同编号的记录", {"data_ids": ["DATA-RECORD"], "state_ids": ["READY"]}),
        ("REQ-数据-05B", "3.5", "归档CSCI应在成功归档时将原始record_id及处理结果SUCCESS保存至DATA-RECORD，并在本次运行的后续查询中提供这两个值。", "UP-05", "成功归档编号A后，在同一次运行中读取归档集合", "存在且仅存在一条编号A的记录，保存的record_id与输入一致且status为SUCCESS", {"data_ids": ["DATA-RECORD"], "state_ids": ["READY", "BUSY"]}),
    ]
    requirements, qualification, forward, reverse = [], [], [], []
    for rid, clause, statement, sid, condition, criterion, extra in definitions:
        requirement = dict(id=rid, name=rid, clause=clause, statement=statement,
                           applicability=condition, qualification_methods=["测试"],
                           source_ids=[sid], evidence="EV-" + rid, **extra)
        requirements.append(requirement)
        qualification.append(dict(requirement_id=rid, methods=["测试"], condition=condition,
                                  evidence="EV-" + rid, pass_criterion=criterion))
        forward.append(dict(source=sid, requirement_id=rid, classification="软件验收需求", location=clause, disposition="落实"))
        reverse.append(dict(requirement_id=rid, source_ids=[sid], methods=["测试"], category=clause, evidence="EV-" + rid))
    return {
        "schema_version": "2.0", "writing_profile": "gjb-standard",
        "metadata": {
            "document_type": "SRS", "title": "归档演示软件需求规格说明", "version": "1.0",
            "classification": "公开测试资料", "csci_id": "DEMO-ARCHIVE", "document_id": "DEMO-SRS",
            "identification": "文档DEMO-SRS，第1.0版；归档演示软件，CSCI标识DEMO-ARCHIVE。",
            "system_overview": "本测试软件以文件交换方式接收归档记录并输出处理结果。所有名称、来源及指标均为回归测试虚构资料。",
            "document_overview": "本说明规定演示软件的需求、合格性方法和来源关系；仅供工具回归验证。",
            "phase": "测试样例", "author": "样例编写角色", "reviewer": "样例审核角色",
            "standards_reviewer": "样例标审角色", "approver": "样例批准角色", "unit": "公开示例",
            "date": "2026-09-15", "change_reason": "建立可重复的工具回归样例",
            "change_content": "建立虚构软件需求及对应来源和验证计划。",
        },
        "references": [
            {"number": "DEMO-SSS", "title": "归档演示系统需求测试基线", "organization": "公开示例", "revision": "1.0", "date": "2026-09-15", "normative": True, "source": "仓库tests/fixtures/DEMO-SSS.md"},
            {"number": "DEMO-IRS", "title": "归档演示文件接口测试基线", "organization": "公开示例", "revision": "1.0", "date": "2026-09-15", "normative": True, "source": "仓库tests/fixtures/DEMO-IRS.md"},
        ],
        "source_requirements": [dict(id=sid, statement=text, reference_id="DEMO-SSS", location="第3章 " + sid,
                                     allocation="allocated", disposition="implemented") for sid, text in sources],
        "states": [
            {"id": "READY", "name": "就绪", "definition": "已完成初始化或上一条合法命令处理，可以接受归档命令。", "entry": "初始化完成或当前结果输出完成", "exit": "接受合法输入进入BUSY；软件结束运行时退出状态机", "transitions": [{"to": "BUSY", "condition": "接受合法归档输入"}]},
            {"id": "BUSY", "name": "处理", "definition": "已接受一条合法归档命令，正在建立记录和输出结果；拒绝当前记录编号的重复命令。", "entry": "READY接受合法归档输入", "exit": "当前结果输出完成后返回READY", "transitions": [{"to": "READY", "condition": "当前结果输出完成"}]},
        ],
        "capabilities": [{"id": "CAP-ARCHIVE", "name": "记录归档", "purpose": "建立唯一编号的归档记录。"}],
        "interfaces": [
            {"id": "IF-INGEST", "name": "归档文件输入", "kind": "external", "definition": "inline",
             "csci_endpoint": "DEMO-ARCHIVE", "external_endpoint": "输入文件提供端", "direction": "输入",
             "data": "无BOM的UTF-8 JSON文件，每个文件恰含一个对象且仅含record_id、payload两个字符串字段；顺序不限。record_id不作空白裁剪，空字符串表示空编号；payload允许空字符串。合法归档输入还须record_id非空且未存在于归档集合。", "protocol": "本地完整文件交换，一份输入文件对应一次提交；无网络会话。试验装置可在BUSY阶段重复提交当前记录编号。",
             "exception": "空编号返回EMPTY_ID。", "status": "测试基线1.0"},
            {"id": "IF-EXPORT", "name": "处理结果输出", "kind": "external", "definition": "inline",
             "csci_endpoint": "DEMO-ARCHIVE", "external_endpoint": "结果文件接收端", "direction": "输出",
             "data": "无BOM的UTF-8 JSON文件，每个文件恰含一个对象且仅含record_id、status两个字符串字段；顺序不限。record_id为原始输入编号（空编号保持空字符串），status为SUCCESS或EMPTY_ID。BUSY期间被拒绝的重复命令不生成额外结果文件。", "protocol": "本地完整文件交换，结果目录可写；无网络会话。结果文件与触发它的输入由试验提交记录关联。",
             "exception": "本测试基线不包含结果目录写入故障注入或相应验收约束，范围见DEMO-SSS第1章。", "status": "测试基线1.0"},
        ],
        "interface_diagram": "接口关系见受控测试文件DEMO-IRS第3章：输入文件提供端经IF-INGEST连接DEMO-ARCHIVE，DEMO-ARCHIVE经IF-EXPORT连接结果文件接收端。",
        "data": [{"id": "DATA-RECORD", "name": "归档记录", "description": "归档记录包含唯一编号和处理结果。",
                  "fields": "record_id：成功归档的原始非空字符串；status：SUCCESS。空编号拒绝不创建归档记录。", "constraints": "记录编号在本次运行的归档集合内唯一；同次运行可读取保存值。跨运行持久化和保留期限不在测试基线范围内。"}],
        "requirements": requirements, "qualification": qualification,
        "forwardTrace": forward, "reverseTrace": reverse, "tbd": [],
        "priority_policy": {"mode": "equal", "statement": "本测试基线内全部需求采用相同优先顺序和关键性权重。"},
        "tailoring": [dict(clause=c, status="not_applicable", reason="该虚构文件归档测试基线未分配此类验收约束。", basis="DEMO-SSS第1章测试范围")
                      for c in ["3.4", "3.6", "3.7", "3.8", "3.9", "3.10", "3.11", "3.12", "3.13", "3.14", "3.15", "3.16", "3.17"]],
        "notes": "CSCI表示计算机软件配置项；本模型中的角色、来源文件和证据编号均为虚构测试数据。验证计划不表示已经取得产品测试结果。",
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(fixture(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
