# Template Structure

Use this file when planning or restructuring document content. Apply the fixed nine-chapter structure only when the document type is `overall-technical-solution`. For standard GJB 438C-2021 software life-cycle documents, first use `references/gjb438c-2021-requirements.md` to identify the document type and its appendix, then map content to that document type instead of forcing the overall technical solution chapters.

Heading numbers shown below are for reference only. Heading paragraph text must NOT contain typed numbers; the template's multilevel list numbering renders them automatically (see `format-lock.md`).

## Applicability Decision

| User request / source document | Structure rule |
|---|---|
| Overall technical solution, `1-2总体技术方案`, project technical scheme, or the bundled project template | Use the fixed nine first-level chapters below. |
| Standard GJB 438C document type such as `SDP`, `SRS`, `SDD`, `STP`, `STD`, `STR`, `SUM`, etc. | Use the corresponding appendix format from GJB 438C-2021. Do not enforce the nine-chapter framework. |
| Style-only repair of an existing DOCX | Preserve the document's existing semantic structure and repair formatting only. |
| Tailored or combined document | Keep required elements complete where included; mark omitted chapters/clauses as `本章无内容` or `本条无内容` with a reason. |

## GJB 438C Document Type Index

The standard identifies 20 principal document types. Use this index for routing; consult `gjb438c-2021-requirements.md` before generating a standard GJB document.

| Abbr. | Chinese document name | Appendix |
|---|---|---|
| SDP | 软件开发计划 | A |
| SIP | 软件安装计划 | B |
| STrP | 软件移交计划 | C |
| STP | 软件测试计划 | D |
| OCD | 运行方案说明 | E |
| SSS | 系统/子系统规格说明 | F |
| IRS | 接口需求规格说明 | G |
| SSDD | 系统/子系统设计说明 | H |
| IDD | 接口设计说明 | I |
| SRS | 软件需求规格说明 | J |
| SDD | 软件设计说明 | K |
| DBDD | 数据库设计说明 | L |
| STD | 软件测试说明 | M |
| STR | 软件测试报告 | N |
| SPS | 软件产品规格说明 | O |
| SVD | 软件版本说明 | P |
| SUM | 软件用户手册 | Q |
| CPM | 计算机编程手册 | R |
| FSM | 固件保障手册 | S |
| SDSR | 软件研制总结报告 | T |

## Front Matter (before chapter 1)

The template front matter must be preserved in this order:

1. Cover table (a real table, not free paragraphs) with five elements:
   - number/classification line (`编号密级`)
   - document class line `项目技术文件` (`文头字`)
   - document identifier `SJZT-XXXX-NNNN-ZF【YYYY/MM/DD】` (`文件标识号`)
   - document title `总体技术方案` (`文件名称`)
   - issuing unit and date, e.g. `技术总师组` + `2024年4月` (`单位名称`)
2. TOC title `目    录` (`311-目录标题`) followed by a real Word `TOC \o "1-3" \h \z` field.
3. A section break: the front-matter section uses upperRoman page numbers with `titlePg`; the body section carries its own headers/footers. Never delete this break.

## Fixed First-Level Chapters

This section is mandatory only for `overall-technical-solution`.

1. `1 概述`
2. `2 使用与技术指标要求`
3. `3 总体架构`
4. `4 分系统设计`
5. `5 系统集成设计`
6. `6 关键技术分析`
7. `7 效能分析`
8. `8 工程组织管理`
9. `9 初步工作计划`

## Required Second-Level Framework

### 1 概述

- `1.1 任务依据`
- `1.2 编制目的`
- `1.3 指导原则`
- `1.4 建设目标`
- `1.5 主要工作`
- `1.6 与其他项目关系`
- `1.7 名词术语`
- `1.8 缩略语`

### 2 使用与技术指标要求

- `2.1 使用要求`
- `2.2 技术指标要求`

### 3 总体架构

- `3.1 体系架构设计`
- `3.2 逻辑架构设计`
- `3.3 系统架构设计`
- `3.4 功能架构设计`
- `3.5 数据架构设计`
- `3.6 部署架构设计`
- `3.7 数据交换设计`
- `3.8 关键性能指标设计`
- `3.9 组织运用模式`
- `3.10 技术体制`

### 4 分系统设计

- Keep `4.x` second-level headings as project-specific分系统 names.
- Each分系统 must use the same internal pattern: `系统结构`, `运行过程`, `接口设计`, `功能组成`, and necessary `数据/配置/模型映射`.

### 5 系统集成设计

- `5.1 集成联试`
- Add only project-specific `5.x` sections when the source material clearly requires them.

### 6 关键技术分析

- Use one `6.x` section per key technology.
- Each key technology section must include: `问题描述`, `技术描述`, and `实现途径`.

### 7 效能分析

- Use one `7.x` section per measurable effectiveness area.
- Prefer operational, quality, delivery, resource, and maintainability effects.

### 8 工程组织管理

- `8.1 组织机构`
- `8.2 考核验收`
- `8.3 安全保密`

### 9 初步工作计划

- Use a milestone table aligned to the template's `序号`, `任务名称`, `开始时间`, `完成时间` pattern.

## Table Patterns

- Terminology table: `编号`, `名称`, `说明`
- Abbreviation table: `序号`, `简写`, `全称`, `解释说明`
- Standard/spec list table: `序号`, `分类`, `名称`, `遵循修订`, `拓展新撰`
- Interface table: `序号`, `接口名称`, `发送方`, `接收方`, `接口描述`, `接口协议`
- Hardware environment table: `序号`, `名称`, `型号规格`, `计量单位`, `数量`
- Software environment table: `序号`, `软件类型`, `软件项名称`, `版本`, `数量`
- Test data table: `分类`, `数据类型`, `数据表`
- Organization table: `序号`, `单位名称`, `项目分工`, `负责人`, `备注`
- Schedule table: `序号`, `任务名称`, `开始时间`, `完成时间`

Table header cells use `145表头`; table body cells use `145表正文`. Captions before tables use `Caption` with chapter-based numbering `表X-Y`.

## Legacy Content Guardrail

The template contains old example business terms. Do not reuse them as content:

- `ZBZQ`
- `XX数据中台`
- `联合XX数据资源体系`
- `XX数据融合治理`
