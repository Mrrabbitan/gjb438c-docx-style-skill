# 总体技术方案模板结构

仅用于用户选用`overall-technical-solution`项目模板时的规划或内容重排。SRS使用[附录J](srs-appendix-j.md)与[P-09配置](srs-p09-format-lock.md)，其他标准文档使用对应附录。九章和模板样式不属于标准对所有软件文档的要求。

下方数字用于说明层级。实际Word标题用多级编号自动生成，不在标题文字中重复手打数字，见[格式规则](format-lock.md)。

## 适用性判断

| 用户选择/文档类型 | 结构规则 |
|---|---|
| 用户选用本项目的总体技术方案模板`1-2总体技术方案` | 使用下方九章；不能仅因名称中含“技术方案”就覆盖用户原结构。 |
| GJB标准文档类型，如SDP、SRS、SDD、STP、STD、STR、SUM等 | 使用对应附录，不执行九章检查。 |
| 已有DOCX纯格式修复 | 保留现有语义结构，只修格式。 |
| 裁剪或合并文档 | 根据4.4保持保留标题的标准顺序及要素完整；裁剪最高节点写`本章无内容`或`本条无内容`并说明理由，整个子树已裁剪时不重复列下级标题。未知内容不能自动当作不适用。 |

## 文档类型路由索引

标准列出20种主要文档。以下索引只用于路由；编写其他类型前读取对应附录，不能宣称本skill已实现其全文语义检查。详见[通用矩阵](gjb438c-2021-requirements.md)。

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

## 正文前置材料

前置材料沿用下列角色和顺序，项目值必须由来源替换，不能继承示例编号或日期：

1. Cover table (a real table, not free paragraphs) with five elements:
   - number/classification line (`编号密级`)
   - document class line `项目技术文件` (`文头字`)
   - document identifier `SJZT-XXXX-NNNN-ZF【YYYY/MM/DD】` (`文件标识号`)
   - document title `总体技术方案` (`文件名称`)
   - issuing unit and date, e.g. `技术总师组` + `2024年4月` (`单位名称`)
2. 有实际修改历史时保留修改页/表；未知时记录缺口，不编造初版批准。
3. 目录标题`目    录`（`311-目录标题`）及真实Word `TOC \o "1-3" \h \z`域。
4. 保留前置材料与正文分节；原模板`upperRoman`是源模板状态。符合4.3.7.1的生成规则为正文前小写罗马页码、正文/附录阿拉伯页码，每页唯一；首页不能因`titlePg`被无检查地隐藏。

## 总体方案一级章名

仅在选择`overall-technical-solution`模板时使用。

1. `1 概述`
2. `2 使用与技术指标要求`
3. `3 总体架构`
4. `4 分系统设计`
5. `5 系统集成设计`
6. `6 关键技术分析`
7. `7 效能分析`
8. `8 工程组织管理`
9. `9 初步工作计划`

## 总体方案二级结构

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

## 常用表格结构

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

## 旧业务内容辨识

模板含有以下旧业务示例。 仅当确认为遗留示例时移除，不从模板复制为当前业务；若已由当前来源证实则以来源为准：

- `ZBZQ`
- `XX数据中台`
- `联合XX数据资源体系`
- `XX数据融合治理`
