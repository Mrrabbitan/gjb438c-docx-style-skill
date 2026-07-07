# Template Structure

Use this file when planning or restructuring document content. The first-level chapter names are fixed. The second-level headings are the main formatting boundary.

## Fixed First-Level Chapters

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
- Interface table: `序号`, `接口名称`, `发送方`, `接收方`, `接口描述`, `接口协议`
- Hardware environment table: `序号`, `名称`, `型号规格`, `计量单位`, `数量`
- Software environment table: `序号`, `软件类型`, `软件项名称`, `版本`, `数量`
- Organization table: `序号`, `单位名称`, `项目分工`, `负责人`, `备注`
- Schedule table: `序号`, `任务名称`, `开始时间`, `完成时间`

## Legacy Content Guardrail

The template contains old example business terms. Do not reuse them as content:

- `ZBZQ`
- `XX数据中台`
- `联合XX数据资源体系`
- `XX数据融合治理`

