# Format Lock

Use this file when applying or auditing format. The document must be formatted by paragraph role, table role, and heading level.

## Paragraph Style Map

| Role | Required Style |
|---|---|
| Cover number/classification line | `编号密级` |
| TOC title | `311-目录标题` |
| TOC level 1 | `toc 1` |
| TOC level 2 | `toc 2` |
| TOC level 3 | `toc 3` |
| First-level heading | `Heading 1` |
| Second-level heading | `Heading 2` |
| Third-level heading | `Heading 3` |
| Fourth-level heading | `Heading 4` |
| Fifth-level heading | `Heading 5` |
| Normal body paragraph | `145正文` |
| Figure caption | `Caption` |
| Table caption | `Caption` |
| Table header cell paragraph | `DB表头` |
| Table body cell paragraph | `DB表正文` |

## Line-Level Rules

- Every visible paragraph must map to a role in the style map.
- Do not create new styles for project-specific content.
- Do not rely on manual font overrides where a template paragraph style exists.
- Do not introduce environment-specific absolute paths into generated documents or repository files.
- Blank spacing should come from the template styles, not manual empty paragraph stacks.
- If a paragraph starts with a first-level number such as `1 概述`, apply `Heading 1`.
- If a paragraph starts with a second-level number such as `3.2 逻辑架构设计`, apply `Heading 2`.
- If a paragraph starts with a third-level number such as `4.2.1 系统结构`, apply `Heading 3`.
- Use `Caption` for any line beginning with `图 `, `图`, `表 `, or `表` when it is a caption.
- Use `145正文` for explanatory text, requirement descriptions, assumptions, constraints, and implementation descriptions.

## Second-Level Alignment Rule

Second-level headings are the formatting control unit:

- Do not add arbitrary first-level chapters.
- Do not let third-level headings replace a required second-level section.
- Keep every table, caption, interface list, or resource list inside its owning second-level section.
- If source content has a different structure, remap it into the nearest existing second-level section.
- For project-specific分系统 content, use `4.x` second-level headings and keep each section's internal writing pattern consistent.

## Table Rules

- The first table row is always header content and uses `DB表头`.
- All subsequent rows use `DB表正文`.
- Captions should be separate paragraphs before the table and use `Caption`.
- Use template table schemas when possible:
  - Term: `编号`, `名称`, `说明`
  - Abbreviation: `序号`, `简写`, `全称`, `解释说明`
  - Interface: `序号`, `接口名称`, `发送方`, `接收方`, `接口描述`, `接口协议`
  - Hardware: `序号`, `名称`, `型号规格`, `计量单位`, `数量`
  - Software: `序号`, `软件类型`, `软件项名称`, `版本`, `数量`
  - Organization: `序号`, `单位名称`, `项目分工`, `负责人`, `备注`
  - Schedule: `序号`, `任务名称`, `开始时间`, `完成时间`

## Audit Failure Conditions

Treat the output as non-compliant when:

- Any first-level heading is missing, renamed, or reordered.
- A first-level, second-level, or third-level numbered heading uses the wrong style.
- Normal paragraphs use `Normal` instead of `145正文`.
- Table cell paragraphs do not use `DB表头` or `DB表正文`.
- Legacy template business terms remain in the final document.
- The document contains local machine paths such as `/Users/`, `/home/`, or `C:\Users\`.

