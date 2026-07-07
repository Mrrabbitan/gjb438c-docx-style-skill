---
name: gjb438c-docx-style
description: Generate, reformat, repair, or audit Chinese DOCX technical方案 documents against the bundled GJB-438C-style overall technical solution template. Use when the user mentions GJB438C, GJB-438C, 总体技术方案模板, 1-2总体技术方案模板, 固化格式, 按模板格式调整, 二级标题对齐, or asks to produce/reformat a DOCX using the provided GJB438C structure and styles.
---

# GJB438C DOCX Style

## Purpose

Use this skill to generate or reformat `.docx` documents so they follow the bundled GJB-438C-style overall technical solution template. Treat the template as the authority for structure and formatting, not as a source of reusable business content.

## Required Workflow

1. Identify the task type:
   - New document from source materials.
   - Existing DOCX reformatting.
   - Style-only repair.
   - Compliance audit.
2. Read the minimum necessary references:
   - For chapter layout, read `references/template-structure.md`.
   - For exact paragraph, heading, table, and caption rules, read `references/format-lock.md`.
   - For reusable Chinese prompts, read `references/writing-prompts.md`.
3. Use the bundled template at `assets/1-2总体技术方案-模板参考.docx`.
4. Keep the template's nine first-level chapters fixed. Align all generated content to the second-level chapter framework.
5. Never reuse legacy business text from the template. Remove or reject old placeholders such as `ZBZQ`, `XX数据中台`, and `联合XX数据资源体系`.
6. Validate the output with `scripts/audit_gjb438c_docx.py` before delivery.

## Generation Rules

- Keep first-level headings exactly as the template's nine chapters.
- Use second-level headings as the main formatting boundary. Each second-level section should have a consistent block pattern: heading, body paragraphs, optional tables, optional captions, and optional interface/resource descriptions.
- Use third-level and lower headings only inside the corresponding second-level section. Do not let lower-level headings change the second-level structure.
- Use only template styles for visible document content:
  - `Heading 1` to `Heading 5` for headings.
  - `145正文` for normal body paragraphs.
  - `Caption` for table and figure captions.
  - `DB表头` for table header cells.
  - `DB表正文` for table body cells.
  - `311-目录标题`, `toc 1`, `toc 2`, and `toc 3` for table-of-contents placeholders.
  - `编号密级` for cover number/classification lines.
- Avoid introducing new document styles, ad hoc fonts, manual spacing, or environment-specific absolute paths.

## Script Usage

Create a minimal sample document:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --demo --output /tmp/gjb438c-demo.docx
```

Generate from a structured content JSON file:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --content-json content.json --output output.docx
```

Audit a generated document:

```bash
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py output.docx --strict-secondary
```

Validate the skill bundle itself:

```bash
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py --skill-dir gjb438c-docx-style --validate-skill
```

## Content JSON Shape

Use this minimal shape when automation is helpful:

```json
{
  "metadata": {
    "title": "项目总体技术方案",
    "subtitle": "GJB438C模板版"
  },
  "sections": [
    {
      "level": 1,
      "title": "1 概述",
      "paragraphs": ["本章说明项目任务依据、编制目的、指导原则和建设目标。"]
    },
    {
      "level": 2,
      "title": "1.1 任务依据",
      "paragraphs": ["本节说明任务来源、合同依据、需求依据和相关标准依据。"],
      "tables": [
        {
          "caption": "表 1 任务依据清单",
          "headers": ["序号", "依据名称", "说明"],
          "rows": [["1", "需求文件", "说明需求来源"]]
        }
      ]
    }
  ]
}
```

The script applies styles by `level`, not by guessed visual appearance.
