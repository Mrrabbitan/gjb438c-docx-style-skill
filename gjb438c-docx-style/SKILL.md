---
name: gjb438c-docx-style
description: Generate, reformat, repair, or audit Chinese DOCX technical方案 and GJB 438C-2021 software life-cycle documents. Use when the user mentions GJB438C, GJB-438C, GJB 438C-2021, SDP, SRS, SDD, STP, STD, STR, 总体技术方案模板, 1-2总体技术方案模板, 固化格式, 按模板格式调整, 二级标题对齐, or asks to produce/reformat a DOCX using the provided GJB438C structure and styles.
---

# GJB438C DOCX Style

## Purpose

Use this skill to generate, reformat, repair, or audit `.docx` documents so they follow GJB 438C-2021 general document requirements and the bundled GJB-438C-style Word template mechanics.

Treat GJB 438C-2021 as the upper-level constraint for document kind, composition, cover, modification page, table of contents, body, appendices, page numbering, clause numbering, figure/table expression, and tailoring. Treat the bundled template as the DOCX implementation baseline for styles, cover table, TOC field, section break, page fields, and overall technical solution formatting. Do not reuse legacy business content from the template.

## Required Workflow

1. Identify the task type:
   - New document from source materials.
   - Existing DOCX reformatting.
   - Style-only repair.
   - Compliance audit.
2. Identify the document type before restructuring content:
   - Use `overall-technical-solution` only for the bundled nine-chapter project technical solution template.
   - Use a standard GJB 438C type when requested: `SDP`, `SIP`, `STrP`, `STP`, `OCD`, `SSS`, `IRS`, `SSDD`, `IDD`, `SRS`, `SDD`, `DBDD`, `STD`, `STR`, `SPS`, `SVD`, `SUM`, `CPM`, `FSM`, or `SDSR`.
   - Use style-only repair when the user asks to preserve the current document semantics and only fix formatting.
3. Read the minimum necessary references:
   - For GJB 438C-2021 general requirements and document-kind routing, read `references/gjb438c-2021-requirements.md`.
   - For chapter layout, read `references/template-structure.md`.
   - For exact paragraph, heading, table, and caption rules, read `references/format-lock.md`.
   - For reusable Chinese prompts, read `references/writing-prompts.md`.
4. Use the bundled template at `assets/1-2总体技术方案-模板参考.docx` for DOCX style mechanics unless the user explicitly provides another approved template.
5. Apply the correct structure:
   - Overall technical solution: keep the template's nine first-level chapters fixed and align content to the second-level framework.
   - Standard GJB 438C document: follow the corresponding appendix format; do not force the nine-chapter overall technical solution structure.
   - Style-only repair: preserve semantic order and only normalize styles, fields, numbering, captions, and tables.
6. Never reuse legacy business text from the template. Remove or reject old placeholders such as `ZBZQ`, `XX数据中台`, and `联合XX数据资源体系`.
7. Validate the output with `scripts/audit_gjb438c_docx.py` before delivery, including page-number, footer, section-break, field-refresh, table-style, and GJB composition checks.

## Generation Rules

- For `overall-technical-solution`, keep first-level headings exactly as the template's nine chapters.
- For standard GJB 438C document types, use the document-kind appendix outline from the standard and keep missing-but-required clauses as `本章无内容` or `本条无内容` with a reason when tailoring is needed.
- Headings are auto-numbered by the template's multilevel list. Heading paragraph text must NOT contain typed number prefixes (write `概述`, never `1 概述`); typed numbers cause double numbering.
- Use second-level headings as the main formatting boundary. Each second-level section should have a consistent block pattern: heading, body paragraphs, optional figures, optional tables, optional captions, and optional interface/resource descriptions.
- Use third-level and lower headings only inside the corresponding second-level section. Do not let lower-level headings change the second-level structure.
- Use only template styles for visible document content:
  - `Heading 1` to `Heading 5` for headings (no typed numbers in text).
  - `145正文` for normal body paragraphs.
  - `145图样式` for figure placeholder paragraphs (each figure on its own paragraph, followed by a `Caption`).
  - `Caption` for table and figure captions, with chapter-based numbering `图 X-Y` / `表X-Y`.
  - `145表头` for table header cells.
  - `145表正文` for table body cells.
  - `311-目录标题`, `toc 1`, `toc 2`, and `toc 3` for table-of-contents content.
  - `编号密级`, `文头字`, `文件标识号`, `文件名称`, `单位名称` for cover lines (inside the template cover table).
- Never use `DB表头` or `DB表正文`: they are defined in the template style sheet but never used in the document body.
- Preserve or repair the GJB document composition: cover, modification page/history when available, table of contents, body, and appendices when applicable.
- Preserve the template front matter: the cover table, the real `TOC \o "1-3" \h \z` field, and the section break between front matter and body (roman page numbers, body headers/footers).
- Preserve page-number fields and footer relationships exactly as template-owned structure. The front-matter section uses upper-roman page numbers starting at 1, and the body section keeps the template body footer with a native Word `PAGE` field rendered in the `— PAGE —` pattern.
- When generating or reformatting a document, replace content only from the first `Heading 1` onward. Do not rebuild the cover, TOC, section break, page-number fields, header/footer relationships, or `w:updateFields` setting.
- Avoid introducing new document styles, ad hoc fonts, manual spacing, or environment-specific absolute paths.

## Script Usage

Create a minimal sample document:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --demo --output /tmp/gjb438c-demo.docx
```

Create a standard GJB document-type sample without forcing the nine-chapter overall technical solution:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --demo --document-type SDP --output /tmp/gjb438c-sdp-demo.docx
```

Generate from a structured content JSON file:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --content-json content.json --document-type SRS --output output.docx
```

Audit a generated document:

```bash
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py output.docx --strict-overall --strict-secondary
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
    "document_type": "overall-technical-solution",
    "title": "总体技术方案",
    "number": "XXXX",
    "identifier": "SJZT-XXXX-NNNN-ZF【YYYY/MM/DD】",
    "unit": "技术总师组",
    "date": "2026年7月"
  },
  "sections": [
    {
      "level": 1,
      "title": "概述",
      "paragraphs": ["本章说明项目任务依据、编制目的、指导原则和建设目标。"]
    },
    {
      "level": 2,
      "title": "任务依据",
      "paragraphs": ["本节说明任务来源、合同依据、需求依据和相关标准依据。"],
      "figures": [
        {"caption": "图 1-1 任务依据关系示意图"}
      ],
      "tables": [
        {
          "caption": "表1-1 任务依据清单",
          "headers": ["序号", "依据名称", "说明"],
          "rows": [["1", "需求文件", "说明需求来源"]]
        }
      ]
    }
  ]
}
```

The script applies styles by `level`, not by guessed visual appearance. Section titles must not contain typed number prefixes; if present they are stripped automatically because the template headings auto-number. `metadata` fields fill the preserved template cover table in place; omit a field to keep the template's placeholder text.
