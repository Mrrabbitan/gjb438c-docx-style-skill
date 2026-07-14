# Format Lock

Use this file when applying or auditing format. The document must be formatted by paragraph role, table role, and heading level. GJB 438C-2021 is the upper-level rule for document composition and numbering; the facts below were measured from the bundled template (`assets/1-2总体技术方案-模板参考.docx`) at XML level and are the DOCX implementation baseline.

## GJB 438C General Composition Rules

Every generated or repaired document should be evaluated against the GJB 438C-2021 composition model:

1. Cover: include document number, version/revision/volume number when available, classification, document name, system/software identifier, preparing organization, author/reviewer/countersignature/approver when available, and date.
2. Modification page: preserve or add modification history when source data exists; record reason, content, version/revision, date, and responsible party where available.
3. Table of contents: use a real Word TOC field, not static text. TOC should include chapters, clauses, appendices, and page numbers. Add figure/table/note lists only when requested or present in source material.
4. Body: structure follows the detected document type. `overall-technical-solution` uses the nine-chapter template; standard GJB documents use their corresponding appendix outline.
5. Appendices: preserve separately maintainable figures, tables, classified data, or detailed lists as appendices when source material contains them. Appendix identifiers use capital letters such as A and B and should be referenced from the body.
6. Tailoring: do not silently delete required chapters or clauses. If a section is omitted by lifecycle, contract, or actual project activity, mark it as `本章无内容` or `本条无内容` and state the reason.

## Paragraph Style Map

| Role | Required Style |
|---|---|
| Cover number/classification line | `编号密级` |
| Cover document-class line (项目技术文件) | `文头字` |
| Cover document identifier line (SJZT-XXXX-NNNN-ZF【YYYY/MM/DD】) | `文件标识号` |
| Cover document title (总体技术方案) | `文件名称` |
| Cover issuing unit and date | `单位名称` |
| TOC title (目 录) | `311-目录标题` |
| TOC level 1 | `toc 1` |
| TOC level 2 | `toc 2` |
| TOC level 3 | `toc 3` |
| First-level heading | `Heading 1` |
| Second-level heading | `Heading 2` |
| Third-level heading | `Heading 3` |
| Fourth-level heading | `Heading 4` |
| Fifth-level heading | `Heading 5` |
| Normal body paragraph | `145正文` |
| Figure placeholder paragraph (holds the image) | `145图样式` |
| Figure caption | `Caption` |
| Table caption | `Caption` |
| Table header cell paragraph | `145表头` |
| Table body cell paragraph | `145表正文` |

Forbidden styles for generated content: `DB表头` and `DB表正文` are defined in the template style sheet but are never used in the document body. Do NOT apply them; they exist only for template completeness. Also do not use `Normal` for visible body text.

## Heading Auto-Numbering Rule (critical)

`Heading 1` through `Heading 5` are bound to the multilevel list `numId=1` (`%1`, `%1.%2`, `%1.%2.%3`, ...). Word renders the chapter numbers automatically.

- Heading paragraph TEXT must NOT contain a typed number prefix. Write `概述`, not `1 概述`; write `逻辑架构设计`, not `3.2 逻辑架构设计`.
- Writing a typed number produces double numbering ("1 1 概述") and is an audit failure.
- Heading level is decided by semantic level of the section, never by parsing numbers out of the text.

## Measured Style Facts (line-level audit baseline)

| Style | East Asian font | Size | Key paragraph format |
|---|---|---|---|
| `Heading 1` | 黑体 | 14pt (sz 28) | keepNext, keepLines, line 560 exact, after 120, multilevel ilvl 0 |
| `Heading 2` | 黑体 | 14pt | same, multilevel ilvl 1 |
| `Heading 3` | 黑体 | 14pt | same, multilevel ilvl 2 |
| `Heading 4` | 宋体 bold | 14pt | same, multilevel ilvl 3 |
| `Heading 5` | 宋体 bold | 14pt | same, multilevel ilvl 4 |
| `145正文` | 宋体 (ASCII Times New Roman) | 14pt | first-line indent 2 chars, line 560 exact, justified |
| `145图样式` | inherits 145正文 | 14pt | based on 145正文, centered, no first-line indent, keepNext, single line spacing |
| `Caption` | 黑体 | 12pt (sz 24) | centered, beforeLines 20, afterLines 20 |
| `145表头` | 黑体 | 12pt | centered (based on 145表正文) |
| `145表正文` | 宋体 | 12pt | centered, beforeLines 20, afterLines 20 |
| `编号密级` | 黑体 | 14pt | centered, 1.5x-class spacing |
| `311-目录标题` | 黑体 | 16pt (sz 32) | centered, before 240, after 360 |
| `toc 1` | 宋体 | 12pt | line 440 exact |

Page setup: A4 (11906 x 16838 twips). Two sections: the front-matter section uses roman page numbers implemented by the bundled template as `upperRoman` starting at 1 with `titlePg`; the body section has its own header and footer references and Arabic page-number fields. Do not remove or merge the section break.

## Page Number and Footer Rules

The template page-number structure is part of the GJB438C format lock and must be preserved at OOXML level.

| Area | Required structure |
|---|---|
| Page size | Every section uses A4 page size, `w:pgSz w="11906" h="16838"` |
| Section count | Exactly two `w:sectPr` blocks: front matter and body |
| Front-matter page numbering | First section has `w:pgNumType fmt="upperRoman" start="1"` |
| Front-matter first page behavior | First section has `w:titlePg`; the title page can have no visible page number |
| Front-matter default footer | Default footer contains a native Word `PAGE` field with instruction text containing `PAGE` and centered paragraph alignment |
| Body header/footer ownership | Body section preserves template header/footer references and must not be merged into the front-matter section |
| Body page numbering | Body page number is a native Word `PAGE` field in the body footer; cached field result text such as `22` is not authoritative |
| Body page-number visual pattern | Body footer keeps the template `— PAGE —` pattern: a leading em dash text run, a `PAGE` field, and a trailing em dash text run |
| Field refresh | `word/settings.xml` includes `w:updateFields w:val="true"` so Word/WPS can refresh TOC and page-number fields on open |

Implementation notes:

- Never hand-type visible page numbers in normal paragraphs or footer text as the source of truth.
- Never replace the `PAGE` field with cached text, even if the cached value looks correct.
- For documents split into multiple volumes, restart numbering per volume only when the source contract or volume metadata requires it.
- When generating from structured content, replace body content only from the first `Heading 1` onward and leave the cover, TOC, section break, header references, footer references, and page-number fields intact.
- If a renderer shows stale page numbers, refresh fields in Word/WPS; do not alter the OOXML field structure merely to match a cached display value.

## Line-Level Rules

- Every visible paragraph must map to a role in the style map.
- Do not create new styles for project-specific content.
- Do not rely on manual font overrides where a template paragraph style exists.
- Do not introduce environment-specific absolute paths into generated documents or repository files.
- Blank spacing should come from the template styles, not manual empty paragraph stacks.
- Heading paragraphs never carry typed numbers (see Heading Auto-Numbering Rule).
- Every figure occupies its own `145图样式` paragraph, immediately followed by a `Caption` paragraph.
- Use `Caption` for caption lines. Caption numbering is chapter-based: `图 X-Y 名称` for figures and `表X-Y 名称` (or `表 X-Y 名称`) for tables, where X is the chapter number and Y is the per-chapter sequence.
- Use `145正文` for explanatory text, requirement descriptions, assumptions, constraints, and implementation descriptions.
- Inline enumerations inside body text follow the template pattern `（1）`, `（2）`, ... as plain `145正文` paragraphs, not list styles.

## Second-Level Alignment Rule

Second-level headings are the formatting control unit:

- Do not add arbitrary first-level chapters.
- Do not let third-level headings replace a required second-level section.
- Keep every table, caption, interface list, or resource list inside its owning second-level section.
- If source content has a different structure, remap it into the nearest existing second-level section.
- For project-specific分系统 content, use `4.x` second-level headings and keep each section's internal writing pattern consistent.

## Table Rules

- The first table row is always header content and uses `145表头`.
- All subsequent rows use `145表正文`.
- Captions should be separate paragraphs before the table and use `Caption`, numbered `表X-Y`.
- Table style in the template is `Normal Table` or `Table Grid1`; generated tables should use `Table Grid1` so borders render.
- Use template table schemas when possible:
  - Term: `编号`, `名称`, `说明`
  - Abbreviation: `序号`, `简写`, `全称`, `解释说明`
  - Standard/spec list: `序号`, `分类`, `名称`, `遵循修订`, `拓展新撰`
  - Interface: `序号`, `接口名称`, `发送方`, `接收方`, `接口描述`, `接口协议`
  - Hardware: `序号`, `名称`, `型号规格`, `计量单位`, `数量`
  - Software: `序号`, `软件类型`, `软件项名称`, `版本`, `数量`
  - Test data: `分类`, `数据类型`, `数据表`
  - Organization: `序号`, `单位名称`, `项目分工`, `负责人`, `备注`
  - Schedule: `序号`, `任务名称`, `开始时间`, `完成时间`

## Front Matter Rules

- The cover is a table (not free paragraphs) whose cells use the cover styles listed in the style map. Keep the template's cover table and replace only the cell texts.
- If the target document has a modification-history page/table, place it after the cover and before the TOC. If the source lacks modification history, warn during strict GJB composition audit rather than inventing entries.
- The table of contents is a real Word `TOC \o "1-3" \h \z` field preceded by a `311-目录标题` paragraph. Never fake the TOC with plain text; keep the field so Word/WPS can refresh it.
- The TOC should cover chapters, clauses, and appendices. Figure/table lists are optional and should be added only when requested or when source material already contains them.
- Keep the section break between front matter and body so roman page numbering and body headers/footers survive.

## Document Type and Tailoring Rules

- Overall technical solution documents use the nine fixed first-level chapters in `template-structure.md`.
- Standard GJB 438C document types (`SDP`, `SIP`, `STrP`, `STP`, `OCD`, `SSS`, `IRS`, `SSDD`, `IDD`, `SRS`, `SDD`, `DBDD`, `STD`, `STR`, `SPS`, `SVD`, `SUM`, `CPM`, `FSM`, `SDSR`) use the corresponding appendix outline listed in `gjb438c-2021-requirements.md`.
- Strict nine-chapter checking is valid only for `overall-technical-solution`; do not apply it to standard GJB documents unless the user explicitly asks to use the project template.
- When combining or splitting documents, preserve cover, modification history, TOC, body, appendix, page-number, and clause-number elements for each resulting document.

## Audit Failure Conditions

Treat the output as non-compliant when:

- In `overall-technical-solution` mode, any first-level heading is missing, renamed, or reordered (compare by unnumbered chapter names).
- A standard GJB document is incorrectly forced into the overall technical solution's nine-chapter structure.
- Any heading paragraph text starts with a typed number prefix such as `1 ` or `3.2 ` (double-numbering).
- A numbered-level heading uses the wrong `Heading N` style for its semantic level.
- Normal paragraphs use `Normal` instead of `145正文`.
- Table cell paragraphs do not use `145表头` (header row) or `145表正文` (body rows).
- Figure placeholder paragraphs do not use `145图样式`.
- Legacy template business terms remain in the final document.
- The document contains local machine paths such as macOS user-home paths, Linux home-directory paths, or Windows user-profile paths.
