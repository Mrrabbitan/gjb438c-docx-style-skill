# GJB 438C-2021 Requirements Matrix

This file is the skill's requirement cross-check against
`资源文档/GJB 438C-2021 军用软件开发文档通用要求.pdf`.

The PDF available in this workspace is a scanned image PDF. The facts below
were manually checked from rendered pages, especially the table of contents and
pages covering chapters 4 and 5. Do not invent unverified clause text. When a
task requires a document-specific appendix not listed here, inspect the
corresponding PDF pages first.

## Authority Order

1. GJB 438C-2021 general requirements are the upper-level constraint for
   document kind, document composition, page numbering, clauses, appendices, and
   tailoring.
2. The bundled DOCX template is the implementation baseline for Word styles,
   section structure, cover table, TOC field, footer fields, and generated DOCX
   mechanics.
3. Project-specific templates such as the overall technical solution's nine
   chapters are allowed only when the user requests that project template.

## Clause-to-Skill Matrix

| Source | Requirement summary | Skill rule |
|---|---|---|
| 4.1 General | Documents may use paper or electronic media. Electronic document formats may include database format, WPS/document-processor-compatible format, or formats stored in software engineering tools. | DOCX generation uses a Word/WPS-compatible `.docx` package. Database-style data must be represented through tables or exported records, not hidden local files. |
| 4.2 Document kinds | The standard identifies 20 principal software life-cycle document kinds. | The skill must recognize all 20 abbreviations and must not force every document into the overall technical solution's nine-chapter structure. |
| 4.3.1 Document composition | A document normally consists of cover, modification page, table of contents, body, and appendices. | Generated/audited documents must preserve or warn about these composition elements. Missing modification pages are warnings unless the user requests strict GJB composition audit. |
| 4.3.2 Cover | Cover information may include document number, version/revision/volume number, classification, document name, system/software identifier, preparing organization, author, reviewer, countersignature, approver, and date. | Preserve the template cover table and cover styles. Fill available metadata in place. Warn when expected cover concepts are absent. |
| 4.3.3 Modification page | Modification page records modification history, including reason, content, version, and date. | Add or preserve a modification-history section/table when source data exists. Audit warns when absent. |
| 4.3.4 Table of contents | TOC includes chapters, clauses, appendices, and page numbers; when needed it may include lists for figures, tables, notes, or other references. | Keep a real Word TOC field. Do not fake TOC with plain text. Add figure/table lists only when requested or present in the source. |
| 4.3.5 Body | Body is the document's main content; chapter-specific detailed requirements are in chapter 5 and appendices. | Body structure must match the detected document type. Overall technical solution uses the bundled nine-chapter framework; standard GJB document kinds use the corresponding appendix outline. |
| 4.3.6 Appendices | Appendices hold separately maintainable information such as figures, tables, classified data, and should be cited in body text; appendices are marked by capital letters such as A, B. | Preserve appendices, use appendix labels, and keep appendix content referenced from the body when source material provides it. |
| 4.3.7.1 Page numbers / page marks | Each page has a unique page number. Front matter uses roman numbering; body and appendices use Arabic numbering. Multi-volume documents may restart numbering per volume. | Preserve section breaks and native PAGE fields. The bundled template currently implements front matter with upper-roman numbering (`I`, `II`, `III`) and body with Arabic `— PAGE —`. |
| 4.3.7.2 Clauses and subclauses | Clauses/subclauses are hierarchically numbered; X and Y in appendix examples are variables to be replaced by actual content. | Headings use template multilevel numbering. Do not type number prefixes into heading text because Word renders numbers automatically. |
| 4.3.7.3 Representation | Content may be represented with figures, tables, or other forms when that improves readability. | Figures use `145图样式` plus `Caption`; tables use `145表头` and `145表正文`; captions use chapter-based numbering. |
| 4.4 Tailoring | Documents may be combined, split, or tailored according to lifecycle model, contract, and actual activities. Omitted chapters/clauses should be marked as no content and explained. | Do not silently delete required sections. When tailoring, mark omitted sections as `本章无内容` or `本条无内容` and include a reason. |
| 5.1-5.20 Detailed requirements | Each document kind has a detailed purpose and corresponding appendix body format. | Use the document-type index below to map abbreviations to appendix outlines and writing prompts. |

## GJB 438C Document Type Index / 文档种类索引

| Abbr. | Chinese document name | English expansion | Appendix | Skill handling |
|---|---|---|---|---|
| SDP | 软件开发计划 | software development plan | A | Standard GJB document type |
| SIP | 软件安装计划 | software installation plan | B | Standard GJB document type |
| STrP | 软件移交计划 | software transition plan | C | Standard GJB document type |
| STP | 软件测试计划 | software test plan | D | Standard GJB document type |
| OCD | 运行方案说明 | operational concept description | E | Standard GJB document type |
| SSS | 系统/子系统规格说明 | system/subsystem specification | F | Standard GJB document type |
| IRS | 接口需求规格说明 | interface requirement specification | G | Standard GJB document type |
| SSDD | 系统/子系统设计说明 | system/subsystem design description | H | Standard GJB document type |
| IDD | 接口设计说明 | interface design description | I | Standard GJB document type |
| SRS | 软件需求规格说明 | software requirement specification | J | Standard GJB document type |
| SDD | 软件设计说明 | software design description | K | Standard GJB document type |
| DBDD | 数据库设计说明 | database design description | L | Standard GJB document type |
| STD | 软件测试说明 | software test description | M | Standard GJB document type |
| STR | 软件测试报告 | software test report | N | Standard GJB document type |
| SPS | 软件产品规格说明 | software product specification | O | Standard GJB document type |
| SVD | 软件版本说明 | software version description | P | Standard GJB document type |
| SUM | 软件用户手册 | software user manual | Q | Standard GJB document type |
| CPM | 计算机编程手册 | computer programming manual | R | Standard GJB document type |
| FSM | 固件保障手册 | firmware support manual | S | Standard GJB document type |
| SDSR | 软件研制总结报告 | software development summary report | T | Standard GJB document type |

Machine-readable slash list for validation:
`SDP/SIP/STrP/STP/OCD/SSS/IRS/SSDD/IDD/SRS/SDD/DBDD/STD/STR/SPS/SVD/SUM/CPM/FSM/SDSR`.

## Document-Type Decision Rule

- Use `overall-technical-solution` only when the user asks for the bundled
  project overall technical solution template or the existing nine-chapter
  project format.
- Use the standard document type (`SDP`, `SRS`, `SDD`, etc.) when the user asks
  for a named GJB 438C life-cycle document.
- Use style-only repair when the user asks to preserve an existing document's
  semantics and repair format only. In this mode, do not rewrite content into
  another document kind unless explicitly requested.

## Audit Severity Defaults

- Errors: broken DOCX package, missing required template styles, heading
  double-numbering, body text using `Normal`, table styles not matching strict
  table mode, broken section/page fields, local path leakage, or forbidden
  generated styles.
- Warnings: missing modification-history page, caption numbering gaps, missing
  cover concepts not available in the template, absent appendix when source
  material has none, or non-strict table mismatch.
