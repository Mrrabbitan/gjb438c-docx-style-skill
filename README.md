# GJB438C DOCX Style Skill

This repository packages a reusable Codex Skill for generating, reformatting,
repairing, and auditing Chinese `.docx` documents against GJB 438C-2021 general
document requirements and a bundled GJB-438C-style Word template.

## What It Provides

- A Codex Skill: `gjb438c-docx-style`
- A bundled DOCX template asset
- A GJB 438C-2021 requirements matrix for document-kind routing and common
  composition checks
- Prompt references for generation, reformatting, style repair, and audit
- Cross-platform Python scripts for DOCX generation and validation
- A local installer for copying or linking the skill into a Codex skills directory

The skill supports the bundled `overall-technical-solution` template and the
standard GJB 438C software life-cycle document types:

`SDP/SIP/STrP/STP/OCD/SSS/IRS/SSDD/IDD/SRS/SDD/DBDD/STD/STR/SPS/SVD/SUM/CPM/FSM/SDSR`

The repository does **not** include the original GJB 438C-2021 PDF. The
requirements matrix records the checked rules needed by this skill and should be
reviewed against an authorized standard copy for formal compliance programs.

## Install

Install dependencies:

```bash
python -m pip install -e .
```

Install the skill into the default Codex skills directory:

```bash
python tools/install_skill.py
```

Install with overwrite:

```bash
python tools/install_skill.py --force
```

Install by symlink instead of copy:

```bash
python tools/install_skill.py --mode symlink --force
```

Install into a custom skills root:

```bash
python tools/install_skill.py --dest /path/to/skills --force
```

## Validate

Validate the skill bundle:

```bash
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py --skill-dir gjb438c-docx-style --validate-skill
```

Generate a sample document:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --demo --output build/gjb438c-demo.docx
```

Audit the generated sample:

```bash
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py build/gjb438c-demo.docx --strict-overall --strict-secondary --strict-tables
```

Generate and audit a standard GJB document-type sample without forcing the
nine-chapter overall technical solution template:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --demo --document-type SDP --output build/gjb438c-sdp-demo.docx
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py build/gjb438c-sdp-demo.docx --strict-tables
```

## Skill Usage

After installation, ask Codex to use:

```text
Use $gjb438c-docx-style to reformat this DOCX into the GJB438C overall technical solution template and audit the result.
```

For standard GJB 438C document types, specify the target type:

```text
Use $gjb438c-docx-style to create an SRS-style GJB 438C DOCX from these source materials and audit the result.
```

The nine-chapter first-level structure is strict only for the bundled overall
technical solution template. Standard GJB document types use their corresponding
document-type format and should not be forced into the project template.

## License

MIT
