# GJB438C DOCX Style Skill

This repository packages a reusable Codex Skill for generating, reformatting, repairing, and auditing Chinese `.docx` documents against a bundled GJB-438C-style overall technical solution template.

## What It Provides

- A Codex Skill: `gjb438c-docx-style`
- A bundled DOCX template asset
- Prompt references for generation, reformatting, style repair, and audit
- Cross-platform Python scripts for DOCX generation and validation
- A local installer for copying or linking the skill into a Codex skills directory

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
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py build/gjb438c-demo.docx --strict-secondary --strict-tables
```

## Skill Usage

After installation, ask Codex to use:

```text
Use $gjb438c-docx-style to reformat this DOCX into the GJB438C overall technical solution template and audit the result.
```

The first-level chapters are fixed to the template's nine-chapter structure. The second-level heading framework is the primary formatting boundary.

