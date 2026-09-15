# GJB438C DOCX Style Skill

Version **0.2.0** provides a Codex skill for drafting, repairing, and auditing
Chinese GJB 438C-2021 documents, with a dedicated software requirements
specification (SRS) workflow.

The [validation record](docs/srs-refinement-validation.md) describes the checked
standard clauses, resolved defects, regression tests, and rendered-page evidence.

## What the skill checks

The SRS route follows Appendix J's content structure and records requirement
identifiers, applicable conditions, qualification methods, source relationships,
tailoring decisions, and unresolved inputs. It supports two working modes:

- **Draft** preserves unresolved information and reports gaps for further work.
  The demo is a drafting example, not an approved project baseline.
- **Review** requires a complete content model and applies stricter checks before
  generating a document for human review. Review mode does not grant approval.

The standard governs applicable content and composition requirements. The P-09
template supplies implementation choices such as fonts, margins, and Word
styles. These choices must not be presented as provisions of the standard.
Under 4.3.7.1, the cover starts at lower-case Roman `i` and front matter
continues that sequence; body pages restart with Arabic `1`.

Machine checks cover only what their inputs and implementation can establish.
An unperformed check is **not checked**, never a pass. Source authenticity,
engineering suitability, field refresh, and visual inspection of every rendered
page remain explicit review work. Successful generation or an audit exit code
does not prove compliance, accepted performance, or review approval.

The existing `overall-technical-solution` template route remains available.
Other supported document types use their corresponding outlines and must not
be forced into the SRS or nine-chapter project outline:

`SDP/SIP/STrP/STP/OCD/SSS/IRS/SSDD/IDD/SRS/SDD/DBDD/STD/STR/SPS/SVD/SUM/CPM/FSM/SDSR`

## Install and update

From this repository checkout, install dependencies and the skill:

```bash
python -m pip install -e .
python tools/install_skill.py
```

The installer uses `$CODEX_HOME/skills` when configured, otherwise
`~/.codex/skills`. `--dest` accepts a skills directory or the exact
`gjb438c-docx-style` target path.

```bash
python tools/install_skill.py --dest /path/to/skills
python tools/install_skill.py --force
python tools/install_skill.py --mode symlink --force
```

Without `--force`, an existing installation is left intact. With `--force`, the
installer prepares and verifies a staged copy before replacing the target and
preserves the previous installation as a sibling `.backup-...` path. It rejects
overlapping source and target paths and does not follow the target's final
symbolic link when replacing it. A failed final replacement restores the
previous installation where possible. Machine caches and Word lock files are
excluded from copied bundles.

Before updating, compare any changes made in your installed skill with this
checkout. A backup protects the old files; the installer does not merge local
edits. To restore one, move the current target aside and rename the selected
backup to `gjb438c-docx-style`. Symlink installation keeps using this checkout,
so subsequent checkout changes affect the installed skill. A symlink backup
preserves the link, not a snapshot of its destination.

## Generate and audit an SRS

Generate a draft example:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --demo --document-type SRS --mode draft --output build/srs-draft.docx
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py build/srs-draft.docx --document-type SRS --json
```

Use a project content model following
`gjb438c-docx-style/references/srs-content-schema.json` for review:

```bash
python gjb438c-docx-style/scripts/apply_gjb438c_template.py --content-json content.json --document-type SRS --mode review --output build/srs-review.docx
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py build/srs-review.docx --document-type SRS --strict-srs --content-json content.json --json
```

Passing the same model to the audit enables checks against the supplied
requirement and source records. Follow the skill's review gate for evidence and
rendered-page inspection after these commands.

Example Codex request:

```text
Use $gjb438c-docx-style to prepare this software requirements specification in draft mode, identify missing source information, and audit the result against Appendix J. Separate standard requirements from project-template formatting.
```

## Validate the repository

```bash
python -m unittest discover -s tests -v
python gjb438c-docx-style/scripts/audit_gjb438c_docx.py --skill-dir gjb438c-docx-style --validate-skill
python tools/build_srs_fixture.py --output build/srs-review.json
```

The fixture builder creates synthetic review data for regression checks. Its
independent source baseline and interface definition are in
[`tests/fixtures/DEMO-SSS.md`](tests/fixtures/DEMO-SSS.md) and
[`tests/fixtures/DEMO-IRS.md`](tests/fixtures/DEMO-IRS.md); none is a real project
specification. CI runs tests, legacy overall-solution and SDP
examples, an SRS draft, an SRS review fixture with its model, and validation of a
temporary installation. The matrix covers Python 3.9 and 3.11 on Linux and
Windows, and Python 3.11 on macOS. CI does not perform human source verification
or full visual approval.

## Public template and source material

The bundled P-09 asset is a sanitized reusable template. Its former organization,
personal editing metadata, historical dates, and custom metadata have been
removed. Long embedded standard instructions have been replaced with short,
original drafting prompts. The cover uses unresolved placeholders, while
section structure, styles, numbering definitions, and native Word fields are
retained. Prompts and placeholders are not project requirements or facts.

This repository does **not** include the GJB 438C-2021 PDF, extracted standard
text, private project materials, or review scratch files. The references are
implementation aids and clause mappings; use an authorized standard copy to
verify formal interpretations. Generated DOCX files and other build output
belong under `build/` and are not source assets.

## License

MIT. The license does not grant rights in the external GJB standard.
