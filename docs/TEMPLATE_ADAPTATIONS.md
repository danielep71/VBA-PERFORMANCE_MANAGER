# Template adaptations

This repository adopts practices from the maintainer's own project template.
The template is licensed MPL-2.0, and some of its files are MIT; this
repository is licensed Apache-2.0 with a [NOTICE](../NOTICE).

Each adaptation is written in this repository's own terms: its `src/`,
`test/`, `demo/` and `tools/` layout and its `cPM_` / `M_cPM_` names. No
template names, issue numbers or history are copied. Before a template file is
adapted, its licence is decided and recorded here.

| File in this repository | Template source | Template revision | Licence decision | Issue |
|---|---|---|---|---|
| `tools/workflow_controls.py` | `tests/test_workflow_controls.py`, `tests/strict_yaml.py` | `91a57f2` | Rewritten, not copied: the rules and the duplicate-key loader are reimplemented as a checker with fixtures. Apache-2.0 under this repository's licence. | #80 |
| `tools/test_workflow_controls.py` | `tests/test_workflow_controls.py` | `91a57f2` | Rewritten, not copied: new fixtures for this repository's rules. Apache-2.0. | #80 |
| `tools/requirements-ci.txt` | `.github/requirements.txt` | `91a57f2` | Not copied: the PyYAML 6.0.3 hashes were taken from the PyPI JSON API. A hash list carries no licensable expression. | #80 |
| `pyproject.toml` | `pyproject.toml` | `91a57f2` | Adapted: the same Ruff rule set and strict mypy settings are written for `tools/`. Configuration values carry no licensable expression. Apache-2.0. | #81 |
| `tools/requirements-quality.txt` | `tools/requirements-quality-ci.txt` | `91a57f2` | Not copied: the versions match the template, but every hash was taken from the PyPI JSON API, with `types-PyYAML` added. Apache-2.0. | #81 |
| `.github/workflows/static-checks.yml` (quality steps) | `.github/workflows/repository-integrity.yml` | `91a57f2` | Practices only: checksum-verified actionlint, hash-locked installs and unittest discovery. The actionlint checksum was taken from the upstream release. Apache-2.0. | #81 |
| `tools/binary_review.py`, `tools/test_binary_review.py` | `tools/check_binaries.py` | `91a57f2` | Rewritten, not copied: the same detection rule (text attribute unset, or a NUL byte in the first 8000 bytes). The approved paths, reasons and stale-entry rules are this repository's own. Apache-2.0. | #84 |
| `tools/binary-allowlist.json` | `.github/binary-allowlist.json` | `91a57f2` | Not copied: this repository's own entries and schema, adding a reason per file. Apache-2.0. | #84 |
| `tools/markdown_links.py`, `tools/test_markdown_links.py` | `tools/check_links.py` | `91a57f2` | Rewritten, not copied: link extraction, GitHub heading-slug rules and the external mode are reimplemented for this repository. Apache-2.0. | #85 |
| `.github/dependabot.yml` | `.github/dependabot.yml` | `91a57f2` | Adapted: the same weekly grouped layout, written for this repository's `tools/` requirements, labels and release branch. Configuration only. Apache-2.0. | #85 |
| `.github/workflows/scorecard.yml` | `.github/workflows/scorecard.yml` | `91a57f2` | Adapted: the four-step job is the shape the Scorecard service requires. Pins re-verified upstream. Apache-2.0. | #85 |
| `.github/workflows/link-check.yml` | `.github/workflows/link-check.yml` | `91a57f2` | Adapted to `tools/markdown_links.py` and the single Python pin. Apache-2.0. | #85 |
| `tools/test_supply_chain.py` | `tests/test_workflow_controls.py` (Dependabot and Scorecard tests) | `91a57f2` | Rewritten for this repository's configuration. Apache-2.0. | #85 |
| `.github/ISSUE_TEMPLATE/bug_report.yml`, `feature_request.yml` | `.github/ISSUE_TEMPLATE/bug.yml`, `enhancement.yml` | `91a57f2` | Not copied: converted from this repository's own Markdown templates. The template contributed the idea of default priority, assignee, dependency and target-date fields. Apache-2.0. | #86 |
| `tools/labels.json`, `tools/labels.py`, `tools/test_labels.py` | `.github/labels.json`, `.github/scripts/labels-sync.mjs`, `labels-drift.mjs` | `91a57f2` | Rewritten, not copied: this repository's 17 labels, and a Python sync/drift tool instead of Node scripts. Labels are never deleted. Apache-2.0. | #86 |
| `.github/workflows/labels-sync.yml`, `labels-drift.yml`, `pr-milestones.yml` | `.github/workflows/labels-sync.yml`, `labels-drift.yml`, `pr-milestones.yml` | `91a57f2` | Adapted: the same triggers and least-privilege shape, rewritten for `tools/labels.py` and a single API step. Apache-2.0. | #86 |
| `tools/issue_triage.py`, `tools/test_issue_triage.py`, `.github/workflows/issue-triage.yml` | `tools/check_issues.py`, `.github/workflows/issue-triage.yml` | `91a57f2` | Rewritten, not copied: these rules read this repository's Schedule and Dependencies conventions. Read-only. Apache-2.0. | #86 |
| `tools/repository-settings.json`, `tools/repository_settings.py`, `tools/test_repository_settings.py` | `docs/setup/repository-settings.json`, `tools/setup_repository.py` | `91a57f2` | Rewritten, not copied: the baseline records this repository's live settings, with the security features enabled; the preview, apply, verify and evidence modes are its own. Apache-2.0. | #86 |
| `tools/readme_facts.py`, `tools/test_readme_facts.py`, `tools/gate_checks.py`, `tools/certified-evidence.json` | `tools/render_docs.py` | `91a57f2` | Rewritten, not copied: the idea of marked, checked Markdown blocks. The marker syntax, renderers, the split between derived and certified tag-bound facts, and the gate registry are this repository's own. Apache-2.0. | #89 |
| `tools/vba_source_rules.py`, `tools/test_vba_source_rules.py` | `tools/check_vba_source.py`, `tools/check_vba_public_api.py`, `tools/check_vba_conditionals.py` | `91a57f2` | Rewritten, not copied: the rule list is the template's idea, but the rules run on this repository's conditional-compilation model (`tools/vba_compile_safety.py`), compare `VB_Name` case-insensitively and check `PtrSafe` in both directions. Apache-2.0. | #83 |
| `tools/vba_format.py`, `tools/test_vba_format.py` | `tools/format_vba.py` | `91a57f2` | Rewritten, not copied: the rules were measured from this repository's source (indented module-level statements, one or two blank lines after `End`) rather than taken from the template's house style. Apache-2.0. | #83 |
| `.github/workflows/static-checks.yml` (hardening) | `.github/workflows/repository-integrity.yml` | `91a57f2` | Practices only (pins, fixed runner, timeouts, `persist-credentials: false`, minimum-Python job); the workflow is this repository's own. Apache-2.0. | #80 |
