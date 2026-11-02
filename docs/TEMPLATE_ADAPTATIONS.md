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
| `.github/workflows/static-checks.yml` (hardening) | `.github/workflows/repository-integrity.yml` | `91a57f2` | Practices only (pins, fixed runner, timeouts, `persist-credentials: false`, minimum-Python job); the workflow is this repository's own. Apache-2.0. | #80 |
