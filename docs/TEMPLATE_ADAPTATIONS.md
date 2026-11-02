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
| `.github/workflows/static-checks.yml` (hardening) | `.github/workflows/repository-integrity.yml` | `91a57f2` | Practices only (pins, fixed runner, timeouts, `persist-credentials: false`, minimum-Python job); the workflow is this repository's own. Apache-2.0. | #80 |
