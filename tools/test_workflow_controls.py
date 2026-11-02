"""Fixtures for the workflow hardening rules (#80).

Run with ``python tools/test_workflow_controls.py``. The rules need PyYAML,
pinned with hashes in tools/requirements-ci.txt:

    python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-ci.txt

Without PyYAML the tests are skipped locally; in GitHub Actions they fail.
"""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    import yaml

    import workflow_controls as wc
except ImportError:  # pragma: no cover - exercised only without the parser
    yaml = None  # type: ignore[assignment]

SHA = "3d3c42e5aac5ba805825da76410c181273ba90b1"
README_OK = "The tools need Python 3.11 or later.\n"

STATIC = f"""\
name: Static checks
on:
  push:
permissions:
  contents: read
jobs:
  vba-lint:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@{SHA} # v7.0.1
        with:
          persist-credentials: false
      - uses: actions/setup-python@{SHA} # v7.1.0
        with:
          python-version-file: .python-version
  minimum-python:
    runs-on: ubuntu-24.04
    timeout-minutes: 5
    steps:
      - uses: actions/checkout@{SHA} # v7.0.1
        with:
          persist-credentials: false
      - uses: actions/setup-python@{SHA} # v7.1.0
        with:
          python-version: "3.11.18"
"""

LABELS = f"""\
name: Labels
on:
  workflow_dispatch:
permissions:
  contents: read
jobs:
  sync:
    runs-on: ubuntu-24.04
    timeout-minutes: 5
    permissions:
      issues: write
    steps:
      - uses: actions/github-script@{SHA} # v9.0.0
"""


def run(static: str = STATIC, labels: str = LABELS, version: str | None = "3.13.18\n",
        readme: str | None = README_OK) -> list[str]:
    return wc.check({"static-checks.yml": static, "sync-label-colors.yml": labels},
                    version, readme)


@unittest.skipIf(yaml is None and not os.environ.get("GITHUB_ACTIONS"),
                 "PyYAML is not installed: see tools/requirements-ci.txt")
class WorkflowControlTests(unittest.TestCase):
    def assertRejects(self, problems: list[str], *fragments: str) -> None:
        self.assertTrue(problems, "expected a violation, got none")
        text = "\n".join(problems)
        for fragment in fragments:
            self.assertIn(fragment, text)

    def test_parser_is_available(self) -> None:
        self.assertIsNotNone(yaml, "install tools/requirements-ci.txt before running the tests")

    def test_repository_complies(self) -> None:
        self.assertEqual(wc.repository_problems(), [])

    def test_valid_fixture_complies(self) -> None:
        self.assertEqual(run(), [])

    def test_duplicate_keys_are_rejected(self) -> None:
        with self.assertRaisesRegex(yaml.YAMLError, "duplicate key"):
            wc.load("jobs:\n  a:\n    runs-on: x\n    runs-on: y\n")
        self.assertRejects(run(static=STATIC.replace("    timeout-minutes: 10\n",
                                                     "    timeout-minutes: 10\n    timeout-minutes: 10\n")),
                           "invalid YAML")

    def test_workflow_token_must_be_read_only(self) -> None:
        self.assertRejects(run(labels=LABELS.replace("  contents: read\n", "  issues: write\n", 1)),
                           "workflow-level token may write ['issues']")
        self.assertRejects(run(static=STATIC.replace("permissions:\n  contents: read\n", "")),
                           "workflow-level permissions are not explicit")

    def test_only_listed_jobs_write(self) -> None:
        widened = STATIC.replace("    timeout-minutes: 10\n",
                                 "    timeout-minutes: 10\n    permissions:\n      contents: write\n")
        self.assertRejects(run(static=widened), "job vba-lint: may write ['contents']")
        self.assertRejects(run(labels=LABELS.replace("      issues: write\n",
                                                     "      issues: write\n      contents: write\n")),
                           "job sync: may write ['contents']")

    def test_every_job_has_a_timeout(self) -> None:
        self.assertRejects(run(static=STATIC.replace("    timeout-minutes: 10\n", "")),
                           "job vba-lint: timeout-minutes must be a positive integer")
        for bad in ("0", "true", '"10"'):
            with self.subTest(timeout=bad):
                self.assertRejects(run(static=STATIC.replace("timeout-minutes: 10", f"timeout-minutes: {bad}")),
                                   "timeout-minutes must be a positive integer")

    def test_runner_image_is_fixed(self) -> None:
        for bad in ("ubuntu-latest", "windows-latest", "[self-hosted]", "${{ matrix.os }}"):
            with self.subTest(runner=bad):
                self.assertRejects(run(static=STATIC.replace("runs-on: ubuntu-24.04", f"runs-on: {bad}", 1)),
                                   "runs-on must be a fixed runner image")

    def test_checkout_drops_credentials(self) -> None:
        self.assertRejects(run(static=STATIC.replace("persist-credentials: false", "persist-credentials: true", 1)),
                           "job vba-lint: actions/checkout must set persist-credentials: false")
        missing = STATIC.replace("        with:\n          persist-credentials: false\n", "", 1)
        self.assertRejects(run(static=missing), "persist-credentials: false")

    def test_python_version_is_defined_once(self) -> None:
        self.assertRejects(run(static=STATIC.replace("python-version-file: .python-version",
                                                     'python-version: "3.12"')),
                           "job vba-lint: actions/setup-python must read python-version-file")
        for bad in ("3.13\n", "3.x\n", "", None):
            with self.subTest(version=bad):
                self.assertRejects(run(version=bad), ".python-version must hold one exact version")

    def test_minimum_python_job_is_exact_and_documented(self) -> None:
        self.assertRejects(run(static=STATIC.replace('python-version: "3.11.18"', 'python-version: "3.11"')),
                           "must pin the minimum Python as an exact python-version")
        self.assertRejects(run(static=STATIC.split("  minimum-python:")[0]),
                           "no minimum-python job pins the minimum Python")
        self.assertRejects(run(readme="The tools need Python 3.12 or later.\n"),
                           "README.md must state the minimum once as 'Python 3.11 or later'")
        self.assertRejects(run(readme=None), "'Python 3.11 or later'")


if __name__ == "__main__":
    unittest.main(verbosity=2)
