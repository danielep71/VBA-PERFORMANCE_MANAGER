"""Issue forms and governance workflows (#86).

Read with the strict YAML loader of tools/workflow_controls.py; without
PyYAML the tests are skipped locally and fail in GitHub Actions.
"""
from __future__ import annotations

import json
import os
import re
import unittest
from pathlib import Path
from typing import Any

try:
    import workflow_controls as wc
except ImportError:  # pragma: no cover - exercised only without the parser
    wc = None  # type: ignore[assignment]

ROOT = Path(__file__).resolve().parent.parent
FORMS = ROOT / ".github" / "ISSUE_TEMPLATE"
ELEMENTS = {"markdown", "textarea", "input", "dropdown", "checkboxes"}
PRIORITY = re.compile(r"P[123]")


def load(path: Path) -> Any:
    assert wc is not None
    return wc.load(path.read_text(encoding="utf-8"))


def declared_labels() -> set[str]:
    data = json.loads((ROOT / "tools" / "labels.json").read_text(encoding="utf-8"))
    return {label["name"] for label in data["labels"]}


@unittest.skipIf(wc is None and not os.environ.get("GITHUB_ACTIONS"),
                 "PyYAML is not installed: see tools/requirements-ci.txt")
class IssueFormTests(unittest.TestCase):
    def forms(self) -> dict[str, Any]:
        return {path.name: load(path) for path in sorted(FORMS.glob("*.yml"))
                if path.name != "config.yml"}

    def test_only_issue_forms_remain(self) -> None:
        self.assertEqual(sorted(path.name for path in FORMS.glob("*.md")), [])
        self.assertEqual(sorted(self.forms()), ["bug_report.yml", "documentation.yml",
                                                "feature_request.yml"])
        self.assertIs(load(FORMS / "config.yml")["blank_issues_enabled"], False)

    def test_forms_set_triage_metadata(self) -> None:
        labels = declared_labels()
        for name, form in self.forms().items():
            with self.subTest(form=name):
                self.assertTrue(form["name"] and form["description"])
                self.assertEqual(form["assignees"], ["danielep71"])
                self.assertTrue(set(form["labels"]) <= labels, form["labels"])
                self.assertEqual(len([label for label in form["labels"]
                                      if PRIORITY.fullmatch(label)]), 1)

    def test_form_elements_are_valid(self) -> None:
        for name, form in self.forms().items():
            ids = [element.get("id") for element in form["body"] if element["type"] != "markdown"]
            with self.subTest(form=name):
                self.assertEqual(len(ids), len(set(ids)), "duplicate element id")
                self.assertTrue(all(ids), "every input element needs an id")
                for element in form["body"]:
                    self.assertIn(element["type"], ELEMENTS)
                    attributes = element["attributes"]
                    if element["type"] != "markdown":
                        self.assertTrue(attributes.get("label"))
                    if element["type"] == "dropdown":
                        self.assertTrue(attributes["options"])

    def test_forms_capture_dependencies_and_target_date(self) -> None:
        for name, form in self.forms().items():
            elements = {element.get("id"): element for element in form["body"]}
            with self.subTest(form=name):
                self.assertEqual(elements["dependencies"]["attributes"]["label"], "Dependencies")
                self.assertIs(elements["dependencies"]["validations"]["required"], True)
                self.assertEqual(elements["target-date"]["type"], "input")
                self.assertEqual(elements["target-date"]["attributes"]["label"], "Target date")


@unittest.skipIf(wc is None and not os.environ.get("GITHUB_ACTIONS"),
                 "PyYAML is not installed: see tools/requirements-ci.txt")
class GovernanceWorkflowTests(unittest.TestCase):
    def test_milestone_workflow_never_runs_pull_request_code(self) -> None:
        document = load(ROOT / ".github/workflows/pr-milestones.yml")
        self.assertEqual(document["permissions"], {})
        self.assertEqual(document["on"]["pull_request_target"]["types"],
                         ["opened", "reopened", "edited", "milestoned"])
        job = document["jobs"]["clear"]
        self.assertEqual(job["permissions"], {"pull-requests": "write"})
        self.assertEqual(len(job["steps"]), 1)
        step = job["steps"][0]
        self.assertNotIn("uses", step)
        self.assertNotIn("${{", step["run"])

    def test_label_sync_writes_only_from_the_default_branch(self) -> None:
        job = load(ROOT / ".github/workflows/labels-sync.yml")["jobs"]["sync"]
        self.assertEqual(job["if"], "github.ref_name == github.event.repository.default_branch")
        self.assertIn("--apply", job["steps"][-1]["run"])
        drift = load(ROOT / ".github/workflows/labels-drift.yml")
        self.assertNotIn("push", drift["on"])
        self.assertEqual(drift["jobs"]["drift"]["permissions"], {"contents": "read", "issues": "read"})
        self.assertIn("--check", drift["jobs"]["drift"]["steps"][-1]["run"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
