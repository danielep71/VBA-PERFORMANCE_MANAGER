"""Dependabot and OpenSSF Scorecard configuration invariants (#85).

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


def load(rel: str) -> Any:
    assert wc is not None
    return wc.load((ROOT / rel).read_text(encoding="utf-8"))


@unittest.skipIf(wc is None and not os.environ.get("GITHUB_ACTIONS"),
                 "PyYAML is not installed: see tools/requirements-ci.txt")
class DependabotTests(unittest.TestCase):
    def test_covers_actions_and_every_locked_requirements_file(self) -> None:
        updates = load(".github/dependabot.yml")["updates"]
        covered = {(u["package-ecosystem"], u["directory"]) for u in updates}
        self.assertEqual(covered, {("github-actions", "/"), ("pip", "/tools")})
        locked = sorted(p.name for p in (ROOT / "tools").glob("requirements*.txt"))
        self.assertEqual(locked, ["requirements-ci.txt", "requirements-quality.txt"])
        for name in locked:
            self.assertIn("--hash=sha256:", (ROOT / "tools" / name).read_text(encoding="utf-8"))

    def test_updates_are_grouped_scheduled_and_triaged(self) -> None:
        declared = {label["name"] for label in
                    json.loads((ROOT / "tools/labels.json").read_text(encoding="utf-8"))["labels"]}
        for update in load(".github/dependabot.yml")["updates"]:
            with self.subTest(ecosystem=update["package-ecosystem"]):
                self.assertEqual(update["schedule"]["interval"], "weekly")
                self.assertTrue(update["groups"])
                self.assertRegex(update["target-branch"], r"^release/v\d+\.\d+\.\d+$")
                self.assertEqual(update["assignees"], ["danielep71"])
                priorities = [label for label in update["labels"] if re.fullmatch(r"P[123]", label)]
                self.assertEqual(len(priorities), 1)
                for label in update["labels"]:
                    self.assertIn(label, declared, f"label {label} is not in tools/labels.json")


@unittest.skipIf(wc is None and not os.environ.get("GITHUB_ACTIONS"),
                 "PyYAML is not installed: see tools/requirements-ci.txt")
class ScorecardTests(unittest.TestCase):
    def test_publishes_only_from_the_default_branch_with_approved_steps(self) -> None:
        document = load(".github/workflows/scorecard.yml")
        self.assertNotIn("pull_request", document["on"])
        self.assertIn("schedule", document["on"])
        job = document["jobs"]["analysis"]
        self.assertIn("github.event.repository.private == false", job["if"])
        self.assertIn("github.ref_name == github.event.repository.default_branch", job["if"])
        self.assertTrue(all("uses" in step and "run" not in step for step in job["steps"]))
        self.assertEqual([step["uses"].split("@")[0] for step in job["steps"]],
                         ["actions/checkout", "ossf/scorecard-action", "actions/upload-artifact",
                          "github/codeql-action/upload-sarif"])
        self.assertIs(job["steps"][1]["with"]["publish_results"], True)

    def test_link_check_runs_weekly_and_never_on_pull_requests(self) -> None:
        document = load(".github/workflows/link-check.yml")
        self.assertEqual(set(document["on"]), {"schedule", "workflow_dispatch"})
        self.assertIn("--external", document["jobs"]["external"]["steps"][-1]["run"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
