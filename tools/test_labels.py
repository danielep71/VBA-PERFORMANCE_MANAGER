"""Fixtures for labels as data (#86)."""
from __future__ import annotations

import json
import unittest
from typing import Any

import labels as lb

REPO = "owner/repo"


def label_file(*entries: dict[str, Any]) -> str:
    return json.dumps({"schema_version": 1, "labels": list(entries)})


def entry(name: str, color: str = "0052cc", description: str = "A label.") -> dict[str, Any]:
    return {"name": name, "color": color, "description": description}


class FakeGitHub:
    """Records calls and serves a label list, paginated like the REST API."""

    def __init__(self, live: list[dict[str, Any]]) -> None:
        self.live = live
        self.calls: list[tuple[str, str, Any]] = []

    def send(self, method: str, path: str, payload: dict[str, Any] | None) -> Any:
        self.calls.append((method, path, payload))
        if method == "GET":
            page = int(path.rsplit("page=", 1)[1])
            return self.live[(page - 1) * 100:page * 100]
        return {}


class FileTests(unittest.TestCase):
    def test_repository_file_is_valid(self) -> None:
        declared, problems = lb.parse(lb.LABELS_FILE.read_text(encoding="utf-8"))
        self.assertEqual(problems, [])
        names = {label.name for label in declared}
        self.assertTrue({"P1", "P2", "P3", "bug", "enhancement", "documentation", "CI"} <= names)

    def test_invalid_files_fail(self) -> None:
        cases = {
            "{": "invalid JSON",
            '{"labels": []}': "expected",
            label_file(entry("a"), entry("A")): "A is declared twice",
            label_file(entry("a", color="0052CC")): "six lowercase hex digits",
            label_file(entry("a", color="#0052cc")): "six lowercase hex digits",
            label_file(entry("a", description=" ")): "a description is required",
            label_file(entry("a", description="x" * 101)): "exceeds 100 characters",
            label_file({"name": "a", "color": "0052cc"}): "exactly name, color and description",
            label_file({**entry("a"), "extra": 1}): "exactly name, color and description",
        }
        for text, fragment in cases.items():
            with self.subTest(text=text):
                self.assertIn(fragment, "\n".join(lb.parse(text)[1]))


class SyncTests(unittest.TestCase):
    DECLARED = [lb.Label("bug", "d73a4a", "Defect."), lb.Label("good first issue", "7057ff", "Easy."),
                lb.Label("P1", "b60205", "Critical.")]

    def test_differences(self) -> None:
        live = [lb.Label("bug", "d73a4a", "Defect."), lb.Label("P1", "000000", "Critical."),
                lb.Label("stale", "ffffff", "Old.")]
        create, update, extra = lb.differences(self.DECLARED, live)
        self.assertEqual([label.name for label in create], ["good first issue"])
        self.assertEqual([label.name for label in update], ["P1"])
        self.assertEqual(extra, ["stale"])

    def test_apply_creates_and_updates_but_never_deletes(self) -> None:
        github = FakeGitHub([{"name": "bug", "color": "D73A4A", "description": "Defect."},
                             {"name": "P1", "color": "000000", "description": "Critical."},
                             {"name": "stale", "color": "ffffff", "description": None}])
        report = lb.apply(self.DECLARED, github.send, REPO)
        writes = [(method, path) for method, path, _ in github.calls if method != "GET"]
        self.assertEqual(writes, [("POST", f"/repos/{REPO}/labels"),
                                  ("PATCH", f"/repos/{REPO}/labels/P1")])
        self.assertEqual(report, ["created good first issue", "updated P1",
                                  "not declared, left in place: stale"])

    def test_names_with_spaces_are_encoded(self) -> None:
        github = FakeGitHub([{"name": "good first issue", "color": "000000", "description": "Easy."}])
        lb.apply([lb.Label("good first issue", "7057ff", "Easy.")], github.send, REPO)
        self.assertIn(("PATCH", f"/repos/{REPO}/labels/good%20first%20issue"),
                      [(m, p) for m, p, _ in github.calls])

    def test_drift_reports_every_difference_and_passes_when_equal(self) -> None:
        github = FakeGitHub([{"name": "bug", "color": "d73a4a", "description": "Changed."},
                             {"name": "stale", "color": "ffffff", "description": ""}])
        self.assertEqual(lb.drift(self.DECLARED, github.send, REPO), [
            "missing on GitHub: good first issue", "missing on GitHub: P1",
            "differs on GitHub: bug", "on GitHub but not in tools/labels.json: stale"])
        same = FakeGitHub([{"name": label.name, "color": label.color, "description": label.description}
                           for label in self.DECLARED])
        self.assertEqual(lb.drift(self.DECLARED, same.send, REPO), [])
        self.assertTrue(all(method == "GET" for method, _, _ in same.calls))

    def test_live_labels_follow_pagination(self) -> None:
        github = FakeGitHub([{"name": f"l{i}", "color": "000000", "description": "d"} for i in range(150)])
        self.assertEqual(len(lb.live_labels(github.send, REPO)), 150)
        self.assertEqual([p for _, p, _ in github.calls],
                         [f"/repos/{REPO}/labels?per_page=100&page=1",
                          f"/repos/{REPO}/labels?per_page=100&page=2"])

    def test_live_modes_need_a_token(self) -> None:
        import os
        saved = {key: os.environ.pop(key, None) for key in ("GITHUB_TOKEN", "GH_TOKEN")}
        try:
            self.assertEqual(lb.main(["--check"]), 1)
        finally:
            os.environ.update({key: value for key, value in saved.items() if value is not None})


if __name__ == "__main__":
    unittest.main(verbosity=2)
