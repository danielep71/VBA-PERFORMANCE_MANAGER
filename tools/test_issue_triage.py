"""Fixtures for the issue triage audit (#86)."""
from __future__ import annotations

import unittest
from typing import Any

import issue_triage as it

MILESTONE = {"title": "v1.4.1", "state": "open", "due_on": "2026-11-30T08:00:00Z"}
PLANNING = "## Dependencies\n\n- None\n\n## Schedule\n\n- Start date: 2026-11-03\n- Target date: 2026-11-27\n"


def issue(number: int = 1, body: str = PLANNING, **overrides: Any) -> dict[str, Any]:
    item: dict[str, Any] = {
        "number": number, "state": "open", "body": body,
        "assignees": [{"login": "danielep71"}],
        "labels": [{"name": "bug"}, {"name": "P2"}],
        "milestone": dict(MILESTONE),
    }
    item.update(overrides)
    return item


def findings(*items: dict[str, Any]) -> list[str]:
    return it.audit(list(items))


class TriageRuleTests(unittest.TestCase):
    def test_a_complete_issue_passes(self) -> None:
        self.assertEqual(findings(issue()), [])

    def test_assignee_is_required(self) -> None:
        self.assertEqual(findings(issue(assignees=[{"login": "someone"}])),
                         ["#1: not assigned to danielep71"])

    def test_exactly_one_priority(self) -> None:
        self.assertIn("exactly one priority", findings(issue(labels=[{"name": "bug"}]))[0])
        two = issue(labels=[{"name": "P1"}, {"name": "P3"}])
        self.assertEqual(findings(two), ["#1: needs exactly one priority label, has ['P1', 'P3']"])

    def test_open_milestone_is_required(self) -> None:
        self.assertEqual(findings(issue(milestone=None)), ["#1: no milestone"])
        closed = dict(MILESTONE, state="closed")
        self.assertEqual(findings(issue(milestone=closed)), ["#1: milestone v1.4.1 is closed"])

    def test_target_date_rules(self) -> None:
        cases = {
            PLANNING.replace("2026-11-27", "2026-11-30"): [],   # on the due date
            PLANNING.replace("2026-11-27", "2026-12-01"):
                ["#1: target date 2026-12-01 is after the milestone due date 2026-11-30"],
            PLANNING.replace("2026-11-27", "27/11/2026"):
                ["#1: target date '27/11/2026' is not YYYY-MM-DD"],
            PLANNING.replace("2026-11-27", "2026-02-30"):
                ["#1: target date 2026-02-30 is not a valid date"],
            PLANNING.replace("- Target date: 2026-11-27\n", ""): ["#1: no target date"],
            PLANNING.replace("2026-11-03", "2026-11-28"):
                ["#1: start date 2026-11-28 is after the target date 2026-11-27"],
            PLANNING.replace("Target date", "Target end date"): [],
        }
        for body, expected in cases.items():
            with self.subTest(body=body):
                self.assertEqual(findings(issue(body=body)), expected)

    def test_issue_form_fields_are_read(self) -> None:
        form = ("### Summary\n\nText\n\n### Dependencies\n\n#2\n\n"
                "### Target date\n\n2026-11-20\n")
        self.assertEqual(findings(issue(body=form), issue(2)), [])
        empty = form.replace("2026-11-20", "_No response_")
        self.assertEqual(findings(issue(body=empty), issue(2)), ["#1: no target date"])

    def test_planning_line_wins_over_the_form_field(self) -> None:
        body = PLANNING + "\n### Target date\n\n2027-05-01\n"
        self.assertEqual(findings(issue(body=body)), [])

    def test_dependency_section_rules(self) -> None:
        def with_dependencies(lines: str) -> str:
            return PLANNING.replace("- None\n", lines)
        self.assertEqual(findings(issue(body=PLANNING.split("## Schedule")[1])),
                         ["#1: no Dependencies section"])
        self.assertEqual(findings(issue(body=with_dependencies(""))),
                         ["#1: the Dependencies section is empty; write None if there are none"])
        ok = with_dependencies("- #2 (alignment seams)\n- External condition: an Excel run\n")
        self.assertEqual(findings(issue(body=ok), issue(2)), [])
        self.assertEqual(findings(issue(body=with_dependencies("- #99\n"))),
                         ["#1: dependency #99 does not exist"])
        self.assertEqual(findings(issue(body=with_dependencies("- wait for the release\n"))),
                         ["#1: dependency line 'wait for the release' names no #issue, None or "
                          "External condition"])

    def test_pull_requests_and_closed_issues_are_not_audited(self) -> None:
        self.assertEqual(findings(issue(1, body="", pull_request={"url": "x"}),
                                  issue(2, body="", state="closed")), [])

    def test_a_pull_request_can_be_a_dependency(self) -> None:
        body = PLANNING.replace("- None\n", "- #5\n")
        self.assertEqual(findings(issue(body=body), issue(5, pull_request={"url": "x"})), [])


class FetchTests(unittest.TestCase):
    def test_fetch_follows_pagination(self) -> None:
        pages = {1: [issue(n) for n in range(1, 101)], 2: [issue(101)]}
        seen: list[str] = []

        def get(path: str) -> Any:
            seen.append(path)
            return pages[int(path.rsplit("page=", 1)[1])]

        self.assertEqual(len(it.fetch(get, "o/r")), 101)
        self.assertEqual(seen, ["/repos/o/r/issues?state=all&per_page=100&page=1",
                                "/repos/o/r/issues?state=all&per_page=100&page=2"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
