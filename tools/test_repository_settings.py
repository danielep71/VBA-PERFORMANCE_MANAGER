"""Fixtures for the repository settings baseline (#86)."""
from __future__ import annotations

import copy
import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import repository_settings as rs

REPO = "danielep71/VBA-PERFORMANCE_MANAGER"
BASELINE = rs.BASELINE.read_text(encoding="utf-8")


class FakeGitHub:
    """A repository whose settings start as today's and change through the API."""

    def __init__(self, *, readable_alerts: bool = True, readable_fixes: bool = True,
                 fixes_delay: int = 0, summary_fixes: bool = True) -> None:
        baseline = json.loads(BASELINE)
        self.general = dict(baseline["general"])
        self.analysis = {key: "disabled" for key in rs.ANALYSIS}
        self.alerts = False
        self.fixes = False
        self.fixes_delay = fixes_delay      # reads that still see the old state after a PUT
        self.reporting = False
        self.readable_alerts = readable_alerts
        self.readable_fixes = readable_fixes
        self.summary_fixes = summary_fixes  # whether security_and_analysis reports them
        self.rulesets = {7: {"id": 7, **copy.deepcopy(baseline["rulesets"][0])}}
        self.writes: list[tuple[str, str, Any]] = []

    def call(self, method: str, endpoint: str, payload: Any = None) -> tuple[int, Any]:
        base = f"repos/{REPO}"
        if method != "GET":
            self.writes.append((method, endpoint, copy.deepcopy(payload)))
        payload = copy.deepcopy(payload)
        if endpoint == base and method == "GET":
            summary = {key: {"status": value} for key, value in self.analysis.items()}
            if self.summary_fixes:
                summary[rs.DEPENDABOT] = {"status": "enabled" if self.fixes_visible(consume=False)
                                          else "disabled"}
            return 200, {**self.general, "security_and_analysis": summary}
        if endpoint == f"{base}/automated-security-fixes":
            if method == "GET":
                return (200, {"enabled": self.fixes_visible(), "paused": False}) \
                    if self.readable_fixes else (403, {"message": "forbidden"})
            self.fixes = method == "PUT"
            return 204, None
        if endpoint == base and method == "PATCH":
            analysis = payload.pop("security_and_analysis", {})
            self.general.update(payload)
            self.analysis.update({key: value["status"] for key, value in analysis.items()})
            return 200, {}
        if endpoint == f"{base}/vulnerability-alerts":
            if method == "GET":
                return (204 if self.alerts else 404) if self.readable_alerts else 403, None
            self.alerts = method == "PUT"
            return 204, None
        if endpoint == f"{base}/private-vulnerability-reporting":
            if method == "GET":
                return 200, {"enabled": self.reporting}
            self.reporting = method == "PUT"
            return 204, None
        if endpoint == f"{base}/rulesets" and method == "GET":
            return 200, [{"id": key} for key in self.rulesets]
        if endpoint == f"{base}/rulesets" and method == "POST":
            self.rulesets[99] = {"id": 99, **payload}
            return 201, {}
        if endpoint.startswith(f"{base}/rulesets/"):
            number = int(endpoint.rsplit("/", 1)[1])
            if method == "GET":
                return 200, copy.deepcopy(self.rulesets[number])
            self.rulesets[number] = {"id": number, **payload}
            return 200, {}
        return 404, {"message": "not found"}


    def fixes_visible(self, consume: bool = True) -> bool:
        """GitHub turns security updates on asynchronously: the first snapshots lag.

        Each snapshot reads the summary first and the dedicated endpoint last,
        so only the endpoint read counts one lagging snapshot.
        """
        if self.fixes and self.fixes_delay:
            if consume:
                self.fixes_delay -= 1
            return False
        return self.fixes


def no_wait(_: float) -> None:
    """Stand-in for time.sleep, so the re-read loop runs instantly."""


def apply(github: FakeGitHub, **options: object) -> tuple[int, list[str]]:
    return rs.run("apply", github.call, wait=no_wait, **options)  # type: ignore[arg-type]


class SettingsTests(unittest.TestCase):
    def test_baseline_file_is_valid_and_enables_security(self) -> None:
        baseline, problems = rs.load_baseline(BASELINE)
        self.assertEqual(problems, [])
        self.assertEqual(baseline["repository"], REPO)
        self.assertTrue(all(baseline["security"].values()))
        self.assertEqual([r["name"] for r in baseline["rulesets"]], ["Protect main"])

    def test_preview_lists_differences_and_changes_nothing(self) -> None:
        github = FakeGitHub()
        status, out = rs.run("preview", github.call)
        self.assertEqual(status, 0)
        self.assertEqual(github.writes, [])
        self.assertEqual([line for line in out if line.startswith("would call")], [
            f"would call: PUT repos/{REPO}/vulnerability-alerts",
            f"would call: PUT repos/{REPO}/automated-security-fixes",
            f"would call: PATCH repos/{REPO}",
            f"would call: PUT repos/{REPO}/private-vulnerability-reporting"])
        self.assertIn("5 setting(s) differ from the baseline", out)

    def test_apply_converges_and_verify_then_passes(self) -> None:
        github = FakeGitHub()
        status, out = apply(github)
        self.assertEqual(status, 0, out)
        self.assertEqual([w[:2] for w in github.writes[:2]], [
            ("PUT", f"repos/{REPO}/vulnerability-alerts"),
            ("PUT", f"repos/{REPO}/automated-security-fixes")])
        self.assertIn("verified with 1 read(s)", out)
        patch = github.writes[2][2]
        self.assertEqual(patch, {"security_and_analysis": {key: {"status": "enabled"}
                                                           for key in rs.ANALYSIS}})
        self.assertEqual(rs.run("verify", github.call), (0, ["settings match the baseline"]))

    def test_verify_fails_on_any_difference(self) -> None:
        status, out = rs.run("verify", FakeGitHub().call)
        self.assertEqual(status, 1)
        self.assertIn("differs: security.secret_scanning: expected true, live false", out)

    def test_unreadable_setting_is_a_difference(self) -> None:
        github = FakeGitHub(readable_alerts=False)
        status, out = apply(github)
        self.assertEqual(status, 1)
        self.assertIn("verified with 5 read(s)", out)
        status, out = rs.run("verify", github.call)
        self.assertEqual(status, 1)
        self.assertIn('differs: security.vulnerability_alerts: expected true, live "unreadable"', out)

    def test_general_and_ruleset_drift_are_restored(self) -> None:
        github = FakeGitHub()
        github.general["delete_branch_on_merge"] = False
        github.rulesets[7]["rules"] = list(reversed(github.rulesets[7]["rules"]))   # order only
        github.rulesets[7]["enforcement"] = "disabled"
        apply(github)
        self.assertIn(("PUT", f"repos/{REPO}/rulesets/7"), [w[:2] for w in github.writes])
        self.assertEqual(rs.run("verify", github.call)[0], 0)

    def test_missing_ruleset_is_created(self) -> None:
        github = FakeGitHub()
        github.rulesets.clear()
        apply(github)
        self.assertIn(("POST", f"repos/{REPO}/rulesets"), [w[:2] for w in github.writes])

    def test_default_branch_is_never_patched(self) -> None:
        github = FakeGitHub()
        github.general["default_branch"] = "develop"
        status, _ = apply(github)
        self.assertEqual(status, 1)
        self.assertFalse(any("default_branch" in (w[2] or {}) for w in github.writes))

    def test_evidence_record_is_dated_and_complete(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            now = datetime(2026, 11, 3, 9, 30, tzinfo=timezone.utc)
            _, out = rs.run("verify", FakeGitHub().call, evidence=Path(td), now=now)
            path = Path(td) / "repository-settings-20261103T093000Z.json"
            self.assertIn(f"evidence written to {path}", out)
            record = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(record["checked_at"], "2026-11-03T09:30:00Z")
        self.assertEqual(record["result"], "differs")
        self.assertEqual(len(record["differences"]), 5)
        self.assertEqual(len(record["baseline_sha256"]), 64)

    def test_a_late_setting_passes_after_a_re_read(self) -> None:
        """The #103 case: security updates read as off just after the apply."""
        github = FakeGitHub(fixes_delay=2)
        waits: list[float] = []
        status, out = rs.run("apply", github.call, wait=waits.append)
        self.assertEqual(status, 0, out)
        self.assertEqual(waits, [rs.PAUSE, rs.PAUSE])
        self.assertIn("re-read 1 of 5: 1 setting(s) still differ; waiting 3 s", out)
        self.assertIn("verified with 3 read(s)", out)

    def test_verify_reads_once(self) -> None:
        github = FakeGitHub(fixes_delay=1)
        apply(github, attempts=1)
        status, out = rs.run("verify", github.call, wait=self.fail)   # never waits
        self.assertEqual(status, 0, out)

    def test_dependabot_read_falls_back_then_is_unreadable(self) -> None:
        github = FakeGitHub(readable_fixes=False)
        apply(github)
        self.assertEqual(rs.run("verify", github.call)[0], 0)    # from security_and_analysis
        github = FakeGitHub(readable_fixes=False, summary_fixes=False)
        status, out = apply(github)
        self.assertEqual(status, 1)
        self.assertIn('still differs: security.dependabot_security_updates: expected true, '
                      'live "unreadable"', out)

    def test_disabling_turns_security_updates_off_before_alerts(self) -> None:
        github = FakeGitHub()
        apply(github)
        baseline = json.loads(BASELINE)
        baseline["security"].update(vulnerability_alerts=False, dependabot_security_updates=False)
        github.writes.clear()
        apply(github, baseline_text=json.dumps(baseline))
        self.assertEqual([w[:2] for w in github.writes[:2]], [
            ("DELETE", f"repos/{REPO}/automated-security-fixes"),
            ("DELETE", f"repos/{REPO}/vulnerability-alerts")])

    def test_invalid_baseline_fails(self) -> None:
        self.assertEqual(rs.run("verify", FakeGitHub().call, baseline_text="{")[0], 1)
        broken = json.loads(BASELINE)
        broken["rulesets"][0].pop("bypass_actors")
        status, out = rs.run("verify", FakeGitHub().call, baseline_text=json.dumps(broken))
        self.assertEqual(status, 1)
        self.assertIn("each ruleset needs exactly", out[0])


if __name__ == "__main__":
    unittest.main(verbosity=2)
