"""Run the traffic workflow's embedded scripts against isolated fixtures (#165).

The scripts of .github/workflows/daily-traffic.yml are extracted and executed
against local traffic responses, a local HTTP server and local Git
repositories; nothing reaches GitHub. Read with the strict YAML loader of
tools/workflow_controls.py; without PyYAML the tests are skipped locally and
fail in GitHub Actions.
"""
from __future__ import annotations

import csv
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

try:
    import workflow_controls as wc
except ImportError:  # pragma: no cover - exercised only without the parser
    wc = None  # type: ignore[assignment]

ROOT = Path(__file__).resolve().parent.parent
WORKFLOW = ROOT / ".github" / "workflows" / "daily-traffic.yml"


def jobs() -> dict[str, Any]:
    assert wc is not None
    document: dict[str, Any] = wc.load(WORKFLOW.read_text(encoding="utf-8"))
    found: dict[str, Any] = document["jobs"]
    return found


def step(name: str) -> dict[str, Any]:
    return next(s for s in jobs()["traffic"]["steps"] if s["name"] == name)


def script(name: str) -> str:
    return str(step(name)["run"])


def embedded(name: str) -> str:
    match = re.search(r"<< 'PY'\n(.*?)\nPY", script(name), re.S)
    assert match is not None
    return match[1]


@unittest.skipIf(wc is None and not os.environ.get("GITHUB_ACTIONS"),
                 "PyYAML is not installed: see tools/requirements-ci.txt")
class TrafficTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.work = self.root / "work"
        self.work.mkdir()
        self.env = {**os.environ, "PUSH_TOKEN": "fixture", "TRAFFIC_TOKEN": "fixture"}
        self.build = self.root / "build-traffic.py"
        self.build.write_text(embedded("Build CSV history").replace("/tmp/", f"{self.root}/"))
        self.snapshot("2026-11-07", clones=10)

    def run_script(self, source: str, *, python: bool = True) -> subprocess.CompletedProcess[str]:
        return subprocess.run([sys.executable, "-c", source] if python else ["bash", "-c", source],
                              cwd=self.work, env=self.env, capture_output=True, text=True,
                              check=False)

    def build_data(self) -> None:
        result = self.run_script(self.build.read_text())
        self.assertEqual(result.returncode, 0, result.stderr)

    def snapshot(self, day: str, *, clones: int = 10, views: int = 50, uniques: int = 2,
                 stars: int = 3, forks: int = 2, referrer: str = "") -> None:
        for metric, count in (("views", views), ("clones", clones)):
            (self.root / f"{metric}.json").write_text(json.dumps({"count": count, "uniques": uniques,
                metric: [{"timestamp": day + "T00:00:00Z", "count": count, "uniques": uniques}]}))
        (self.root / "repo.json").write_text(json.dumps({"stargazers_count": stars, "forks_count": forks}))
        (self.root / "referrers.json").write_text(json.dumps(
            [{"referrer": referrer, "count": 5, "uniques": 2}] if referrer else []))
        (self.root / "paths.json").write_text("[]")
        (self.root / "traffic-snapshot-time").write_text(day + "T12:00:00+00:00")

    def rows(self, name: str) -> list[dict[str, str]]:
        with (self.work / "data" / name).open(newline="") as stream:
            return list(csv.DictReader(stream))

    def badge(self, name: str) -> dict[str, Any]:
        data: dict[str, Any] = json.loads((self.work / "data/badges" / f"{name}.json").read_text())
        return data

    def test_history_and_badges_accumulate_across_days(self) -> None:
        self.build_data()
        self.snapshot("2026-11-08", clones=100, uniques=5)
        self.build_data()
        self.snapshot("2026-11-08", clones=100, uniques=5)   # a re-run on the same day
        self.build_data()
        self.assertEqual([r["snapshot_date"] for r in self.rows("traffic.csv")],
                         ["2026-11-07", "2026-11-08"])
        self.assertEqual(self.rows("traffic.csv")[-1]["clones_14d"], "100")
        self.assertEqual([r["clones"] for r in self.rows("traffic_daily.csv")], ["10", "100"])
        self.assertEqual(self.badge("clones"), {"schemaVersion": 1, "label": "clones",
                                                "message": "110", "color": "217346"})
        self.assertEqual(self.badge("visits")["message"], "7")       # 2 + 5 daily uniques
        self.assertEqual(self.badge("visitors")["label"], "visitors (14d)")
        self.assertEqual(self.badge("visitors")["message"], "5")     # today's 14-day uniques

    def test_large_counts_are_compacted(self) -> None:
        self.snapshot("2026-11-07", clones=1500)
        self.build_data()
        self.assertEqual(self.badge("clones")["message"], "1.5k")

    def test_bad_response_leaves_history_unchanged(self) -> None:
        self.build_data()
        before = {p: p.read_bytes() for p in (self.work / "data").rglob("*") if p.is_file()}
        self.snapshot("2026-11-08")
        (self.root / "paths.json").write_text('{"message": "failure"}')
        result = self.run_script(self.build.read_text())
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(before, {p: p.read_bytes() for p in before})

    def test_older_snapshot_cannot_replace_newer_remote_data(self) -> None:
        self.snapshot("2026-11-09", clones=500)
        self.build_data()
        before = (self.work / "data/traffic.csv").read_bytes()
        self.snapshot("2026-11-08", clones=20)
        self.build_data()
        self.assertEqual((self.work / "data/traffic.csv").read_bytes(), before)

    def git(self, *args: str, cwd: Path | None = None) -> str:
        result = subprocess.run(["git", *args], cwd=cwd or self.work, capture_output=True, text=True,
                                check=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout.strip()

    @unittest.skipUnless(shutil.which("bash") and shutil.which("git"), "requires Bash and Git")
    def test_rejected_push_rebuilds_on_concurrent_remote_history(self) -> None:
        remote = self.root / "remote.git"
        self.git("init", "--bare", str(remote))
        self.git("init", "-b", "traffic-history")
        self.git("config", "user.email", "fixture@example.invalid")
        self.git("config", "user.name", "Traffic fixture")
        self.build_data()
        self.git("add", ".")
        self.git("commit", "-m", "Initial snapshot")
        self.git("remote", "add", "origin", str(remote))
        self.git("push", "origin", "HEAD:traffic-history")
        # A second writer publishes after this checkout was prepared.
        other = self.root / "other"
        self.git("clone", "-b", "traffic-history", str(remote), str(other))
        (other / "concurrent.txt").write_text("Keep this concurrent update\n")
        with (other / "data/traffic_daily.csv").open("a") as stream:
            stream.write("2026-11-01,7,3,4,2\n")
        self.git("add", ".", cwd=other)
        self.git("-c", "user.name=Other", "-c", "user.email=other@example.invalid",
                 "commit", "-m", "Concurrent history", cwd=other)
        concurrent = self.git("rev-parse", "HEAD", cwd=other)
        self.git("push", "origin", "HEAD", cwd=other)
        self.snapshot("2026-11-08", clones=100)
        self.build_data()
        source = script("Commit traffic history").replace("/tmp/", f"{self.root}/")
        result = self.run_script(source, python=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("rejected", result.stderr)
        self.git("merge-base", "--is-ancestor", concurrent, "HEAD")
        self.assertTrue((self.work / "concurrent.txt").exists())
        self.assertEqual({r["date"] for r in self.rows("traffic_daily.csv")},
                         {"2026-11-01", "2026-11-07", "2026-11-08"})
        self.assertEqual(self.git("rev-parse", "HEAD"),
                         self.git("rev-parse", "traffic-history", cwd=remote))

        # Persistent rejection stops at three attempts and publishes nothing.
        hook = remote / "hooks/pre-receive"
        hook.write_text("#!/bin/sh\nexit 1\n")
        hook.chmod(0o755)
        accepted = self.git("rev-parse", "traffic-history", cwd=remote)
        self.snapshot("2026-11-09", clones=200)
        self.build_data()
        result = self.run_script(source, python=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(result.stderr.count("[remote rejected]"), 3)
        self.assertIn("failed after three attempts", result.stdout)
        self.assertEqual(self.git("rev-parse", "traffic-history", cwd=remote), accepted)

    @unittest.skipUnless(shutil.which("bash") and shutil.which("curl"), "requires Bash and curl")
    def test_http_retry_and_persistent_failure(self) -> None:
        statuses = [503, 200]
        requests: list[str] = []

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self) -> None:
                requests.append(self.path)
                self.send_response(statuses.pop(0) if statuses else 200)
                self.end_headers()
                self.wfile.write(b"{}")

            def log_message(self, format: str, *args: object) -> None:
                pass

        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.addCleanup(server.server_close)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.shutdown)
        source = next(str(step["run"]) for step in jobs()["traffic"]["steps"]
                      if "TRAFFIC_TOKEN" in step.get("env", {}))
        source = source.replace("${{ github.repository }}", "fixture/repo")
        source = source.replace("https://api.github.com/repos/$REPO",
                                f"http://127.0.0.1:{server.server_port}")
        source = source.replace("/tmp/", f"{self.root}/")
        for bound in ("--retry 5", "--retry-max-time 300", "--max-time 60", "--connect-timeout 15"):
            self.assertIn(bound, source)
        self.build_data()
        before = (self.work / "data/traffic.csv").read_bytes()
        result = self.run_script(source, python=False)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(requests), 6)  # five endpoints, first attempted twice
        requests.clear()
        statuses[:] = [401]
        result = self.run_script(source, python=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(len(requests), 1)
        self.assertEqual((self.work / "data/traffic.csv").read_bytes(), before)

    def test_job_credentials_and_permissions_are_scoped(self) -> None:
        job = jobs()["traffic"]
        self.assertEqual(list(jobs()), ["traffic"])
        self.assertEqual(job["permissions"], {"contents": "write"})
        self.assertEqual(job["environment"], "analytics")
        self.assertIn("github.event.repository.default_branch", job["if"])
        holders = {s["name"] for s in job["steps"] if "secrets." in repr(s)}
        self.assertEqual(holders, {"Check the traffic App configuration",
                                   "Create the traffic App token"})
        self.assertNotIn("secrets.TRAFFIC_APP_PRIVATE_KEY }}",
                         repr(step("Check the traffic App configuration")))
        token = {s["name"] for s in job["steps"] if "steps.app-token" in repr(s)}
        self.assertEqual(token, {"Fetch GitHub traffic API data"})
        self.assertEqual(step("Create the traffic App token")["with"]["permission-administration"],
                         "read")
        checkout = job["steps"][0]
        self.assertTrue(checkout["uses"].startswith("actions/checkout@"))
        self.assertIs(checkout["with"]["persist-credentials"], False)


if __name__ == "__main__":
    unittest.main(verbosity=2)
