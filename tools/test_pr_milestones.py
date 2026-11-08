"""Execute the milestone workflow shell against a stateful fake GitHub CLI."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

from test_governance import ROOT, load

FAKE_GH = r"""
import json
import os
import sys
from pathlib import Path

args = sys.argv[1:]
state_path = Path(os.environ["FAKE_STATE"])
state = json.loads(state_path.read_text())
with Path(os.environ["FAKE_LOG"]).open("a") as log:
    log.write(json.dumps(args) + "\n")
if os.environ.get("FAIL_API") == "list" and "--method" not in args:
    sys.exit(1)
if "--method" in args:
    assert args[args.index("--method") + 1] == "PATCH"
    if os.environ.get("FAIL_API") == "patch":
        sys.exit(1)
    path = next(a for a in args if a.startswith("repos/"))
    assert json.load(sys.stdin) == {"milestone": None}
    number = path.rsplit("/", 1)[1]
    state[number]["milestone"] = None
    state_path.write_text(json.dumps(state))
elif any("/pulls?" in a for a in args):
    path = next(a for a in args if "/pulls?" in a)
    items = list(state.values())
    if "state=open" in path:
        items = [item for item in items if item["state"] == "open"]
    else:
        assert "state=all" in path
    # Emulate the CLI's default first page versus its paginated result stream.
    if "--paginate" not in args:
        items = items[:100]
    for item in items:
        if item["milestone"] is not None:
            print(item["number"])
else:
    path = next(a for a in args if a.startswith("repos/"))
    item = state[path.rsplit("/", 1)[1]]
    if item["milestone"] is not None:
        print(item["number"])
"""


@unittest.skipUnless(shutil.which("bash"), "the workflow requires Bash")
class MilestoneWorkflowTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.state = self.root / "state.json"
        self.log = self.root / "calls.jsonl"
        gh = self.root / "gh"
        gh.write_text("#!" + sys.executable + "\n" + FAKE_GH)
        gh.chmod(0o755)
        step = load(ROOT / ".github/workflows/pr-milestones.yml")["jobs"]["clear"]["steps"][0]
        self.assertEqual(step["shell"], "bash", "explicit Bash enables pipefail in Actions")
        self.script = step["run"]

    def seed(self, items: list[dict[str, Any]]) -> None:
        self.state.write_text(json.dumps({str(item["number"]): item for item in items}))

    def run_workflow(self, number: str = "", fail: str = "") -> subprocess.CompletedProcess[str]:
        env = {**os.environ, "PATH": str(self.root) + os.pathsep + os.environ["PATH"],
               "REPO": "owner/repo", "PR_NUMBER": number, "FAKE_STATE": str(self.state),
               "FAKE_LOG": str(self.log), "FAIL_API": fail}
        return subprocess.run(["bash", "--noprofile", "--norc", "-eo", "pipefail", "-c",
                               self.script], env=env, capture_output=True, text=True, timeout=15)

    def calls(self) -> list[list[str]]:
        return [json.loads(line) for line in self.log.read_text().splitlines()]

    def patches(self) -> list[list[str]]:
        return [call for call in self.calls() if "--method" in call]

    def test_sweep_all_states_and_pages_then_is_idempotent(self) -> None:
        items = [{"number": n, "state": "open" if n == 1 else "closed",
                  "merged": n == 101, "milestone": {"number": 4}}
                 for n in range(1, 102)]
        self.seed(items)
        result = self.run_workflow()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(self.patches()), 101)
        self.assertTrue(all(item["milestone"] is None
                            for item in json.loads(self.state.read_text()).values()))
        result = self.run_workflow()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(self.patches()), 101, "second sweep must not PATCH again")

    def test_event_targets_open_closed_and_merged_prs_only(self) -> None:
        for state, merged in (("open", False), ("closed", False), ("closed", True)):
            with self.subTest(state=state, merged=merged):
                self.seed([{"number": 7, "state": state, "merged": merged, "milestone": 4},
                           {"number": 8, "state": "open", "milestone": 4}])
                self.log.write_text("")
                result = self.run_workflow("7")
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(len(self.patches()), 1)
                self.assertEqual(json.loads(self.state.read_text())["8"]["milestone"], 4)
                result = self.run_workflow("7")
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(len(self.patches()), 1)

    def test_event_without_milestone_is_a_noop(self) -> None:
        self.seed([{"number": 7, "state": "closed", "milestone": None}])
        result = self.run_workflow("7")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(self.patches(), [])

    def test_api_failures_fail_the_workflow(self) -> None:
        for fail in ("list", "patch"):
            with self.subTest(fail=fail):
                self.seed([{"number": 7, "state": "closed", "milestone": 4}])
                self.assertNotEqual(self.run_workflow(fail=fail).returncode, 0)


if __name__ == "__main__":
    unittest.main()
