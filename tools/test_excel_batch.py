"""Keep tools/Invoke-ExcelBatch.ps1 in step with the runner it drives (#11).

The batch script needs Windows and desktop Excel, so these tests check its
contract against the repository: the runner it calls, the manifest keys it
reads, the shipped example manifest and its safety rules.
"""
from __future__ import annotations

import json
import re
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BATCH = ROOT / "tools" / "Invoke-ExcelBatch.ps1"
RUNNER = ROOT / "tools" / "Invoke-ExcelRegression.ps1"
EXAMPLE = ROOT / "tools" / "excel-batch.example.json"
README = ROOT / "README.md"
SHA = re.compile(r"^[0-9a-f]{40}$")
MANIFEST_KEYS = {"schema_version", "runs", "label", "sha", "red",
                 "component", "expected_case", "from_ref", "find", "replace"}


def batch() -> str:
    return BATCH.read_text(encoding="ascii")


def components(text: str) -> list[str]:
    listed = re.search(r"\$Components = @\((.*?)\)", text, re.S)
    assert listed
    return re.findall(r"'([^']+)'", listed.group(1))


def git_blob(ref: str, path: str) -> str | None:
    shown = subprocess.run(["git", "-C", str(ROOT), "show", f"{ref}:{path}"],
                           capture_output=True, check=False)
    return shown.stdout.decode("latin-1") if shown.returncode == 0 else None


class BatchContractTests(unittest.TestCase):
    def test_script_is_ascii_with_crlf(self) -> None:
        data = BATCH.read_bytes()
        data.decode("ascii")
        self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))

    def test_components_match_the_runner(self) -> None:
        self.assertEqual(components(batch()), components(RUNNER.read_text(encoding="ascii")))

    def test_manifest_keys_read_are_the_documented_ones(self) -> None:
        read = set(re.findall(r"Get-Field \$\w+ '(\w+)'", batch()))
        self.assertEqual(read, MANIFEST_KEYS)

    def test_each_run_uses_its_own_worktree_and_runner(self) -> None:
        text = batch()
        self.assertIn("worktree add --detach", text)
        self.assertIn("Join-Path $worktree 'tools\\Invoke-ExcelRegression.ps1'", text)
        self.assertIn("'-ExpectedSha', $entry.sha, '-Repo', $worktree", text)
        self.assertRegex(text, r"(?s)\} finally \{\s*Remove-Worktree \$worktree")

    def test_safety_rules(self) -> None:
        text = batch()
        self.assertIn("Excel is already running", text)              # never touches the user's Excel
        self.assertIn("-OutDir must be outside the repository", text)
        for registry in ("Set-ItemProperty", "Remove-ItemProperty", "HKCU:"):
            self.assertNotIn(registry, text)                          # the runner owns AccessVBOM
        for verb in ("checkout", "reset", "push", "commit", "stash"):
            self.assertNotRegex(text, rf"git[^\r\n]*\s{verb}\b", verb)  # the user's checkout is left alone
        # The whole manifest is validated before Excel or a worktree is used.
        self.assertLess(text.index("$entries = @(Read-Manifest $Manifest)"),
                        text.index("worktree add --detach --quiet"))

    def test_missing_commits_are_fetched_by_sha(self) -> None:
        # A PR head can be reachable only from refs/pull/*, which a plain
        # `git fetch origin` does not map (review finding on #149).
        text = batch()
        self.assertIn("fetch --quiet origin @missing", text)
        self.assertNotRegex(text, r"fetch --quiet origin\r?$")

    def test_readme_documents_the_batch(self) -> None:
        readme = README.read_text(encoding="utf-8")
        self.assertIn("tools\\Invoke-ExcelBatch.ps1", readme)
        self.assertIn("tools/excel-batch.example.json", readme)


class ExampleManifestTests(unittest.TestCase):
    def setUp(self) -> None:
        self.data = json.loads(EXAMPLE.read_text(encoding="utf-8"))

    def test_shape(self) -> None:
        self.assertEqual(self.data["schema_version"], 1)
        allowed = set(components(batch()))
        self.assertTrue(self.data["runs"])
        for run in self.data["runs"]:
            with self.subTest(run=run.get("label")):
                self.assertLessEqual(set(run) | set(run.get("red", {})), MANIFEST_KEYS)
                self.assertRegex(run["sha"], SHA)
                red = run.get("red")
                if red is None:
                    continue
                self.assertIn(red["component"], allowed)
                self.assertIsInstance(red["expected_case"], int)
                self.assertGreaterEqual(red["expected_case"], 1)
                if "from_ref" in red:
                    self.assertRegex(red["from_ref"], SHA)
                    self.assertNotIn("find", red)
                else:
                    self.assertTrue(red["find"])
                    self.assertNotEqual(red["find"], red["replace"])

    def test_find_occurs_once_where_history_is_available(self) -> None:
        for run in self.data["runs"]:
            red = run.get("red") or {}
            if "find" not in red:
                continue
            text = git_blob(run["sha"], red["component"].replace("\\", "/"))
            if text is None:
                self.skipTest(f"{run['sha']} is not in this clone")
            self.assertEqual(text.count(red["find"]), 1, run["label"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
