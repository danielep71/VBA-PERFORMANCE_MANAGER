"""Fixtures for the derived VBA component inventory (#82)."""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from release_provenance import OPTIONAL, REQUIRED
from vba_inventory import ROLES, inventory_problems, module_name, tracked_vba, under

ROOT = Path(__file__).resolve().parent.parent
CURRENT = sorted(REQUIRED + OPTIONAL)


def git(cwd: Path, *args: str) -> None:
    subprocess.run(["git", *args], cwd=cwd, check=True, capture_output=True)


class InventoryRuleTests(unittest.TestCase):
    def test_repository_inventory_is_complete(self) -> None:
        tracked = tracked_vba(ROOT)
        self.assertEqual(sorted(under(tracked, "src", "test", "demo")), CURRENT)
        self.assertEqual(inventory_problems(tracked, REQUIRED + OPTIONAL), [])

    def test_current_tree_fixture_passes(self) -> None:
        self.assertEqual(inventory_problems(CURRENT, CURRENT), [])

    def test_new_component_must_be_hashed_by_provenance(self) -> None:
        problems = inventory_problems(CURRENT + ["demo/M_cPM_New.bas"], CURRENT)
        self.assertEqual(problems, ["demo/M_cPM_New.bas: tracked component missing from release "
                                    "provenance (REQUIRED or OPTIONAL in tools/release_provenance.py)"])

    def test_component_outside_the_checked_roots_fails(self) -> None:
        for path in ("M_Stray.bas", "tools/M_Stray.cls", "docs/Form.frm"):
            with self.subTest(path=path):
                problems = inventory_problems(CURRENT + [path], CURRENT)
                self.assertEqual(len(problems), 1)
                self.assertIn(f"{path}: VBA component outside src/, test/, demo/", problems[0])

    def test_missing_role_file_fails(self) -> None:
        tracked = [path for path in CURRENT if path != ROLES["test"]]
        problems = inventory_problems(tracked, tracked)
        self.assertEqual(problems, [f"test role file {ROLES['test']} is not a tracked component"])

    def test_role_needs_its_extension_and_a_distinct_file(self) -> None:
        roles = dict(ROLES, module=ROLES["class"])
        problems = inventory_problems(CURRENT, CURRENT, roles)
        self.assertIn(f"module role file {ROLES['class']} must be a .bas export", problems)
        self.assertIn("two roles name the same file", problems)

    def test_provenance_cannot_list_an_untracked_component(self) -> None:
        problems = inventory_problems(CURRENT, CURRENT + ["demo/M_Gone.bas"])
        self.assertEqual(problems, ["demo/M_Gone.bas: listed by release provenance but not a "
                                    "tracked component"])

    def test_module_name_reads_vb_name(self) -> None:
        self.assertEqual(module_name('Attribute VB_Name = "M_cPM_TimeWasters"\r\n'), "M_cPM_TimeWasters")
        self.assertIsNone(module_name("Option Explicit\n"))


class TrackedListingTests(unittest.TestCase):
    def test_only_tracked_vba_files_are_listed(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            repo = Path(td)
            git(repo, "init", "-q")
            for rel in ("src/classes/A.cls", "demo/B.bas", "test/C.BAS", "docs/notes.md",
                        "demo/Untracked.bas"):
                (repo / rel).parent.mkdir(parents=True, exist_ok=True)
                (repo / rel).write_text("x\n", encoding="utf-8")
            git(repo, "add", "src", "demo/B.bas", "test", "docs")
            self.assertEqual(tracked_vba(repo), ["demo/B.bas", "src/classes/A.cls", "test/C.BAS"])

    def test_git_failure_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            with self.assertRaisesRegex(RuntimeError, "git ls-files failed"):
                tracked_vba(Path(td) / "not-a-repository")


class StaticGateTests(unittest.TestCase):
    def test_new_tracked_component_is_checked(self) -> None:
        """A new demo module is checked without being listed anywhere."""
        with tempfile.TemporaryDirectory(prefix="cpm-inventory-") as td:
            dest = Path(td) / "repo"
            shutil.copytree(ROOT, dest, ignore=shutil.ignore_patterns(".git", "__pycache__"))
            git(dest, "init", "-q")
            new = dest / "demo" / "M_cPM_NewDemo.bas"
            new.write_text('Attribute VB_Name = "M_cPM_NewDemo"\nOption Explicit\n'
                           "Public Sub Demo()\n    Dim Empty As Long\nEnd Sub\n", encoding="utf-8")
            git(dest, "add", "-A")
            result = dest / "result.json"
            proc = subprocess.run([sys.executable, str(dest / "tools" / "vba_lint.py"),
                                   "--json", str(result)], cwd=dest, capture_output=True,
                                  text=True, timeout=300)
            self.assertNotEqual(proc.returncode, 0)
            report = json.loads(result.read_text(encoding="utf-8"))
            failed = {r["check"]: r["problems"] for r in report["results"] if not r["passed"]}
            self.assertEqual(sorted(failed), ["no reserved words as identifiers",
                                              "tracked VBA components are inventoried"])
            self.assertTrue(any("M_cPM_NewDemo.bas" in p
                                for p in failed["no reserved words as identifiers"]))
            self.assertTrue(any(p.startswith("demo/M_cPM_NewDemo.bas: tracked component missing")
                                for p in failed["tracked VBA components are inventoried"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
