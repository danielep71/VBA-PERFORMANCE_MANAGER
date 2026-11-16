"""Keep tools/Verify-SyncRefusals.ps1 in step with the sync it exercises (#138).

The script needs Windows and desktop Excel, so these tests check its contract
against the repository: the tool it runs, the refusal messages it expects,
its launcher and its safety rules.
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "tools" / "Verify-SyncRefusals.ps1"
LAUNCHER = ROOT / "tools" / "Verify-SyncRefusals.cmd"
SYNC = ROOT / "tools" / "Manage-PerformanceManagerVBA.ps1"
DOC = ROOT / "docs" / "LOCAL_EXCEL_TOOLS.md"


def text(path: Path) -> str:
    return path.read_text(encoding="ascii")


class SyncRefusalContractTests(unittest.TestCase):
    def test_files_are_ascii_with_crlf(self) -> None:
        for path in (SCRIPT, LAUNCHER):
            with self.subTest(path=path.name):
                data = path.read_bytes()
                data.decode("ascii")
                self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))

    def test_expected_refusals_are_the_sync_messages(self) -> None:
        expected = re.findall(r"Expected = '([^']+)'", text(SCRIPT))
        self.assertEqual(len(expected), 2)
        sync = text(SYNC)
        self.assertIn("throw 'VBA project is protected.", sync)
        self.assertIn('throw "Component type conflict: $name"', sync)
        self.assertTrue(expected[0].startswith("VBA project is protected"))
        self.assertEqual(expected[1], "Component type conflict: M_cPM_TIMEWASTERS")
        # The conflicting class must carry the name the sync manages first.
        self.assertRegex(sync, r"\$components = @\('src\\modules\\M_cPM_TIMEWASTERS\.bas'")
        self.assertIn("New-TestWorkbook $conflict 2 'M_cPM_TIMEWASTERS'", text(SCRIPT))

    def test_it_runs_the_sync_tool_in_sync_mode(self) -> None:
        self.assertIn("Join-Path $PSScriptRoot 'Manage-PerformanceManagerVBA.ps1'", text(SCRIPT))
        self.assertIn("-File $tool -Mode Sync", text(SCRIPT))
        self.assertIn("-AllowVbomAccess -Workbook $case.Workbook", text(SCRIPT))
        self.assertIn('"%~dp0Verify-SyncRefusals.ps1"', text(LAUNCHER))

    def test_every_documented_check_is_evaluated(self) -> None:
        checks = re.findall(r"^\s{8}(\w+)\s+= \(", text(SCRIPT), re.M)
        self.assertEqual(checks, ["exit_nonzero", "result_fail", "expected_error", "not_passed",
                                  "cleanup_completed", "input_unchanged", "vbom_restored",
                                  "no_excel_left"])

    def test_safety_rules(self) -> None:
        script = text(SCRIPT)
        # AccessVBOM is restored on every path of every private Excel instance.
        self.assertRegex(script, r"(?s)\} finally \{.*?Restore-Vbom \$vbom\s*\}")
        # Only invisible Excel instances are closed; a visible window stops the run.
        self.assertIn("if ($process.MainWindowTitle) {", script)
        self.assertIn("-WorkDir must be outside the repository.", script)
        self.assertIn("The checkout is not clean", script)
        for verb in ("checkout", "reset", "push", "commit", "stash", "clean"):
            self.assertNotRegex(script, rf"git[^\r\n]*\s{verb}\b", verb)

    def test_documented(self) -> None:
        self.assertIn("`Verify-SyncRefusals.cmd`", DOC.read_text(encoding="utf-8"))
        self.assertIn("`Verify-SyncRefusals.cmd`", (ROOT / "README.md").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
