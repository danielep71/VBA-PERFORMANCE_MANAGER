"""Keep tools/Build-DemoWorkbook.ps1 in step with the demo it builds (#47).

The build needs Windows and desktop Excel, so these tests check its contract
against the repository: the components it imports, the entry point and the
sheets it expects, the demo actions it smoke-tests, the metadata kinds it
removes, and its safety rules.
"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "tools" / "Build-DemoWorkbook.ps1"
LAUNCHER = ROOT / "tools" / "Build-DemoWorkbook.cmd"
DEMO = ROOT / "demo" / "M_cPM_DEMO.bas"
SYNC = ROOT / "tools" / "Manage-PerformanceManagerVBA.ps1"
PROVENANCE = ROOT / "tools" / "release_provenance.py"


def script() -> str:
    return SCRIPT.read_text(encoding="ascii")


def demo() -> str:
    return DEMO.read_text(encoding="latin-1")


def ps_list(text: str, name: str) -> list[str]:
    listed = re.search(rf"\${name} = @\((.*?)\)", text, re.S)
    assert listed, name
    return re.findall(r"'([^']+)'", listed.group(1))


class DemoBuildContractTests(unittest.TestCase):
    def test_files_are_ascii_with_crlf(self) -> None:
        for path in (SCRIPT, LAUNCHER):
            with self.subTest(path=path.name):
                data = path.read_bytes()
                data.decode("ascii")
                self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))

    def test_components_are_the_six_managed_ones(self) -> None:
        self.assertEqual(sorted(ps_list(script(), "Components")), sorted(ps_list(SYNC.read_text(encoding="ascii"), "components")))
        for component in ps_list(script(), "Components"):
            self.assertTrue((ROOT / component.replace("\\", "/")).is_file(), component)

    def test_entry_point_sheets_and_log_exist_in_the_demo(self) -> None:
        source = demo()
        self.assertRegex(source, r"(?m)^Public Sub cPM_CreateDemoSheets\(\)")
        sheets = dict(re.findall(r'Private Const (cPM_SHEET_\w+)\s+As String = "([^"]+)"', source))
        self.assertEqual(sorted(ps_list(script(), "Sheets")), sorted(sheets.values()))
        self.assertIn('Private Const cPM_TABLE_LOG             As String = "Tbl_cPM_Log"', source)
        self.assertIn("$LogTable = 'Tbl_cPM_Log'", script())

    def test_every_demo_action_is_smoke_tested(self) -> None:
        grid = re.findall(r'Array\("btn_cPM_\w+", "[^"]+", "(\w+)"\)', demo())
        self.assertEqual(sorted(ps_list(script(), "SmokeActions") + ["cPM_ClearResultsLog"]), sorted(grid))
        for helper in ("cPM_Demo_SetUnattended", "cPM_Demo_FailureCount"):
            self.assertRegex(demo(), rf"(?m)^Public (Sub|Function) {helper}\(")
            self.assertIn(f"'{helper}'", script())

    def test_actions_report_failures_instead_of_raising(self) -> None:
        for action in ps_list(script(), "SmokeActions"):
            body = re.search(rf"(?ms)^Public Sub {action}\(\)(.*?)^End Sub", demo())
            assert body, action
            self.assertIn("cPM_Demo_ReportFailure", body.group(1))
            self.assertNotIn("Err.Raise SavedErrNumber", body.group(1))

    def test_metadata_kinds_match_the_provenance_check(self) -> None:
        text = script()
        for marker in ("docMetadata/", "classificationlabels", "MSIP_Label_", "category|contentStatus",
                       "absPath", "(odd|even|first)(Header|Footer)"):
            self.assertIn(marker, text, marker)
        provenance = PROVENANCE.read_text(encoding="utf-8")
        for marker in ("docmetadata/", "classificationlabels", "MSIP_Label_", "category|contentStatus",
                       "absPath", "(odd|even|first)(Header|Footer)"):
            self.assertIn(marker, provenance, marker)

    def test_safety_rules(self) -> None:
        text = script()
        self.assertRegex(text, r"(?s)\} finally \{.*?if \(\$vbomTouched\)")
        self.assertIn("-OutDir must be outside the repository.", text)
        self.assertIn("The checkout has local changes", text)
        self.assertIn("if ($process.MainWindowTitle) {", text)
        for verb in ("checkout", "reset", "push", "commit", "stash", "clean"):
            self.assertNotRegex(text, rf"git[^\r\n]*\s{verb}\b", verb)
        self.assertIn('"%~dp0Build-DemoWorkbook.ps1" -AllowVbomAccess', LAUNCHER.read_text(encoding="ascii"))

    def test_documented(self) -> None:
        self.assertIn("`Build-DemoWorkbook.cmd`", (ROOT / "docs" / "LOCAL_EXCEL_TOOLS.md").read_text(encoding="utf-8"))
        self.assertIn("`Build-DemoWorkbook.cmd`", (ROOT / "README.md").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
