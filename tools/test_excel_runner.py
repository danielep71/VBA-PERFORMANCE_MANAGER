"""Keep tools/Invoke-ExcelRegression.ps1 in step with the source it drives (#11).

The runner cannot execute here (it needs Windows and desktop Excel). These
tests check its contract against the repository instead, so a change to the
regression module, the README import order or the v1.4.0 result schema fails
the gate rather than the next maintainer run.
"""
from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RUNNER = ROOT / "tools" / "Invoke-ExcelRegression.ps1"
TEST_MODULE = ROOT / "test" / "M_cPM_Test.bas"
README = ROOT / "README.md"
FROZEN_SCHEMA = ROOT / "docs" / "releases" / "v1.4.0" / "IMPLEMENTATION_AND_RELEASE_PLAN.md"


def runner() -> str:
    return RUNNER.read_text(encoding="ascii")


def test_module() -> str:
    return TEST_MODULE.read_text(encoding="ascii")


def assignment(name: str) -> str:
    """The literal a top-level `$Name = ...` line of the runner assigns."""
    match = re.search(rf"^\${name}\s*=\s*(.+?)\s*(?:#.*)?$", runner(), re.M)
    assert match, f"${name} not found in the runner"
    return match.group(1).strip().strip("'")


class RunnerContractTests(unittest.TestCase):
    def test_components_follow_the_readme_import_order(self) -> None:
        listed = re.search(r"\$Components = @\((.*?)\)", runner(), re.S)
        assert listed
        components = [c.replace("\\", "/") for c in re.findall(r"'([^']+)'", listed.group(1))]
        block = re.search(r"Import the optional regression dependencies:\s*```text\n(.*?)```",
                          README.read_text(encoding="utf-8"), re.S)
        assert block
        self.assertEqual(components, block.group(1).split())
        for component in components:
            self.assertTrue((ROOT / component).is_file(), component)

    def test_entry_point_and_log_sheet_exist_in_the_regression_module(self) -> None:
        entry = assignment("EntryPoint")
        self.assertRegex(test_module(), rf"(?m)^Public Sub {entry}\(\)")
        sheet = re.search(r'cPM_SHEET_LOG\s+As String\s*=\s*"([^"]+)"', test_module())
        assert sheet
        self.assertEqual(assignment("LogSheet"), sheet.group(1))

    def test_summary_rows_and_columns_match_what_the_suite_writes(self) -> None:
        reset = re.findall(r"m_SummaryNextRow\s*=\s*(\d+)", test_module())
        self.assertIn(assignment("FirstSummaryRow"), reset)
        written = dict(re.findall(r'\.Cells\(SummaryRow, "([A-Z])"\)\.Value\s*=\s*(\w+)', test_module()))
        self.assertEqual(written, {"C": "m_RunTimestamp", "D": "m_TotalCases",
                                   "E": "m_CurrentCaseName", "F": "CaseAssertions",
                                   "G": "CaseFailures", "H": "ResultText"})
        # The runner reads case number, name, assertions, failures, result: D..H.
        read = re.findall(r"\$sheet\.Cells\.Item\(\$row, (\d)\)", runner())
        self.assertEqual(sorted(set(read)), ["4", "5", "6", "7", "8"])
        self.assertIn("'PASS'", runner())
        self.assertIn('ResultText = "PASS"', test_module())

    def test_total_steps_pattern_matches_the_module(self) -> None:
        pattern = re.search(r"-notmatch '([^']+)'", runner())
        assert pattern
        self.assertTrue(re.search(pattern.group(1), test_module()))

    def test_result_uses_the_frozen_v1_4_0_schema(self) -> None:
        block = re.search(r"### 0\.1 Freeze the execution-result schema.*?```json\n(.*?)```",
                          FROZEN_SCHEMA.read_text(encoding="utf-8"), re.S)
        assert block
        frozen = list(json.loads(block.group(1)))
        result = re.search(r"^    result = \[ordered\]@\{(.*?)\n    \}", runner(), re.S | re.M)
        assert result
        self.assertEqual(re.findall(r"^\s+(\w+) =", result.group(1), re.M), frozen)

    def test_safety_rules(self) -> None:
        text = runner()
        self.assertIn("Excel is already running", text)            # never touches the user's Excel
        self.assertIn("$workbook.Close($false)", text)              # never saves
        self.assertIn("Remove-ItemProperty -Path $VbomKey", text)   # AccessVBOM restored ...
        self.assertRegex(text, r"(?s)\} finally \{\s*if \(\$AllowVbomAccess\)")  # ... on every path
        self.assertIn("-OutDir must be outside the repository", text)

    def test_runner_is_ascii_with_crlf(self) -> None:
        data = RUNNER.read_bytes()
        data.decode("ascii")
        self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
