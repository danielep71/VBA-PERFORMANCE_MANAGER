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


PROGRESS = re.compile(r'^\s*Demo_SB_SetProgress\s+CurrentStep,\s*TotalSteps,\s*"([^"]*)"\s*$', re.M)
CASE_CALL = re.compile(r"^\s*(Test_\w+)\s*$", re.M)
CASE_BEGIN = re.compile(r'^\s*Case_Begin\s+"([^"]*)"', re.M)
CASE_PROC = re.compile(r"^(?:Private |Public )?Sub (Test_\w+)\(.*?^End Sub", re.M | re.S)


def label_problems(text: str) -> list[str]:
    """Runner status-bar labels that differ from their case's Case_Begin name (#78).

    Each `Demo_SB_SetProgress CurrentStep, TotalSteps, "<label>"` followed by a
    `Test_<name>` call must show the one name that procedure passes to
    Case_Begin, so a status-bar label can be found in the log by its text. A
    label not followed by a case call (the preparation step) is not checked.
    """
    text = text.replace("\r\n", "\n")
    names = {m.group(1): CASE_BEGIN.findall(m.group(0)) for m in CASE_PROC.finditer(text)}
    problems: list[str] = []
    for label in PROGRESS.finditer(text):
        following = CASE_CALL.search(text, label.end())
        next_label = PROGRESS.search(text, label.end())
        if following is None or (next_label is not None and next_label.start() < following.start()):
            continue
        case = following.group(1)
        begun = names.get(case)
        if not begun or len(begun) != 1:
            problems.append(f"{case}: needs exactly one Case_Begin name, has {begun or 'none'}")
        elif begun[0] != label.group(1):
            problems.append(f"{case}: runner label {label.group(1)!r} differs from "
                            f"Case_Begin {begun[0]!r}")
    return problems


class CaseLabelTests(unittest.TestCase):
    MODULE = (
        'Public Sub Run()\r\n'
        '        Demo_SB_SetProgress CurrentStep, TotalSteps, "Preparing regression environment"\r\n'
        '        Suite_Prepare\r\n'
        '        Demo_SB_SetProgress CurrentStep, TotalSteps, "First case"\r\n'
        '        Test_First\r\n'
        'End Sub\r\n'
        'Private Sub Test_First()\r\n'
        '        Case_Begin "First case"\r\n'
        'End Sub\r\n'
    )

    def test_matching_labels_pass(self) -> None:
        self.assertEqual(label_problems(self.MODULE), [])

    def test_a_drifted_label_is_reported(self) -> None:
        drifted = self.MODULE.replace('TotalSteps, "First case"', 'TotalSteps, "First"')
        self.assertEqual(label_problems(drifted),
                         ["Test_First: runner label 'First' differs from Case_Begin 'First case'"])

    def test_a_case_needs_exactly_one_name(self) -> None:
        missing = self.MODULE.replace('        Case_Begin "First case"\r\n', "")
        self.assertEqual(label_problems(missing),
                         ["Test_First: needs exactly one Case_Begin name, has none"])
        twice = self.MODULE.replace('        Case_Begin "First case"\r\n',
                                    '        Case_Begin "First case"\r\n        Case_Begin "Again"\r\n')
        self.assertIn("needs exactly one Case_Begin name", label_problems(twice)[0])

    def test_regression_module_labels_match_case_names(self) -> None:
        self.assertEqual(label_problems(test_module()), [])
        total = re.search(r"Const TotalSteps\s+As Long\s*=\s*(\d+)", test_module())
        assert total
        # Every case is labelled: one label per case, plus the preparation step.
        self.assertEqual(len(PROGRESS.findall(test_module())) - 1, int(total.group(1)))


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

    def test_cleanup_releases_every_com_reference_and_waits(self) -> None:
        """A held reference keeps Excel alive after Quit, which forced cleanup once."""
        text = runner()
        cleanup = text[text.index("} finally {", text.index("function Invoke-Suite")):]
        for name in ("$sheet", "$control", "$project", "$workbook", "$excel"):
            self.assertIn(f"Release-Com {name}", cleanup)
        self.assertLess(cleanup.index("Release-Com $project"), cleanup.index("$excel.Quit()"))
        self.assertRegex(text, r"\$ExitWaitSeconds = \d+")
        self.assertIn("while ((Get-Process -Id $excelPid", cleanup)
        self.assertIn('Write-Host "  cleanup: $($run.cleanup)"', text)

    def test_runner_is_ascii_with_crlf(self) -> None:
        data = RUNNER.read_bytes()
        data.decode("ascii")
        self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
