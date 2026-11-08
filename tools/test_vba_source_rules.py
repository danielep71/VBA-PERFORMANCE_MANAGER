"""Positive and negative fixtures for the VBA source rules (#83)."""
from __future__ import annotations

import unittest

import gate_checks
import vba_source_rules as vr
from vba_source_rules import GOOD_MODULE

CLASS_HEADER = ("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
                'Attribute VB_Name = "cGood"\r\nAttribute VB_Exposed = False\r\n')


def run(source: str, path: str = "M_Good.bas") -> dict[str, list[str]]:
    return vr.repository_problems({path: source.encode("latin-1")})


class SourceRuleTests(unittest.TestCase):
    def check(self, source: str, path: str = "M_Good.bas") -> list[str]:
        return run(source, path)[vr.SOURCE_CHECK]

    def test_the_good_module_passes_every_check(self) -> None:
        self.assertEqual(run(GOOD_MODULE), {name: [] for name in run(GOOD_MODULE)})

    def test_each_bad_fixture_fails_only_its_own_check(self) -> None:
        for check, description, path, source in vr.BAD_FIXTURES:
            with self.subTest(description):
                found = run(source, path)
                self.assertTrue(found[check], description)
                others = [name for name, problems in found.items() if problems and name != check]
                self.assertEqual(others, [], description)

    def test_messages(self) -> None:
        cases = {
            GOOD_MODULE.replace('"M_Good"', '"M_Other"'):
                ["M_Good.bas: VB_Name 'M_Other' does not match the file name 'M_Good'"],
            GOOD_MODULE.replace("    Option Explicit\r\n", ""): ["M_Good.bas: no Option Explicit"],
            GOOD_MODULE.replace("text = Greeting", 'text = "open'):
                ["M_Good.bas:14: string literal is not closed"],
            GOOD_MODULE.replace("End Sub\r\n", "End Sub \r\n"): ["M_Good.bas:15: trailing white space"],
            GOOD_MODULE.replace("Option Private Module", "Option Base 1"):
                ["M_Good.bas:3: Option Base is not used; arrays state their bounds"],
        }
        for source, expected in cases.items():
            with self.subTest(expected=expected):
                self.assertEqual(self.check(source), expected)

    def test_vb_name_case_is_not_compared(self) -> None:
        self.assertEqual(self.check(GOOD_MODULE, "M_GOOD.bas"), [])

    def test_option_explicit_must_precede_the_first_procedure(self) -> None:
        late = GOOD_MODULE.replace("    Option Explicit\r\n", "").replace(
            "End Sub\r\n", "End Sub\r\nOption Explicit\r\n")
        self.assertEqual(self.check(late),
                         ["M_Good.bas:15: Option Explicit must precede the first procedure"])

    def test_quotes_in_comments_are_not_literals(self) -> None:
        body = GOOD_MODULE.replace("End Sub\r\n",
                                   "    Rem a \"stray quote\r\n"
                                   "    ' a comment that continues _\r\n"
                                   "      onto a line with \" one quote\r\n"
                                   "    text = \"Rem 'not a comment\"\r\nEnd Sub\r\n")
        self.assertEqual(self.check(body), [])

    def test_inline_rem_comments_allow_quotes_and_continuation(self) -> None:
        for statement in ('If True Then Rem a "quote',
                          'If False Then Beep Else rEm a "quote',
                          'Beep: Rem a "quote'):
            with self.subTest(statement=statement):
                body = GOOD_MODULE
                body = body.replace("End Sub\r\n",
                                    statement + ' _\r\n'
                                    '    another "quote\r\nEnd Sub\r\n')
                self.assertEqual(self.check(body), [])

    def test_rem_inside_literals_or_identifiers_does_not_hide_open_strings(self) -> None:
        for statement in ('text = "Then Rem still open',
                          'ThenRem = "open', 'ElseRem = "open',
                          'Remember = "open'):
            with self.subTest(statement=statement):
                source = GOOD_MODULE.replace("text = Greeting", statement)
                self.assertEqual(self.check(source),
                                 ["M_Good.bas:14: string literal is not closed"])

    def test_the_class_header_is_not_source(self) -> None:
        source = CLASS_HEADER + GOOD_MODULE.split("\r\n", 1)[1].replace("    Option Private Module\r\n", "")
        self.assertEqual(run(source, "cGood.cls"), {name: [] for name in run(GOOD_MODULE)})


class PtrSafeTests(unittest.TestCase):
    def problems(self, source: str) -> list[str]:
        return run(source)[vr.PTRSAFE_CHECK]

    def test_each_vba7_profile_reports_a_missing_ptrsafe(self) -> None:
        self.assertEqual(self.problems(GOOD_MODULE.replace("Declare PtrSafe Sub", "Declare Sub")), [
            "M_Good.bas:6: [VBA7-Win32] Declare Sleep needs PtrSafe",
            "M_Good.bas:6: [VBA7-Win64] Declare Sleep needs PtrSafe"])

    def test_ptrsafe_before_vba7_fails(self) -> None:
        source = GOOD_MODULE.replace("        Private Declare Sub", "        Private Declare PtrSafe Sub")
        self.assertEqual(self.problems(source), [
            "M_Good.bas:8: [pre-VBA7] Declare Sleep uses PtrSafe, which does not compile before VBA7"])

    def test_nested_win64_branches(self) -> None:
        nested = GOOD_MODULE.replace(
            "        Private Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\r\n",
            "        #If Win64 Then\r\n"
            "            Private Declare PtrSafe Function Tick Lib \"kernel32\" Alias \"GetTickCount64\" () As LongLong\r\n"
            "        #Else\r\n"
            "            Private Declare Function Tick Lib \"kernel32\" Alias \"GetTickCount\" () As Long\r\n"
            "        #End If\r\n")
        self.assertEqual(self.problems(nested), ["M_Good.bas:9: [VBA7-Win32] Declare Tick needs PtrSafe"])

    def test_a_continued_declare_is_one_statement(self) -> None:
        source = GOOD_MODULE.replace("Declare PtrSafe Sub Sleep Lib", "Declare PtrSafe Sub _\r\n            Sleep Lib")
        self.assertEqual(self.problems(source), [])

    def test_unsupported_conditional_fails_closed(self) -> None:
        source = GOOD_MODULE.replace("#If VBA7 Then", "#If Mac64 Then")
        self.assertTrue(all("unsupported source" in p for p in self.problems(source)))
        self.assertEqual(len(self.problems(source)), 3)


class VisibilityTests(unittest.TestCase):
    def problems(self, source: str, path: str = "M_Good.bas") -> list[str]:
        return run(source, path)[vr.VISIBILITY_CHECK]

    def test_messages(self) -> None:
        cases = {
            GOOD_MODULE.replace("Public Sub Run()", "Sub Run()"):
                ["M_Good.bas:13: procedure Run needs Public, Private or Friend"],
            GOOD_MODULE.replace("Public Sub Run()", "Static Sub Run()"):
                ["M_Good.bas:13: procedure Run needs Public, Private or Friend"],
            GOOD_MODULE.replace("    Private Const", "    Dim Counter As Long\r\n    Private Const"):
                ["M_Good.bas:4: module-level Dim; use Private or Public"],
            GOOD_MODULE.replace("    Private Const", "    Global Counter As Long\r\n    Private Const"):
                ["M_Good.bas:4: Global; use Public"],
            GOOD_MODULE.replace("    Private Const", "    Const"):
                ["M_Good.bas:4: module-level declaration needs Public or Private"],
            GOOD_MODULE.replace("Private Declare PtrSafe", "Declare PtrSafe"):
                ["M_Good.bas:6: module-level declaration needs Public or Private"],
            GOOD_MODULE.replace("    Public Enum", "    Enum"):
                ["M_Good.bas:10: Enum GoodKind needs Public or Private"],
        }
        for source, expected in cases.items():
            with self.subTest(expected=expected):
                self.assertEqual(self.problems(source), expected)

    def test_accepted_forms(self) -> None:
        source = CLASS_HEADER + (
            "    Option Explicit\r\n"
            "    Implements IGood\r\n"
            "    Private WithEvents mBook As Workbook\r\n"
            "    Public Event Changed(ByVal value As Long)\r\n"
            "    Private Type GoodState\r\n"
            "        Count As Long\r\n"
            "    End Type\r\n"
            "Friend Sub Reset()\r\n"
            "    Dim local As Long\r\n"
            "    Static calls As Long\r\n"
            "End Sub\r\n"
            "Private Property Get Count() As Long\r\n"
            "End Property\r\n")
        self.assertEqual(self.problems(source, "cGood.cls"), [])

    def test_an_event_needs_visibility(self) -> None:
        source = CLASS_HEADER + "    Option Explicit\r\n    Event Changed()\r\n"
        self.assertEqual(self.problems(source, "cGood.cls"),
                         ["cGood.cls:8: module-level declaration needs Public or Private"])


class PublicNameTests(unittest.TestCase):
    TWIN = GOOD_MODULE.replace('"M_Good"', '"M_Twin"').replace("GoodKind", "TwinKind") \
        .replace("GoodFirst", "TwinFirst")

    def problems(self, **sources: str) -> list[str]:
        return vr.repository_problems({path: text.encode("latin-1")
                                       for path, text in sources.items()})[vr.PUBLIC_NAME_CHECK]

    def test_a_procedure_in_two_modules(self) -> None:
        self.assertEqual(self.problems(**{"M_Good.bas": GOOD_MODULE, "M_Twin.bas": self.TWIN}), [
            "public name Run is defined in 2 standard modules: M_Good.bas:13, M_Twin.bas:13"])

    def test_enum_members_constants_and_declares_are_public_names(self) -> None:
        twin = self.TWIN.replace("Public Sub Run()", "Private Sub Walk()") \
            .replace("TwinFirst", "GoodFirst") \
            .replace("Private Const Greeting", "Public Const Greeting")
        good = GOOD_MODULE.replace("Private Const Greeting", "Public Const Greeting")
        self.assertEqual(self.problems(**{"M_Good.bas": good, "M_Twin.bas": twin}), [
            "public name GoodFirst is defined in 2 standard modules: M_Good.bas:11, M_Twin.bas:11",
            "public name Greeting is defined in 2 standard modules: M_Good.bas:4, M_Twin.bas:4"])

    def test_private_names_classes_and_other_profiles_are_not_shared(self) -> None:
        private = self.TWIN.replace("Public Sub Run()", "Private Sub Run()")
        self.assertEqual(self.problems(**{"M_Good.bas": GOOD_MODULE, "M_Twin.bas": private}), [])
        cls = CLASS_HEADER + "    Option Explicit\r\nPublic Sub Run()\r\nEnd Sub\r\n"
        self.assertEqual(self.problems(**{"M_Good.bas": GOOD_MODULE, "cGood.cls": cls}), [])
        split = ('Attribute VB_Name = "{0}"\r\n    Option Explicit\r\n    #If {1}VBA7 Then\r\n'
                 "        Public Const Shared As Long = 1\r\n    #End If\r\n")
        self.assertEqual(self.problems(**{"M_A.bas": split.format("M_A", ""),
                                          "M_B.bas": split.format("M_B", "Not ")}), [])


class FixtureAndRegistryTests(unittest.TestCase):
    def test_in_gate_fixtures_pass(self) -> None:
        self.assertEqual(vr.fixture_problems(), {name: [] for name in vr.fixture_problems()})

    def test_every_check_is_registered(self) -> None:
        for name in (vr.SOURCE_CHECK, vr.PTRSAFE_CHECK, vr.VISIBILITY_CHECK, vr.PUBLIC_NAME_CHECK):
            self.assertIn(name, gate_checks.NAMES)


if __name__ == "__main__":
    unittest.main(verbosity=2)
