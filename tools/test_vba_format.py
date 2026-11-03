"""Fixtures for the VBA layout formatter (#83)."""
from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import gate_checks
import vba_format as vf

HEAD = 'Attribute VB_Name = "M_Fixture"\r\n    Option Explicit\r\n'


def formatted(body: str) -> str:
    return vf.format_text(HEAD + body)[len(HEAD):]


class RuleTests(unittest.TestCase):
    def test_each_rule(self) -> None:
        cases = {
            # 1. CRLF, one final line end, no blank lines at the end
            "Public Sub A()\nEnd Sub": "Public Sub A()\r\nEnd Sub\r\n",
            "Public Sub A()\r\nEnd Sub\r\n\r\n\r\n": "Public Sub A()\r\nEnd Sub\r\n",
            # 2. trailing white space, including on blank lines
            "    Private x As Long  \r\n    \r\n": "    Private x As Long\r\n",
            # 3. tabs in indentation
            "\tPrivate x As Long\r\n\t\tPrivate y As Long\r\n":
                "    Private x As Long\r\n        Private y As Long\r\n",
            # 4. declarations, End statements and labels in column 0
            "  Public Sub A()\r\n  Done:\r\n  End Sub\r\n": "Public Sub A()\r\nDone:\r\nEnd Sub\r\n",
            # 5. one or two blank lines after End
            "Public Sub A()\r\nEnd Sub\r\nPublic Sub B()\r\nEnd Sub\r\n":
                "Public Sub A()\r\nEnd Sub\r\n\r\nPublic Sub B()\r\nEnd Sub\r\n",
            "Public Sub A()\r\nEnd Sub\r\n\r\n\r\n\r\n\r\nPublic Sub B()\r\nEnd Sub\r\n":
                "Public Sub A()\r\nEnd Sub\r\n\r\n\r\nPublic Sub B()\r\nEnd Sub\r\n",
        }
        for body, expected in cases.items():
            with self.subTest(body=body):
                self.assertEqual(formatted(body), expected)

    def test_two_blank_lines_after_end_are_kept(self) -> None:
        body = "Public Sub A()\r\nEnd Sub\r\n\r\n\r\nPublic Sub B()\r\nEnd Sub\r\n"
        self.assertEqual(formatted(body), body)

    def test_what_it_never_changes(self) -> None:
        body = ("Public Sub A( _\r\n"
                "    ByVal x As Long)\r\n"           # continuation keeps its indent
                "    Debug.Print \"  spaced\t  \"\r\n"  # literal kept, tab inside it too
                "    ' comment:  kept  \"as is\r\n"
                "End Sub\r\n")
        self.assertEqual(formatted(body), body)
        header = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
        self.assertTrue(vf.format_text(header + HEAD).startswith(header))

    def test_in_gate_fixtures_pass_and_the_check_is_registered(self) -> None:
        self.assertEqual(vf.fixture_problems(), [])
        self.assertIn(vf.CHECK, gate_checks.NAMES)

    def test_the_repository_is_formatted_and_idempotent(self) -> None:
        self.assertEqual(vf.problems(), [])
        for rel in vf.repository_files():
            text = vf.read(vf.ROOT / rel)
            self.assertEqual(vf.format_text(text), text, rel)


class CommandTests(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = tempfile.TemporaryDirectory()
        self.root = Path(self.dir.name)
        self.messy = HEAD + "  Public Sub A()  \r\nEnd Sub\r\n\r\n\r\n"
        (self.root / "M_Fixture.bas").write_bytes(self.messy.encode("latin-1"))

    def tearDown(self) -> None:
        self.dir.cleanup()

    def main(self, *args: str) -> tuple[int, str]:
        out = io.StringIO()
        with mock.patch.object(vf, "ROOT", self.root), contextlib.redirect_stdout(out):
            status = vf.main([*args, "M_Fixture.bas"])
        return status, out.getvalue()

    def test_check_fails_and_writes_nothing(self) -> None:
        status, out = self.main("--check")
        self.assertEqual(status, 1)
        self.assertIn("M_Fixture.bas: layout differs", out)
        self.assertEqual((self.root / "M_Fixture.bas").read_bytes(), self.messy.encode("latin-1"))

    def test_write_then_check_passes(self) -> None:
        self.assertEqual(self.main("--write")[0], 0)
        self.assertEqual((self.root / "M_Fixture.bas").read_bytes(),
                         (HEAD + "Public Sub A()\r\nEnd Sub\r\n").encode("latin-1"))
        self.assertEqual(self.main("--check"), (0, "1 file(s) formatted\n"))

    def test_a_mode_is_required(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            vf.main([])


if __name__ == "__main__":
    unittest.main(verbosity=2)
