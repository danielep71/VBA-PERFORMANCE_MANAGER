"""Fixtures for the tracked-binary review (#84)."""
from __future__ import annotations

import hashlib
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from typing import Any

from binary_review import ALLOWLIST, parse_allowlist, repository_problems

ROOT = Path(__file__).resolve().parent.parent
PNG = b"\x89PNG\r\n\x1a\n" + b"\0" * 16
ATTRIBUTES = "*.png binary\n*.xlsm binary\n*.frx binary\n"


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def allowlist(files: dict[str, Any]) -> bytes:
    return json.dumps({"schema_version": 1, "files": files}).encode()


def entry(data: bytes, reason: str = "fixture") -> dict[str, str]:
    return {"sha256": digest(data), "reason": reason}


class BinaryReviewTests(unittest.TestCase):
    def problems(self, files: dict[str, bytes], *, untracked: dict[str, bytes] | None = None,
                 attributes: str = ATTRIBUTES) -> list[str]:
        """Review a scratch repository holding files (tracked) and untracked files."""
        with tempfile.TemporaryDirectory() as td:
            repo = Path(td)
            subprocess.run(["git", "init", "-q"], cwd=repo, check=True)
            files = {".gitattributes": attributes.encode(), **files}
            files.setdefault(ALLOWLIST, allowlist({}))
            for name, data in {**files, **(untracked or {})}.items():
                (repo / name).parent.mkdir(parents=True, exist_ok=True)
                (repo / name).write_bytes(data)
            subprocess.run(["git", "add", "--", *files], cwd=repo, check=True)
            return repository_problems(repo)

    def test_repository_binaries_are_reviewed(self) -> None:
        self.assertEqual(repository_problems(ROOT), [])

    def test_allowlisted_binary_with_its_digest_passes(self) -> None:
        self.assertEqual(self.problems({"images/a.png": PNG,
                                        ALLOWLIST: allowlist({"images/a.png": entry(PNG)})}), [])

    def test_unapproved_binary_fails(self) -> None:
        problems = self.problems({"images/a.png": PNG})
        self.assertEqual(len(problems), 1)
        self.assertIn("images/a.png: tracked binary is neither on an approved path", problems[0])

    def test_changed_digest_fails(self) -> None:
        old = {ALLOWLIST: allowlist({"images/a.png": entry(PNG)})}
        problems = self.problems({"images/a.png": PNG + b"changed", **old})
        self.assertEqual(len(problems), 1)
        self.assertIn("differs from the allowlisted", problems[0])

    def test_nul_byte_file_without_attributes_fails(self) -> None:
        problems = self.problems({"data/blob.dat": b"head\0tail"})
        self.assertIn("data/blob.dat: tracked binary", "\n".join(problems))
        self.assertEqual(self.problems({"data/notes.dat": b"plain text\n"}), [])

    def test_committed_workbook_fails(self) -> None:
        problems = self.problems({"PERFORMANCE MANAGER.xlsm": b"PK no nul byte"})
        self.assertIn("PERFORMANCE MANAGER.xlsm: tracked binary", "\n".join(problems))

    def test_userform_binary_is_approved_only_beside_its_form(self) -> None:
        self.assertEqual(self.problems({"demo/Form.frm": b"VERSION 5.00\n", "demo/Form.frx": PNG}), [])
        alone = self.problems({"demo/Form.frx": PNG})
        self.assertIn("demo/Form.frx: tracked binary", "\n".join(alone))
        outside = self.problems({"docs/Form.frm": b"VERSION 5.00\n", "docs/Form.frx": PNG})
        self.assertIn("docs/Form.frx: tracked binary", "\n".join(outside))

    def test_stale_entries_fail(self) -> None:
        gone = self.problems({ALLOWLIST: allowlist({"images/gone.png": entry(PNG)})})
        self.assertEqual(gone, [f"{ALLOWLIST}: images/gone.png is not a tracked file; remove the entry"])
        text = b"now plain text\n"
        not_binary = self.problems({"images/a.txt": text,
                                    ALLOWLIST: allowlist({"images/a.txt": entry(text)})})
        self.assertEqual(not_binary, [f"{ALLOWLIST}: images/a.txt is not binary; remove the entry"])

    def test_untracked_binary_is_not_reviewed(self) -> None:
        self.assertEqual(self.problems({}, untracked={"PERFORMANCE MANAGER.xlsm": PNG}), [])

    def test_malformed_allowlist_fails(self) -> None:
        cases = {
            "{": "invalid JSON",
            '{"files": {}}': "expected",
            '{"schema_version": 1, "files": {"a.png": {"sha256": "ABC", "reason": "r"}}}':
                "sha256 must be 64 lowercase hex digits",
            '{"schema_version": 1, "files": {"a.png": {"sha256": "' + "0" * 64 + '"}}}':
                "a reason is required",
        }
        for text, fragment in cases.items():
            with self.subTest(text=text):
                entries, problems = parse_allowlist(text)
                self.assertEqual(entries, {})
                self.assertIn(fragment, "\n".join(problems))

    def test_git_failure_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            problems = repository_problems(Path(td) / "missing")
        self.assertEqual(len(problems), 1)
        self.assertIn("binaries were not reviewed", problems[0])


if __name__ == "__main__":
    unittest.main(verbosity=2)
