"""Fixtures for the README facts renderer (#89)."""
from __future__ import annotations

import contextlib
import io
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from typing import Any

import binary_review
import gate_checks
import markdown_links
import readme_facts as rf
import vba_compile_safety
import vba_lint
import workflow_pins

SOURCES = ("README.md", "VERSION", rf.TEST_MODULE, rf.EVIDENCE)


def recorded(tag: str) -> str | None:
    """Resolve each tag to the commit its record names, as an untampered tag would."""
    data = json.loads((rf.ROOT / rf.EVIDENCE).read_text(encoding="utf-8"))
    return next((r["commit"] for r in data["releases"] if r["tag"] == tag), None)


class Copy:
    """A temporary repository holding only the files the renderer reads."""

    def __init__(self) -> None:
        self.dir = tempfile.TemporaryDirectory()
        self.root = Path(self.dir.name)
        for rel in SOURCES:
            (self.root / rel).parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(rf.ROOT / rel, self.root / rel)

    def text(self, rel: str) -> str:
        with open(self.root / rel, encoding="utf-8", newline="") as handle:
            return handle.read()

    def edit(self, rel: str, old: str, new: str) -> None:
        text = self.text(rel)
        assert old in text, old
        (self.root / rel).write_text(text.replace(old, new, 1), encoding="utf-8", newline="")

    def evidence(self) -> dict[str, Any]:
        data: dict[str, Any] = json.loads(self.text(rf.EVIDENCE))
        return data

    def set_evidence(self, data: dict[str, Any]) -> None:
        (self.root / rf.EVIDENCE).write_text(json.dumps(data), encoding="utf-8")

    def problems(self) -> list[str]:
        return rf.repository_problems(self.root, recorded)[0]

    def close(self) -> None:
        self.dir.cleanup()


class RepositoryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.copy = Copy()

    def tearDown(self) -> None:
        self.copy.close()

    def test_the_current_readme_is_rendered(self) -> None:
        self.assertEqual(rf.repository_problems(rf.ROOT, recorded), ([], []))

    def test_a_stale_block_fails_and_write_repairs_it(self) -> None:
        count = len(gate_checks.CHECKS)
        self.copy.edit("README.md", f"**{count} checks**", f"**{count - 1} checks**")
        self.assertEqual(self.copy.problems(), [
            "README.md: facts:static-checks is stale; run python3 tools/readme_facts.py --write"])
        evidence = self.copy.text(rf.EVIDENCE)
        self.assertEqual(rf.write(self.copy.root), (["static-checks"], []))
        self.assertEqual(self.copy.problems(), [])
        self.assertEqual(self.copy.text(rf.EVIDENCE), evidence)   # certified data is never written
        self.assertEqual(self.copy.text("README.md"), (rf.ROOT / "README.md").read_text("utf-8"))

    def test_derived_facts_follow_the_source(self) -> None:
        cases = rf.load_facts(rf.ROOT)[0].cases   # type: ignore[union-attr]
        self.copy.edit("VERSION", "1.4.0", "1.4.1")
        self.copy.edit(rf.TEST_MODULE, "        Test_DefaultState\r\n", "")
        stale = [p.split("facts:")[1].split()[0] for p in self.copy.problems()]
        self.assertEqual(stale, ["badges", "source-cases"])
        rf.write(self.copy.root)
        readme = self.copy.text("README.md")
        self.assertIn("Version-1.4.1-", readme)
        self.assertIn(f"invokes **{cases - 1} cases**", readme)

    def test_certified_facts_ignore_the_source(self) -> None:
        before = self.copy.text("README.md")
        self.copy.edit("VERSION", "1.4.0", "2.0.0")
        rf.write(self.copy.root)
        after = self.copy.text("README.md")
        for name in ("evidence:v1.4.0", "evidence:v1.3.0", "summary:v1.4.0"):
            self.assertEqual(self.block(before, name), self.block(after, name))

    def test_certified_blocks_render_only_their_record(self) -> None:
        data = self.copy.evidence()
        data["releases"][0]["assertions"] = 650
        self.copy.set_evidence(data)
        stale = [p.split("facts:")[1].split()[0] for p in self.copy.problems()]
        self.assertEqual(stale, ["badges", "evidence:v1.4.0", "summary:v1.4.0"])

    def test_missing_and_unknown_blocks(self) -> None:
        readme = self.copy.text("README.md")
        start = readme.index("<!-- facts:summary:v1.4.0 -->")
        end = readme.index("<!-- /facts:summary:v1.4.0 -->") + len("<!-- /facts:summary:v1.4.0 -->")
        (self.copy.root / "README.md").write_text(
            readme[:start] + "<!-- facts:summary:v0.1.0 -->\n<!-- /facts:summary:v0.1.0 -->"
            + readme[end:], encoding="utf-8")
        self.assertEqual(self.copy.problems(), [
            "README.md: facts:summary:v0.1.0 has no renderer or certified record",
            "README.md: facts:summary:v1.4.0 is missing"])

    def test_tag_binding(self) -> None:
        def moved(tag: str) -> str | None:
            return "f" * 40 if tag == "v1.4.0" else None
        problems, notes = rf.repository_problems(self.copy.root, moved)
        self.assertEqual(problems, [f"{rf.EVIDENCE}: v1.4.0 records commit "
                                    f"{recorded('v1.4.0')}, but the tag points at {'f' * 40}"])
        self.assertEqual(notes, ["v1.3.0 is not present locally; its commit is not compared"])

    def test_unusable_sources_fail(self) -> None:
        self.copy.edit("VERSION", "1.4.0", "next")
        self.assertEqual(self.copy.problems(), ["VERSION: expected X.Y.Z, found 'next'"])

    @staticmethod
    def block(text: str, name: str) -> str:
        return text.split(f"<!-- facts:{name} -->")[1].split(f"<!-- /facts:{name} -->")[0]


class MarkerTests(unittest.TestCase):
    def test_marker_problems(self) -> None:
        cases = {
            "<!-- facts:badges -->\n": ["README.md: facts:badges is never closed"],
            "<!-- /facts:badges -->\n": ["README.md:1: /facts:badges closes no open block"],
            "<!-- facts:a -->\n<!-- facts:b -->\n<!-- /facts:a -->\n":
                ["README.md:2: facts:b opens inside facts:a"],
            "<!-- facts:a -->\n<!-- /facts:a -->\n<!-- facts:a -->\n<!-- /facts:a -->\n":
                ["README.md:3: facts:a appears twice"],
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                self.assertEqual(rf.blocks(text)[1], expected)

    def test_fixture_passes(self) -> None:
        self.assertEqual(rf.fixture_problems(), [])

    def test_shield_escaping(self) -> None:
        self.assertEqual(rf._shield("source support 32 / 64-bit"), "source_support_32_%2F_64--bit")
        self.assertEqual(rf._shield("a_b"), "a__b")


class CertifiedRecordTests(unittest.TestCase):
    def setUp(self) -> None:
        self.data = json.loads((rf.ROOT / rf.EVIDENCE).read_text(encoding="utf-8"))

    def problems(self, **changes: Any) -> list[str]:
        self.data["releases"][0].update(changes)
        return rf.certified_problems(self.data)

    def test_the_committed_file_is_valid(self) -> None:
        self.assertEqual(rf.certified_problems(self.data), [])

    def test_record_rules(self) -> None:
        where = f"{rf.EVIDENCE}: release 0"
        self.assertEqual(self.problems(commit="a5390b4"), [f"{where}: commit 'a5390b4' is malformed"])
        self.setUp()
        self.assertEqual(self.problems(cases=True), [f"{where}: wrong type for cases"])
        self.setUp()
        self.assertEqual(self.problems(failures=700),
                         [f"{where}: failures must lie between 0 and the assertion count"])
        self.setUp()
        self.assertEqual(self.problems(tag="v1.3.0"),
                         [f"{rf.EVIDENCE}: release 1: v1.3.0 is recorded twice"])
        self.setUp()
        del self.data["releases"][0]["excel"]
        self.assertIn("needs exactly", rf.certified_problems(self.data)[0])
        self.assertIn("expected", rf.certified_problems({"releases": []})[0])


class GateRegistryTests(unittest.TestCase):
    def test_names_are_unique_and_the_facts_check_is_last(self) -> None:
        self.assertEqual(len(set(gate_checks.NAMES)), len(gate_checks.NAMES))
        self.assertEqual(gate_checks.NAMES[-1], rf.CHECK)

    def test_module_check_names_are_registered(self) -> None:
        for name in (binary_review.CHECK, markdown_links.CHECK, workflow_pins.CHECK,
                     vba_compile_safety.LABEL_CHECK, vba_compile_safety.ASSIGN_CHECK,
                     vba_lint.INVENTORY_CHECK):
            self.assertIn(name, gate_checks.NAMES)

    def test_the_gate_fails_when_it_runs_unlisted_checks(self) -> None:
        report = vba_lint.Report()
        with contextlib.redirect_stdout(io.StringIO()):
            report.check("an unlisted check", [])
            vba_lint.check_readme_facts(report)
        self.assertFalse(report.results[-1]["passed"])
        self.assertIn("differ from tools/gate_checks.py", str(report.results[-1]["problems"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
