"""Local positive/negative matrix for the paired release-state gate (#50)."""
from __future__ import annotations

import contextlib
import io
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from typing import cast
from unittest import mock

import release_state as rs
import vba_lint

TAGS = ["v1.0.0", "v1.1.0", "v1.2.0", "v1.3.0", "v1.4.0"]
CHANGELOG = "## [Unreleased]\n\n" + "\n".join(
    f"## [{tag[1:]}] — 2026-01-01\nFrozen body {tag}\n" for tag in reversed(TAGS))
CHANGELOG += f"\n[Unreleased]: {rs.REPO_URL}/compare/v1.4.0...main\n"
CHANGELOG += "\n".join(f"[{tag[1:]}]: {rs.REPO_URL}/" +
                       (f"compare/{TAGS[i-1]}...{tag}" if i else f"releases/tag/{tag}")
                       for i, tag in enumerate(TAGS))
README = "The latest tagged release is **v1.4.0**.\n## v1.4.0 — latest tagged release\n"


def init(root: Path) -> None:
    root.mkdir(parents=True, exist_ok=True)
    rs.git(root, "init")
    rs.git(root, "config", "user.email", "fixture@example.invalid")
    rs.git(root, "config", "user.name", "Fixture")


def commit(root: Path) -> str:
    rs.git(root, "add", ".")
    rs.git(root, "commit", "--allow-empty", "-m", "fixture")
    return rs.git(root, "rev-parse", "HEAD")


class RuleTests(unittest.TestCase):
    def test_changelog_positive_and_negative(self) -> None:
        self.assertEqual(rs.changelog_problems(CHANGELOG, TAGS, "1.4.0"), [])
        for old, new in (("[Unreleased]", "[Next]"), ("2026-01-01", "2026-02-30"),
                         ("## [1.4.0]", "## [v1.4.0]")):
            with self.subTest(old=old):
                self.assertTrue(rs.changelog_problems(CHANGELOG.replace(old, new), TAGS, "1.4.0"))
        self.assertTrue(rs.changelog_problems(CHANGELOG + "\n## [1.4.0] — 2026-01-01", TAGS, "1.4.0"))

    def test_comparison_chain_positive_and_each_negative(self) -> None:
        self.assertEqual(rs.comparison_problems(CHANGELOG, TAGS), [])
        for old in ("v1.4.0...main", "v1.3.0...v1.4.0", "v1.2.0...v1.3.0",
                    "v1.1.0...v1.2.0", "v1.0.0...v1.1.0"):
            with self.subTest(old=old):
                self.assertTrue(rs.comparison_problems(CHANGELOG.replace(old, "wrong"), TAGS))
        self.assertTrue(rs.comparison_problems(CHANGELOG.replace("[1.3.0]:", "[v1.3.0]:"), TAGS))

    def test_readme_current_claims_and_historical_text(self) -> None:
        self.assertEqual(rs.readme_problems(README + "Previous v1.3.0 evidence is historical.", TAGS[-1]), [])
        self.assertTrue(rs.readme_problems(README.replace("v1.4.0", "v1.3.0"), TAGS[-1]))
        self.assertTrue(rs.readme_problems("", TAGS[-1]))

    def test_wiki_badges(self) -> None:
        pages = {"Home.md": "https://img.shields.io/badge/Release-1.4.0-blue",
                 "Version-History.md": "https://img.shields.io/badge/Current-1.4.0-blue"}
        self.assertEqual(rs.wiki_badge_problems(pages, TAGS[-1]), [])
        for path in pages:
            with self.subTest(path=path):
                bad = dict(pages)
                bad[path] = bad[path].replace("1.4.0", "1.3.0")
                self.assertTrue(rs.wiki_badge_problems(bad, TAGS[-1]))
        self.assertTrue(rs.wiki_badge_problems({}, TAGS[-1]))

    def test_status_fields_reject_pending_without_rejecting_history(self) -> None:
        pages = {p: rs.status_block(TAGS[-1], 31) + "\nHistorical v1.3.0 was unreleased then."
                 for p in ("README.md", *("wiki/" + p for p in rs.WIKI_PAGES))}
        self.assertEqual(rs.status_problems(pages, TAGS[-1], 31), [])
        for value in ("unreleased", "superseded", "certification pending", "tag pending", "assets pending"):
            bad = dict(pages)
            bad["README.md"] = bad["README.md"].replace("Release status: published", "Release status: " + value)
            self.assertTrue(rs.status_problems(bad, TAGS[-1], 31))
        self.assertTrue(rs.status_problems(pages, TAGS[-1], 30))
        self.assertTrue(rs.status_problems({}, TAGS[-1], 31))
        duplicated = dict(pages)
        duplicated["README.md"] += "\n" + rs.status_block(TAGS[-1], 31)
        self.assertTrue(rs.status_problems(duplicated, TAGS[-1], 31))

    def test_version_missing_malformed_and_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            self.assertIn("missing", rs.source_version_problems(root, {"source": "1.4.0"})[0])
            for value, expected in (("bad", "malformed"), ("9.9.9", "differs")):
                (root / "VERSION").write_text(value)
                self.assertIn(expected, rs.source_version_problems(root, {"source": "1.4.0"})[0])
            (root / "VERSION").write_text("1.4.0\n")
            self.assertEqual(rs.source_version_problems(root, {"source": "1.4.0"}), [])
            (root / "VERSION").write_text("9.9.9\n")
            report = vba_lint.Report()
            with mock.patch.object(vba_lint, "ROOT", root), contextlib.redirect_stdout(io.StringIO()):
                vba_lint.check_version_consistency(report)
            self.assertIn("version stamps agree", report.failures)
            self.assertTrue(any("VERSION" in p for p in cast(list[str], report.results[0]["problems"])))


class PairedInputTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "source"
        self.wiki = Path(self.temp.name) / "wiki"
        init(self.root)
        init(self.wiki)
        (self.root / "tools").mkdir()
        for page in rs.WIKI_PAGES:
            badge = "Release" if page == "Home.md" else "Current"
            (self.wiki / page).write_text(f"https://img.shields.io/badge/{badge}-1.4.0-blue\n" + rs.status_block(TAGS[-1], 31))
        self.wiki_sha = commit(self.wiki)
        self.pin(self.wiki_sha)
        (self.root / "VERSION").write_text("1.4.0\n")
        (self.root / "CHANGELOG.md").write_text(CHANGELOG)
        (self.root / "README.md").write_text(README + rs.status_block(TAGS[-1], 31))
        self.source_sha = commit(self.root)
        for tag in TAGS:
            rs.git(self.root, "tag", tag)

    def pin(self, sha: str) -> None:
        (self.root / rs.PIN).write_text(json.dumps({"schema_version": 1, "wiki_commit": sha}))

    def test_exact_pair_passes_and_json_identifies_both(self) -> None:
        results, metadata = rs.evaluate(self.root, self.wiki, 31)
        self.assertEqual(results, {name: [] for name in rs.CHECKS})
        self.assertEqual(metadata["wiki_commit"], self.wiki_sha)
        self.assertEqual(metadata["source_commit"], self.source_sha)
        report = vba_lint.Report(release_state=metadata)
        for name, problems in results.items():
            with contextlib.redirect_stdout(io.StringIO()):
                report.check(name, problems)
        path = Path(self.temp.name) / "results.json"
        with contextlib.redirect_stdout(io.StringIO()):
            vba_lint.write_json(report, path)
        data = json.loads(path.read_text())
        self.assertEqual(data["release_state"], metadata)
        self.assertEqual(len(data["results"]), 6)

    def test_missing_dirty_and_mismatched_input_fail(self) -> None:
        self.assertTrue(rs.wiki_input(self.root, self.wiki / "missing")[2])
        (self.wiki / "Home.md").write_text("changed")
        self.assertTrue(rs.wiki_input(self.root, self.wiki)[2])
        new = commit(self.wiki)
        self.assertNotEqual(new, self.wiki_sha)
        self.assertTrue(rs.wiki_input(self.root, self.wiki)[2])
        self.pin("bad")
        self.assertTrue(rs.wiki_input(self.root, self.wiki)[2])

    def test_paired_update_sequence_old_pin_then_new_pin(self) -> None:
        (self.wiki / "Home.md").write_text((self.wiki / "Home.md").read_text() + "\nHistorical note.")
        new_sha = commit(self.wiki)
        self.assertTrue(rs.evaluate(self.root, self.wiki, 31)[0][rs.CHECKS[3]])
        self.pin(new_sha)
        results, metadata = rs.evaluate(self.root, self.wiki, 31)
        self.assertFalse(any(results.values()))
        self.assertEqual(metadata["wiki_commit"], new_sha)
        rs.git(self.wiki, "checkout", "--detach", self.wiki_sha)
        self.pin(self.wiki_sha)
        self.assertFalse(any(rs.evaluate(self.root, self.wiki, 31)[0].values()))

    def test_fetch_uses_pinned_revision_not_moving_head(self) -> None:
        (self.wiki / "Home.md").write_text("new moving head must not be fetched")
        self.assertNotEqual(commit(self.wiki), self.wiki_sha)
        destination = Path(self.temp.name) / "fetched"
        with mock.patch.object(rs, "WIKI_URL", str(self.wiki)):
            rs.fetch_wiki(self.root, destination)
        self.assertEqual(rs.git(destination, "rev-parse", "HEAD"), self.wiki_sha)
        with self.assertRaises(FileExistsError):
            rs.fetch_wiki(self.root, destination)
        missing = Path(self.temp.name) / "unavailable"
        with mock.patch.object(rs, "WIKI_URL", str(self.wiki / "absent")), \
                self.assertRaises(subprocess.CalledProcessError):
            rs.fetch_wiki(self.root, missing)


if __name__ == "__main__":
    unittest.main()
