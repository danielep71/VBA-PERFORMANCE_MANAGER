"""Fixtures for the Markdown link check (#85)."""
from __future__ import annotations

import http.server
import threading
import unittest
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import markdown_links as ml

ROOT = Path(__file__).resolve().parent.parent


class SlugTests(unittest.TestCase):
    def test_github_heading_slugs(self) -> None:
        cases = {
            "Release status": "release-status",
            "🧪 Demo workbook": "-demo-workbook",
            "`TotalSteps` and counts": "totalsteps-and-counts",
            "What's new? (v1.4.1)": "whats-new-v141",
            "[Linked](README.md) heading": "linked-heading",
            "A &amp; B": "a--b",
            "Ünïcode Héading": "ünïcode-héading",
            "snake_case name": "snake_case-name",
        }
        for heading, expected in cases.items():
            with self.subTest(heading=heading):
                self.assertEqual(ml.slug(heading), expected)

    def test_repeated_headings_get_numbered(self) -> None:
        self.assertEqual(ml.anchors("# Notes\n## Notes\n### Notes\n"), {"notes", "notes-1", "notes-2"})

    def test_inline_code_headings_keep_text_and_duplicate_suffixes(self) -> None:
        text = ("## `LastReadStatus`\n## LastReadStatus\n"
                "## `LastReadStatus`\n`TotalSteps` and counts\n===\n")
        self.assertEqual(ml.anchors(text),
                         {"lastreadstatus", "lastreadstatus-1", "lastreadstatus-2",
                          "totalsteps-and-counts"})

    def test_literal_html_anchors_and_fenced_headings_are_excluded(self) -> None:
        text = ('`<a id="fake"></a>`\n<a id="real"></a>\n'
                '~~~markdown\n## `LastReadStatus`\n~~~\n'
                '## `LastReadStatus`\n')
        self.assertEqual(ml.anchors(text), {"real", "lastreadstatus"})

    def test_explicit_and_setext_anchors(self) -> None:
        text = '<a id="date-and-version-rules"></a>\n<a name="Legacy"></a>\nTitle here\n==========\n'
        self.assertEqual(ml.anchors(text), {"date-and-version-rules", "legacy", "title-here"})

    def test_headings_inside_code_blocks_are_not_anchors(self) -> None:
        self.assertEqual(ml.anchors("```\n# not a heading\n```\n# Real\n"), {"real"})


class LinkExtractionTests(unittest.TestCase):
    def test_link_forms(self) -> None:
        text = ('[a](docs/A.md) ![img](images/x.png) [b](<docs/B C.md>) [c](d(1).md)\n'
                '[ref]: RELEASING.md#step\n<a href="NOTICE">n</a> <img src="images/y.png">\n')
        self.assertEqual([d for _, d in ml.links(text)],
                         ["docs/A.md", "images/x.png", "docs/B C.md", "d(1).md",
                          "RELEASING.md#step", "NOTICE", "images/y.png"])

    def test_code_is_not_a_link(self) -> None:
        self.assertEqual(ml.links("`[x](nope.md)`\n```\n[y](nope.md)\n```\n"), [])


class OfflineTests(unittest.TestCase):
    FILES = {
        "README.md": "# Top\n## Release status\n[ok](docs/Guide.md#install) [dir](docs) "
                     "[self](#release-status) [img](images/a.png) [web](https://example.com/x)\n",
        "docs/Guide.md": "# Guide\n## Install\n[up](../README.md#top) [mail](mailto:a@b.c)\n",
    }
    EXISTING = {"README.md", "docs", "docs/Guide.md", "images", "images/a.png"}

    def test_valid_links_pass(self) -> None:
        self.assertEqual(ml.offline_problems(self.FILES, self.EXISTING), [])

    def test_inline_code_heading_links_resolve_end_to_end(self) -> None:
        files = {
            "README.md": "[first](Guide.md#lastreadstatus) "
                         "[second](Guide.md#lastreadstatus-1) "
                         "[setext](Guide.md#totalsteps-and-counts)\n",
            "Guide.md": "## `LastReadStatus`\n## `LastReadStatus`\n"
                        "`TotalSteps` and counts\n---\n"
                        "`[ignored](missing.md)`\n"
                        "```markdown\n## `LastReadStatus`\n[x](missing.md)\n```\n"
                        "[self](#lastreadstatus)\n",
        }
        self.assertEqual(ml.offline_problems(files, set(files)), [])
        files["README.md"] += "[bad](Guide.md#lastreadstatus-2)\n"
        self.assertEqual(ml.offline_problems(files, set(files)),
                         ["README.md:2: missing anchor #lastreadstatus-2 in Guide.md"])

    def test_repository_links_resolve(self) -> None:
        self.assertEqual(ml.repository_problems(ROOT), [])

    def test_broken_links_fail(self) -> None:
        files = {"README.md": "# Top\n[a](docs/Missing.md) [b](#nowhere) [c](docs/Guide.md#gone) "
                              "[d](../outside.md) [e](%20spaced.md)\n",
                 "docs/Guide.md": "# Guide\n"}
        problems = ml.offline_problems(files, {"README.md", "docs", "docs/Guide.md"})
        self.assertEqual(problems, [
            "README.md:2: missing link target: docs/Missing.md",
            "README.md:2: missing anchor #nowhere in README.md",
            "README.md:2: missing anchor #gone in docs/Guide.md",
            "README.md:2: link leaves the repository: ../outside.md",
            "README.md:2: missing link target: %20spaced.md",
        ])

    def test_anchor_match_ignores_case(self) -> None:
        files = {"README.md": "# Release Status\n[a](#Release-Status)\n"}
        self.assertEqual(ml.offline_problems(files, {"README.md"}), [])


class _Handler(http.server.BaseHTTPRequestHandler):
    def do_HEAD(self) -> None:
        self._answer(head=True)

    def do_GET(self) -> None:
        self._answer(head=False)

    def _answer(self, head: bool) -> None:
        if self.path == "/ok":
            code = 200
        elif self.path == "/get-only":
            code = 405 if head else 200
        else:
            code = 404
        self.send_response(code)
        self.end_headers()

    def log_message(self, *args: object) -> None:
        pass


@contextmanager
def server() -> Iterator[str]:
    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{httpd.server_address[1]}"
    finally:
        httpd.shutdown()
        httpd.server_close()


class ExternalTests(unittest.TestCase):
    def test_external_links_are_requested(self) -> None:
        with server() as base:
            files = {"README.md": f"[a]({base}/ok) [b]({base}/get-only) [c]({base}/gone)\n"
                                  f"[d]({base}/gone#x)\n"}
            problems = ml.external_problems(files)
        self.assertEqual(problems, [f"README.md:1: HTTP 404: {base}/gone (and 1 more)"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
