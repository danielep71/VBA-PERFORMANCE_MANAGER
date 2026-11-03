"""Check the links in the tracked Markdown files (#85).

Offline, on every change, through tools/vba_lint.py:

- a relative link must name a file or directory that exists in the
  repository, and must not leave it;
- a fragment (``#anchor``) must name a heading or an explicit ``<a id>`` /
  ``<a name>`` anchor of the Markdown file it points to, using the slug rules
  GitHub applies to headings (lower case, punctuation and symbols removed,
  spaces to hyphens, ``-1``, ``-2`` ... for repeated headings).

With ``--external`` the http(s) links are also requested once each, as the
weekly link-check workflow does; external sites change without a commit, so
that part never blocks a merge. Links inside code spans and fenced code blocks
are not links and are ignored.

Run ``python tools/markdown_links.py [--external]`` from anywhere.
"""
from __future__ import annotations

import argparse
import html
import re
import subprocess
import sys
import unicodedata
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parent.parent
CHECK = "Markdown links and anchors resolve"

FENCE = re.compile(r"^ {0,3}(`{3,}|~{3,})")
HEADING = re.compile(r"^ {0,3}(#{1,6})[ \t]+(.*?)[ \t#]*$")
SETEXT = re.compile(r"^ {0,3}(=+|-+)[ \t]*$")
CODE_SPAN = re.compile(r"(`+)(?:(?!\1).)+?\1")
INLINE = re.compile(r"!?\[(?:[^\[\]]|\[[^\]]*\])*\]\(\s*(<[^>]*>|[^\s()]*(?:\([^\s()]*\)[^\s()]*)*)")
REFERENCE = re.compile(r"^ {0,3}\[[^\]]+\]:\s*(<[^>]*>|\S+)")
HTML_ATTR = re.compile(r"""<(?:a|img|source)\b[^>]*?\b(?:href|src|srcset)\s*=\s*["']([^"']+)["']""", re.I)
ANCHOR_ATTR = re.compile(r"""<a\b[^>]*?\b(?:id|name)\s*=\s*["']([^"']+)["']""", re.I)
USER_AGENT = "VBA-Performance-Manager-link-check"


def slug(heading: str) -> str:
    """The anchor GitHub generates for a heading, before duplicate suffixes."""
    text = re.sub(r"!?\[([^\]]*)\]\([^)]*\)", r"\1", heading)   # links keep their text
    text = re.sub(r"<[^>]+>", "", text)                         # inline HTML is dropped
    text = html.unescape(text).replace("`", "")
    kept = (c for c in text.lower()
            if c in "- " or unicodedata.category(c)[0] in "LM"
            or unicodedata.category(c) in ("Nd", "Pc"))
    return "".join(kept).replace(" ", "-")


def _lines_outside_code(text: str) -> list[tuple[int, str]]:
    """Numbered lines outside fenced code blocks, with code spans blanked."""
    out: list[tuple[int, str]] = []
    fence: str | None = None
    for number, line in enumerate(text.splitlines(), 1):
        match = FENCE.match(line)
        if fence is None and match:
            fence = match.group(1)
            continue
        if fence is not None:
            if match and match.group(1)[0] == fence[0] and len(match.group(1)) >= len(fence):
                fence = None
            continue
        out.append((number, CODE_SPAN.sub(lambda m: " " * len(m.group(0)), line)))
    return out


def anchors(text: str) -> set[str]:
    """Every anchor a Markdown file defines: heading slugs and explicit ids."""
    found: set[str] = set()
    used: dict[str, int] = {}
    previous = ""
    for _, line in _lines_outside_code(text):
        found.update(a.lower() for a in ANCHOR_ATTR.findall(line))
        match = HEADING.match(line)
        title = match.group(2) if match else (previous if previous.strip() and SETEXT.match(line)
                                               and not HEADING.match(previous) else None)
        if title is not None:
            base = slug(title)
            anchor, count = base, used.get(base, -1)
            if count >= 0:
                while anchor in used:
                    count += 1
                    anchor = f"{base}-{count}"
                used[base] = count
            used.setdefault(anchor, 0)
            found.add(anchor)
        previous = line
    return found


def links(text: str) -> list[tuple[int, str]]:
    """Every link destination in a Markdown file, with its line number."""
    found: list[tuple[int, str]] = []
    for number, line in _lines_outside_code(text):
        for pattern in (INLINE, REFERENCE, HTML_ATTR):
            for match in pattern.finditer(line):
                destination = match.group(1).strip()
                if destination.startswith("<") and destination.endswith(">"):
                    destination = destination[1:-1]
                if pattern is HTML_ATTR and "," in destination:   # srcset lists
                    found.extend((number, part.split()[0]) for part in destination.split(",")
                                 if part.strip())
                elif destination:
                    found.append((number, destination))
    return found


def offline_problems(files: dict[str, str], existing: set[str]) -> list[str]:
    """Check relative links and anchors.

    files maps each Markdown path (POSIX, relative to the root) to its text;
    existing holds every file and directory path in the repository.
    """
    defined = {path: anchors(text) for path, text in files.items()}
    problems: list[str] = []
    for path, text in sorted(files.items()):
        base = PurePosixPath(path).parent
        for number, destination in links(text):
            parts = urlsplit(destination)
            if parts.scheme or destination.startswith("//"):
                continue   # external: checked by --external
            where = f"{path}:{number}"
            target = unquote(parts.path)
            if target:
                joined = PurePosixPath(*(base / target).parts)
                resolved: list[str] = []
                for part in joined.parts:
                    if part == "..":
                        if not resolved:
                            resolved = ["<outside>"]
                            break
                        resolved.pop()
                    elif part != ".":
                        resolved.append(part)
                name = "/".join(resolved)
                if resolved[:1] == ["<outside>"]:
                    problems.append(f"{where}: link leaves the repository: {destination}")
                    continue
                if name not in existing:
                    problems.append(f"{where}: missing link target: {destination}")
                    continue
            else:
                name = path
            fragment = unquote(parts.fragment)
            if fragment and name.lower().endswith(".md"):
                if name not in defined:
                    continue   # an untracked Markdown target: nothing to check against
                if fragment.lower() not in defined[name]:
                    problems.append(f"{where}: missing anchor #{fragment} in {name}")
    return problems


def repository_markdown(root: Path) -> tuple[dict[str, str], set[str]]:
    """The tracked Markdown files and every tracked file and directory path."""
    listing = subprocess.run(["git", "ls-files", "-z"], cwd=root, capture_output=True,
                             check=True).stdout.decode("utf-8", errors="replace")
    tracked = [name for name in listing.split("\0") if name]
    existing = set(tracked)
    for name in tracked:
        existing.update(str(parent) for parent in PurePosixPath(name).parents if str(parent) != ".")
    files = {name: (root / name).read_text(encoding="utf-8")
             for name in tracked if name.lower().endswith(".md") and (root / name).is_file()}
    return files, existing


def repository_problems(root: Path = ROOT) -> list[str]:
    try:
        files, existing = repository_markdown(root)
    except (OSError, subprocess.CalledProcessError) as exc:
        return [f"git could not list the Markdown files, so links were not checked: {exc}"]
    return offline_problems(files, existing)


def _request(url: str) -> str | None:
    """Return a problem for one external URL, or None when it answers."""
    for method in ("HEAD", "GET"):
        request = urllib.request.Request(url, method=method, headers={"User-Agent": USER_AGENT})
        try:
            with urllib.request.urlopen(request, timeout=20):
                return None
        except urllib.error.HTTPError as exc:
            if method == "HEAD" and exc.code in (403, 404, 405, 429, 501):
                continue   # some servers refuse HEAD; retry once with GET
            if exc.code == 429:
                return None   # rate limited: reachable, not broken
            return f"HTTP {exc.code}"
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            return f"unreachable: {exc}"
    return None


def external_problems(files: dict[str, str]) -> list[str]:
    urls: dict[str, list[str]] = {}
    for path, text in sorted(files.items()):
        for number, destination in links(text):
            if urlsplit(destination).scheme in ("http", "https"):
                urls.setdefault(destination.split("#", 1)[0], []).append(f"{path}:{number}")
    with ThreadPoolExecutor(max_workers=8) as pool:
        results = dict(zip(urls, pool.map(_request, urls)))
    return [f"{urls[url][0]}: {result}: {url}" + (f" (and {len(urls[url]) - 1} more)"
                                                  if len(urls[url]) > 1 else "")
            for url, result in sorted(results.items()) if result]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--external", action="store_true", help="also request external links")
    args = parser.parse_args(argv)
    problems = repository_problems(ROOT)
    if args.external and not problems:
        files, _ = repository_markdown(ROOT)
        problems = external_problems(files)
    for problem in problems:
        print(problem)
    print(f"{len(problems)} link problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
