"""Deterministic release-state checks over source and an explicitly pinned wiki.

No network access during checking. --fetch-wiki prepares the pinned read-only
input; failure to obtain/verify it is an error, never an implicit skip.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import subprocess
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
WIKI_URL = "https://github.com/danielep71/VBA-PERFORMANCE_MANAGER.wiki.git"
REPO_URL = "https://github.com/danielep71/vba-performance_manager"
PIN = "tools/release-state.json"
SHA = re.compile(r"[0-9a-f]{40}")
VERSION = re.compile(r"\d+\.\d+\.\d+")
HEADINGS = re.compile(r"^## \[([^\]]+)\](.*)$", re.M)
REFERENCES = re.compile(r"^\[([^\]]+)\]:\s*(\S+)\s*$", re.M)
BLOCK = re.compile(r"<!-- release-state -->\s*\n(.*?)\n<!-- /release-state -->", re.S)
CHECKS = (
    "release changelog sections agree with tags",
    "release comparison links form the tag chain",
    "README current-release claims agree with tags",
    "wiki input matches its pinned revision",
    "wiki current-release badges agree with tags",
    "current-status blocks agree with release evidence",
)
WIKI_PAGES = ("Home.md", "Testing-and-Validation.md", "Version-History.md")


def git(root: Path, *args: str) -> str:
    return subprocess.run(["git", *args], cwd=root, check=True, capture_output=True,
                          text=True, encoding="utf-8", timeout=60).stdout.strip()


def version_key(value: str) -> tuple[int, ...]:
    return tuple(map(int, value.removeprefix("v").split(".")))


def source_version_problems(root: Path, stamps: dict[str, str]) -> list[str]:
    """Keep root VERSION in the same comparison as the existing source stamps."""
    try:
        value = (root / "VERSION").read_text(encoding="utf-8").strip()
    except OSError:
        return ["VERSION: missing or unreadable; expected X.Y.Z"]
    if not VERSION.fullmatch(value):
        return [f"VERSION: malformed {value!r}; expected X.Y.Z"]
    return [f"VERSION: {value} differs from {name}: {stamp}"
            for name, stamp in sorted(stamps.items()) if stamp != value]


def changelog_problems(text: str, tags: list[str], source_version: str) -> list[str]:
    found = HEADINGS.findall(text)
    problems = []
    labels = [label for label, _ in found]
    if not labels or labels[0] != "Unreleased" or labels.count("Unreleased") != 1:
        problems.append("CHANGELOG.md: one [Unreleased] must precede released sections")
    for version in {tag[1:] for tag in tags} | {source_version}:
        matches = [tail for label, tail in found if label == version]
        if len(matches) != 1:
            problems.append(f"CHANGELOG.md: expected one dated [{version}] section")
            continue
        date = re.fullmatch(r"\s+[—-]\s+(\d{4}-\d{2}-\d{2})\s*", matches[0])
        try:
            if not date:
                raise ValueError
            dt.date.fromisoformat(date.group(1))
        except ValueError:
            problems.append(f"CHANGELOG.md: [{version}] needs a valid ISO release date")
    for label in labels:
        if label != "Unreleased" and (not VERSION.fullmatch(label) or labels.count(label) != 1):
            problems.append(f"CHANGELOG.md: malformed or duplicate release heading [{label}]")
    return problems


def comparison_problems(text: str, tags: list[str]) -> list[str]:
    refs = REFERENCES.findall(text)
    expected = {"Unreleased": f"{REPO_URL}/compare/{tags[-1]}...main"}
    for index, tag in enumerate(tags):
        expected[tag[1:]] = (f"{REPO_URL}/compare/{tags[index - 1]}...{tag}" if index
                             else f"{REPO_URL}/releases/tag/{tag}")
    problems = []
    for label, url in expected.items():
        actual = [link for name, link in refs if name == label]
        if len(actual) != 1 or actual[0].lower() != url.lower():
            problems.append(f"CHANGELOG.md: [{label}] must reference {url}")
    known = set(expected)
    for label, _ in refs:
        if VERSION.fullmatch(label) and label not in known:
            problems.append(f"CHANGELOG.md: [{label}] reference has no matching release tag")
    return problems


def readme_problems(text: str, latest: str) -> list[str]:
    claims = re.findall(r"latest tagged release is \*\*(v\d+\.\d+\.\d+)\*\*", text, re.I)
    headings = re.findall(r"^## (v\d+\.\d+\.\d+) — latest tagged release\s*$", text, re.M)
    problems = []
    for name, values in (("latest-release note", claims), ("latest-release heading", headings)):
        if values != [latest]:
            problems.append(f"README.md: {name} must name {latest} exactly once; found {values}")
    return problems


def wiki_badge_problems(pages: dict[str, str], latest: str) -> list[str]:
    problems = []
    for path, label in (("Home.md", "Release"), ("Version-History.md", "Current")):
        values = re.findall(rf"img\.shields\.io/badge/{label}-([^-?]+)-", pages.get(path, ""))
        if values != [latest[1:]]:
            problems.append(f"wiki/{path}: {label} badge must name {latest}; found {values}")
    return problems


def status_block(latest: str, checks: int) -> str:
    return ("<!-- release-state -->\n"
            f"Latest release: {latest}\n"
            "Release status: published\n"
            f"Development static checks: {checks}\n"
            "Execution evidence: 64-bit only\n"
            "32-bit follow-up: #29, v1.5.0, contributor-dependent\n"
            "<!-- /release-state -->")


def status_problems(pages: dict[str, str], latest: str, checks: int) -> list[str]:
    """Bounded exact fields; historical prose outside the block is not scanned."""
    expected = BLOCK.fullmatch(status_block(latest, checks)).group(1)  # type: ignore[union-attr]
    problems = []
    for path in ("README.md", *("wiki/" + p for p in WIKI_PAGES)):
        text = pages.get(path, "")
        blocks = BLOCK.findall(text)
        if blocks != [expected] or text.count("<!-- release-state -->") != 1 \
                or text.count("<!-- /release-state -->") != 1:
            problems.append(f"{path}: current-status block must record {latest} as published, "
                            f"{checks} development checks and the documented bitness boundary")
    return problems


def pin(root: Path) -> str:
    data: Any = json.loads((root / PIN).read_text(encoding="utf-8"))
    if not isinstance(data, dict) or set(data) != {"schema_version", "wiki_commit"} \
            or data["schema_version"] != 1 or not isinstance(data["wiki_commit"], str) \
            or not SHA.fullmatch(data["wiki_commit"]):
        raise ValueError(f"{PIN}: expected schema_version=1 and a full lowercase wiki_commit")
    return str(data["wiki_commit"])


def wiki_input(root: Path, wiki: Path) -> tuple[dict[str, str], str | None, list[str]]:
    observed = None
    try:
        expected = pin(root)
        observed = git(wiki, "rev-parse", "HEAD")
        if observed != expected:
            raise ValueError(f"wiki revision {observed} differs from pin {expected}")
        if git(wiki, "status", "--porcelain", "--untracked-files=all"):
            raise ValueError("wiki checkout is dirty; cannot certify pinned content")
        pages = {p: (wiki / p).read_text(encoding="utf-8") for p in WIKI_PAGES}
        return pages, observed, []
    except (OSError, ValueError, subprocess.SubprocessError) as exc:
        return {}, observed, [f"{PIN}: wiki input unavailable or unverified: {exc}"]


def evaluate(root: Path, wiki: Path, check_count: int) -> tuple[dict[str, list[str]], dict[str, Any]]:
    results: dict[str, list[str]] = {name: [] for name in CHECKS}
    metadata: dict[str, Any] = {"wiki_commit": None, "source_commit": None}
    try:
        metadata["source_commit"] = git(root, "rev-parse", "HEAD")
        tags = sorted((t for t in git(root, "tag", "--list").splitlines()
                       if re.fullmatch(r"v\d+\.\d+\.\d+", t)), key=version_key)
        if not tags:
            raise ValueError("no release tags; fetch tags before checking")
        latest = tags[-1]
        metadata["latest_tag"] = latest
        metadata["latest_tag_commit"] = git(root, "rev-parse", latest + "^{commit}")
        source_version = (root / "VERSION").read_text(encoding="utf-8").strip()
        changelog = (root / "CHANGELOG.md").read_text(encoding="utf-8")
        readme = (root / "README.md").read_text(encoding="utf-8")
        results[CHECKS[0]] = changelog_problems(changelog, tags, source_version)
        results[CHECKS[1]] = comparison_problems(changelog, tags)
        results[CHECKS[2]] = readme_problems(readme, latest)
        pages, observed, problems = wiki_input(root, wiki)
        metadata["wiki_commit"] = observed
        metadata["expected_wiki_commit"] = pin(root)
        results[CHECKS[3]] = problems
        results[CHECKS[4]] = wiki_badge_problems(pages, latest)
        results[CHECKS[5]] = status_problems({"README.md": readme,
                                            **{"wiki/" + k: v for k, v in pages.items()}},
                                           latest, check_count)
    except (OSError, ValueError, subprocess.SubprocessError) as exc:
        for name in CHECKS:
            results[name].append(f"release-state input error: {exc}")
    return results, metadata


def fetch_wiki(root: Path, destination: Path) -> None:
    """Prepare a new isolated checkout; never alter a user's existing checkout."""
    expected = pin(root)
    destination.mkdir(parents=True, exist_ok=False)
    git(destination, "init")
    git(destination, "fetch", "--depth=1", WIKI_URL, expected)
    git(destination, "checkout", "--detach", "FETCH_HEAD")
    if git(destination, "rev-parse", "HEAD") != expected:
        raise ValueError("fetched wiki revision differs from pin")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fetch-wiki", action="store_true", required=True)
    parser.add_argument("--wiki-root", type=Path, required=True)
    args = parser.parse_args()
    try:
        fetch_wiki(ROOT, args.wiki_root)
        print(f"Prepared wiki {pin(ROOT)} at {args.wiki_root}")
        return 0
    except (OSError, ValueError, subprocess.SubprocessError) as exc:
        print(f"Cannot prepare pinned wiki: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
