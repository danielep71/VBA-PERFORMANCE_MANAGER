"""Render the README's version, regression and static-check facts (#89).

A rendered block sits between two markers in README.md:

    <!-- facts:NAME -->
    ...
    <!-- /facts:NAME -->

Two kinds of fact are kept apart:

- Derived facts come from the current source and change with it: the version
  in VERSION, the number of regression cases the test module invokes, and the
  static gate's checks listed in tools/gate_checks.py.
- Certified facts are a published tag's Excel evidence: the commit, Excel build,
  bitness, cases, assertions, failures, workbook and manifest digests. They come
  only from tools/certified-evidence.json, one record per tag, which this tool
  reads and never writes. Each record must name its tag's commit when the tag
  is present locally, so a record cannot drift from the release it describes.

Blocks:

- ``badges``: the README's badge row. Version and static checks are derived;
  the regression badge is the latest certified record, and says which tag.
- ``source-cases``: the case count of the regression module in the source.
- ``static-checks``: the gate's check count and numbered list.
- ``evidence:vX.Y.Z``: a tag's certified evidence table.
- ``summary:vX.Y.Z``: a tag's one-sentence certification summary.

The default run fails when a block differs from a fresh render, when markers
are missing, unbalanced or repeated, or when a marker names no renderer;
``--write`` replaces each block with its fresh render. The static gate runs the
check as "README facts match their sources".
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import textwrap
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from gate_checks import CHECKS

ROOT = Path(__file__).resolve().parent.parent
README = "README.md"
EVIDENCE = "tools/certified-evidence.json"
TEST_MODULE = "test/M_cPM_Test.bas"
REPO_URL = "https://github.com/danielep71/VBA-PERFORMANCE_MANAGER"
CHECK = "README facts match their sources"

MARKER = re.compile(r"^<!-- (/?)facts:([a-z0-9-]+(?::v\d+\.\d+\.\d+)?) -->[ \t]*$", re.M)
VERSION_TEXT = re.compile(r"\d+\.\d+\.\d+")
INVOKED_CASE = re.compile(r"^\s+(Test_[A-Za-z_0-9]+)\s*$", re.M)
TAG = re.compile(r"v\d+\.\d+\.\d+")
SHA1 = re.compile(r"[0-9a-f]{40}")
SHA256 = re.compile(r"[0-9a-f]{64}")
DAY = re.compile(r"\d{4}-\d{2}-\d{2}")

# Every key a certified record holds, with the type it must have.
RECORD: dict[str, type | tuple[type, ...]] = {
    "tag": str, "commit": str, "certified": str, "excel": str, "bitness": str,
    "cases": int, "assertions": int, "failures": int, "static_checks": (int, type(None)),
    "workbook": str, "workbook_sha256": str, "manifest_sha256": str,
}

# resolve(tag) -> the tag's commit, or None when the tag is not present locally
Resolve = Callable[[str], "str | None"]


@dataclass(frozen=True)
class Facts:
    version: str
    cases: int
    checks: tuple[tuple[str, str], ...]
    releases: tuple[dict[str, Any], ...]


# --------------------------------------------------------------------------- #
# Sources
# --------------------------------------------------------------------------- #

def certified_problems(data: Any) -> list[str]:
    """Problems in the certified-evidence file's content."""
    if not isinstance(data, dict) or data.get("schema_version") != 1 \
            or not isinstance(data.get("releases"), list) or set(data) != {"schema_version", "releases"}:
        return [f'{EVIDENCE}: expected {{"schema_version": 1, "releases": [...]}}']
    problems: list[str] = []
    seen: set[str] = set()
    for index, record in enumerate(data["releases"]):
        where = f"{EVIDENCE}: release {index}"
        if not isinstance(record, dict) or set(record) != set(RECORD):
            problems.append(f"{where} needs exactly {', '.join(RECORD)}")
            continue
        wrong = [key for key, kind in RECORD.items()
                 if not isinstance(record[key], kind) or isinstance(record[key], bool)]
        if wrong:
            problems.append(f"{where}: wrong type for {', '.join(wrong)}")
            continue
        for key, pattern in (("tag", TAG), ("commit", SHA1), ("certified", DAY),
                             ("workbook_sha256", SHA256), ("manifest_sha256", SHA256)):
            if not pattern.fullmatch(record[key]):
                problems.append(f"{where}: {key} {record[key]!r} is malformed")
        if record["tag"] in seen:
            problems.append(f"{where}: {record['tag']} is recorded twice")
        seen.add(record["tag"])
        if record["failures"] < 0 or record["failures"] > record["assertions"]:
            problems.append(f"{where}: failures must lie between 0 and the assertion count")
    return problems


def tag_binding(releases: tuple[dict[str, Any], ...],
                resolve: Resolve) -> tuple[list[str], list[str]]:
    """(problems, notes): each record must name its tag's commit."""
    problems: list[str] = []
    notes: list[str] = []
    for record in releases:
        commit = resolve(record["tag"])
        if commit is None:
            notes.append(f"{record['tag']} is not present locally; its commit is not compared")
        elif commit != record["commit"]:
            problems.append(f"{EVIDENCE}: {record['tag']} records commit {record['commit']}, "
                            f"but the tag points at {commit}")
    return problems, notes


def git_resolve(root: Path) -> Resolve:
    def resolve(tag: str) -> str | None:
        run = subprocess.run(["git", "rev-parse", "--verify", "--quiet", f"refs/tags/{tag}^{{commit}}"],
                             cwd=root, capture_output=True, text=True)
        return run.stdout.strip() if run.returncode == 0 else None
    return resolve


def load_facts(root: Path) -> tuple[Facts | None, list[str]]:
    """Read every source a renderer needs; (None, problems) when one is unusable."""
    problems: list[str] = []
    version = (root / "VERSION").read_text(encoding="utf-8").strip() \
        if (root / "VERSION").is_file() else ""
    if not VERSION_TEXT.fullmatch(version):
        problems.append(f"VERSION: expected X.Y.Z, found {version!r}")
    test = root / TEST_MODULE
    cases = len(set(INVOKED_CASE.findall(test.read_text(encoding="ascii", errors="replace")))) \
        if test.is_file() else 0
    if not cases:
        problems.append(f"{TEST_MODULE}: no invoked regression case found")
    try:
        data = json.loads((root / EVIDENCE).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return None, problems + [f"{EVIDENCE}: unreadable: {exc}"]
    problems += certified_problems(data)
    if problems:
        return None, problems
    return Facts(version, cases, CHECKS, tuple(data["releases"])), []


# --------------------------------------------------------------------------- #
# Renderers
# --------------------------------------------------------------------------- #

def _shield(text: str) -> str:
    """Escape text for a shields.io static badge path segment."""
    return text.replace("-", "--").replace("_", "__").replace(" ", "_").replace("/", "%2F")


def _badge(label: str, message: str, color: str, target: str, logo: str = "",
           alt: str = "") -> str:
    style = "style=for-the-badge" + (f"&logo={logo}&logoColor=white" if logo else "")
    return (f"[![{alt or label}](https://img.shields.io/badge/{_shield(label)}-"
            f"{_shield(message)}-{color}?{style})]({target})")


def badges(facts: Facts) -> str:
    latest = facts.releases[0]
    passed = latest["assertions"] - latest["failures"]
    total = len(facts.checks)
    return "\n".join([
        _badge("Excel VBA", "source support 32 / 64-bit", "217346", "#requirements",
               "microsoft-excel"),
        _badge("Platform", "Windows desktop", "0078D6", "#requirements", "windows", "Windows"),
        _badge("Version", facts.version, "6f42c1", "#release-status"),
        _badge(f"Regression {latest['tag']}", f"{passed} / {latest['assertions']}", "2ea44f",
               "#regression-testing"),
        _badge("Static checks", f"{total} / {total}", "0969da", "#static-source-analysis"),
        _badge("License", "Apache 2.0", "2ea44f", "LICENSE"),
    ])


def _wrap(text: str) -> str:
    return textwrap.fill(text, 79, break_long_words=False, break_on_hyphens=False)


def source_cases(facts: Facts) -> str:
    return _wrap(f"The regression module in the current source invokes **{facts.cases} cases**, "
                 "and `TotalSteps` must equal that count. Assertion counts are not derived "
                 "from the source: they come from an Excel run, recorded below for each "
                 "published tag.")


def static_checks(facts: Facts) -> str:
    items = [f"{index}. {summary}{'.' if index == len(facts.checks) else ';'}"
             for index, (_, summary) in enumerate(facts.checks, 1)]
    return "\n".join([f"The current development linter performs **{len(facts.checks)} checks**:",
                      "", *items])


def _record(facts: Facts, tag: str) -> dict[str, Any] | None:
    return next((r for r in facts.releases if r["tag"] == tag), None)


def evidence(record: dict[str, Any]) -> str:
    commit = record["commit"]
    rows = [
        f"| Evidence | {record['tag']} result |", "|---|---|",
        f"| Tag target | [`{commit}`]({REPO_URL}/commit/{commit}) |",
        f"| Certified | {record['certified']} |",
        f"| Excel | {record['excel']} |",
        f"| Office bitness | **{record['bitness']}** |",
        f"| Regression | **{record['cases']} cases · {record['assertions']} assertions · "
        f"{record['failures']} failures** |",
    ]
    if record["static_checks"] is not None:
        rows.append(f"| Static checks | {record['static_checks']} checks · 0 failures |")
    rows += [
        f"| Release workbook | `{record['workbook']}` |",
        f"| Workbook SHA-256 | `{record['workbook_sha256']}` |",
        f"| Manifest SHA-256 | `{record['manifest_sha256']}` |",
    ]
    return "\n".join(rows)


def summary(record: dict[str, Any]) -> str:
    checks = record["static_checks"]
    static = f", plus **{checks}/{checks}** static checks" if checks is not None else ""
    return _wrap(f"The release is bound to tag target `{record['commit']}` and certified on "
                 f"{record['bitness']} Excel with **{record['cases']} cases, "
                 f"{record['assertions']} assertions and {record['failures']} failures**"
                 f"{static}.")


DERIVED: dict[str, Callable[[Facts], str]] = {
    "badges": badges, "source-cases": source_cases, "static-checks": static_checks,
}
CERTIFIED: dict[str, Callable[[dict[str, Any]], str]] = {"evidence": evidence, "summary": summary}


def render(name: str, facts: Facts) -> str | None:
    """The block's fresh content, or None when no renderer or record matches."""
    if name in DERIVED:
        return DERIVED[name](facts)
    kind, _, tag = name.partition(":")
    record = _record(facts, tag) if kind in CERTIFIED and tag else None
    return CERTIFIED[kind](record) if record is not None else None


# --------------------------------------------------------------------------- #
# Blocks
# --------------------------------------------------------------------------- #

def blocks(text: str) -> tuple[list[tuple[str, int, int]], list[str]]:
    """(name, content start, content end) per block, and marker problems."""
    found: list[tuple[str, int, int]] = []
    problems: list[str] = []
    seen: set[str] = set()
    open_name: str | None = None
    start = 0
    for match in MARKER.finditer(text):
        closing, name = match.group(1) == "/", match.group(2)
        line = text.count("\n", 0, match.start()) + 1
        if not closing:
            if open_name is not None:
                problems.append(f"{README}:{line}: facts:{name} opens inside facts:{open_name}")
                continue
            if name in seen:
                problems.append(f"{README}:{line}: facts:{name} appears twice")
            seen.add(name)
            open_name, start = name, match.end()
        elif name != open_name:
            problems.append(f"{README}:{line}: /facts:{name} closes no open block")
        else:
            found.append((name, start, match.start()))
            open_name = None
    if open_name is not None:
        problems.append(f"{README}: facts:{open_name} is never closed")
    return found, problems


def _fresh(text: str, facts: Facts) -> tuple[str, list[str], list[str]]:
    """(text with every block rendered, problems, names of stale blocks)."""
    found, problems = blocks(text)
    stale: list[str] = []
    out: list[str] = []
    position = 0
    for name, start, end in found:
        rendered = render(name, facts)
        if rendered is None:
            problems.append(f"{README}: facts:{name} has no renderer or certified record")
            continue
        content = "\n" + rendered + "\n"
        if text[start:end] != content:
            stale.append(name)
        out += [text[position:start], content]
        position = end
    out.append(text[position:])
    return "".join(out), problems, stale


def text_problems(text: str, facts: Facts, required: tuple[str, ...] = ()) -> list[str]:
    """Every problem in a README text against the facts."""
    found, _ = blocks(text)
    present = {name for name, _, _ in found}
    _, problems, stale = _fresh(text, facts)
    problems += [f"{README}: facts:{name} is missing" for name in required if name not in present]
    problems += [f"{README}: facts:{name} is stale; run python3 tools/readme_facts.py --write"
                 for name in stale]
    return problems


def required_blocks(facts: Facts) -> tuple[str, ...]:
    """The blocks the README must carry: every derived one and the latest tag's."""
    latest = facts.releases[0]["tag"]
    return (*DERIVED, f"evidence:{latest}", f"summary:{latest}")


def repository_problems(root: Path = ROOT, resolve: Resolve | None = None) -> tuple[list[str], list[str]]:
    """(problems, notes) for the repository's README."""
    facts, problems = load_facts(root)
    if facts is None:
        return problems, []
    problems, notes = tag_binding(facts.releases, resolve or git_resolve(root))
    text = (root / README).read_text(encoding="utf-8")
    return problems + text_problems(text, facts, required_blocks(facts)), notes


def fixture_problems() -> list[str]:
    """The check must catch a stale block and accept a fresh one."""
    record = {"tag": "v9.9.9", "commit": "0" * 40, "certified": "2026-01-01", "excel": "x",
              "bitness": "64-bit", "cases": 3, "assertions": 9, "failures": 0,
              "static_checks": 2, "workbook": "w.xlsm", "workbook_sha256": "a" * 64,
              "manifest_sha256": "b" * 64}
    facts = Facts("9.9.9", 3, (("a", "first"), ("b", "second")), (record,))
    fresh = "<!-- facts:static-checks -->\n" + static_checks(facts) + "\n<!-- /facts:static-checks -->\n"
    stale = fresh.replace("**2 checks**", "**1 checks**")
    problems = []
    if text_problems(fresh, facts):
        problems.append("fixture: a fresh block was reported")
    if not any("is stale" in p for p in text_problems(stale, facts)):
        problems.append("fixture: a stale block was not reported")
    return problems


def write(root: Path = ROOT) -> tuple[list[str], list[str]]:
    """Render every block in place; (rewritten block names, problems)."""
    facts, problems = load_facts(root)
    if facts is None:
        return [], problems
    path = root / README
    text = path.read_text(encoding="utf-8")
    updated, problems, stale = _fresh(text, facts)
    if not problems and updated != text:
        path.write_text(updated, encoding="utf-8", newline="\n")
    return stale, problems


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--write", action="store_true", help="render every block in place")
    args = parser.parse_args(argv)
    if args.write:
        stale, problems = write()
        print("\n".join(problems or [f"rendered facts:{name}" for name in stale]
                        or ["every block is current"]))
        return 1 if problems else 0
    problems, notes = repository_problems()
    print("\n".join([f"note: {n}" for n in notes] + (problems or ["every block is current"])))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
