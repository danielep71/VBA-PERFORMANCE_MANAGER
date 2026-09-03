#!/usr/bin/env python3
"""
test_release_provenance.py - the strict-mode fixture matrix for the provenance tool.

Every case builds a throwaway git repository, copies the production script into
its tools/ directory, and invokes that copy as a subprocess. The copy is
necessary rather than fastidious: release_provenance.py derives ROOT from
__file__ and passes cwd=ROOT to every git call, so running the real script from
a different working directory would silently verify the real checkout while the
fixture believed it was testing a temporary one.

Run standalone for the closure matrix:

    python tools/test_release_provenance.py -v

vba_lint.py runs it as the `release provenance strict fixtures` check, so a
non-zero result here is an ordinary static-gate failure.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

TOOL = Path(__file__).resolve().parent / "release_provenance.py"
ROOT = TOOL.parent.parent

PROVENANCE_MARKER = "## 🔐 Provenance"

# Mirrors REQUIRED + OPTIONAL in the tool. Content is arbitrary; only identity
# and tag agreement matter here.
SOURCES = [
    "src/modules/M_cPM_TIMEWASTERS.bas",
    "src/classes/cPerformanceManager.cls",
    "test/M_cPM_Test.bas",
    "demo/M_cPM_DEMO.bas",
    "demo/M_cPM_USAGE_EXAMPLES.bas",
    "demo/M_DEMO_BUILDER.bas",
]

ASSET = "PERFORMANCE MANAGER.xlsm"

VALID = {
    "--version": "1.4.0",
    "--tag": "v1.4.0",
    "--asset": ASSET,
    "--excel": "Microsoft 365 MSO, Version 2607, Build 16.0.20228.20188",
    "--bitness": "64-bit",
    "--cases": "80",
    "--assertions": "643",
    "--failures": "0",
}


def _git(repo: Path, *args: str) -> None:
    subprocess.run(["git", *args], cwd=repo, check=True,
                   capture_output=True, text=True, encoding="utf-8", errors="replace")


def make_repo(tmp: Path, tag: str = "v1.4.0") -> Path:
    """A minimal repository at a tag, with the tool copied in and a clean tree."""
    repo = tmp / "repo"
    (repo / "tools").mkdir(parents=True)
    shutil.copy2(TOOL, repo / "tools" / TOOL.name)

    for rel in SOURCES:
        p = repo / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(f"' fixture content for {rel}\n", encoding="utf-8")

    _git(repo, "init", "-q", "-b", "main")
    _git(repo, "config", "user.email", "fixture@example.invalid")
    _git(repo, "config", "user.name", "fixture")
    _git(repo, "add", "-A")
    _git(repo, "commit", "-q", "-m", "fixture baseline")
    if tag:
        _git(repo, "tag", tag)

    # Written after the commit, so it stays untracked. That is how a real
    # release runs: the workbook is a Release asset and is not in the
    # repository, so a fixture that commits it would prove the wrong thing.
    (repo / ASSET).write_bytes(b"fixture workbook bytes")
    return repo


def run(repo: Path, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(repo / "tools" / TOOL.name), *args],
        cwd=repo, capture_output=True, text=True, encoding="utf-8", errors="replace",
    )


def valid_args(**overrides) -> list[str]:
    """The successful invocation, with named flags replaced or dropped.

    Pass None to drop a flag entirely, which is how the missing-input cases are
    expressed without hand-assembling argument lists.
    """
    merged = dict(VALID)
    for key, value in overrides.items():
        flag = "--" + key.replace("_", "-")
        if value is None:
            merged.pop(flag, None)
        else:
            merged[flag] = value
    args: list[str] = []
    for flag, value in merged.items():
        args += [flag, value]
    return args


class Matrix:
    def __init__(self) -> None:
        self.failures: list[str] = []
        self.run_count = 0

    def case(self, name: str, ok: bool, detail: str = "") -> None:
        self.run_count += 1
        if ok:
            print(f"ok    {name}")
        else:
            self.failures.append(f"{name}: {detail}" if detail else name)
            print(f"FAIL  {name}")
            if detail:
                print(f"        {detail}")

    def rejects(self, name: str, repo: Path, args: list[str], *, out_name: str = "release-manifest.json") -> None:
        """A release-contract violation: exit 1, no publishable block, no output file.

        The three assertions travel together deliberately. Exit 1 alone would
        pass even if the tool had already printed a manifest and written JSON
        before deciding to fail, which is the defect this issue exists for.
        """
        out = repo / out_name
        sentinel = "PREVIOUS MANIFEST - MUST NOT BE TOUCHED\n"
        out.write_text(sentinel, encoding="utf-8")
        before = out.stat().st_mtime_ns

        proc = run(repo, *args, "--out", out_name)

        problems = []
        if proc.returncode != 1:
            problems.append(f"exit {proc.returncode}, expected 1")
        if PROVENANCE_MARKER in proc.stdout:
            problems.append("a publishable provenance block was printed")
        if out.read_text(encoding="utf-8") != sentinel:
            problems.append("the existing --out file was overwritten")
        if out.stat().st_mtime_ns != before:
            problems.append("the existing --out file was touched")
        self.case(name, not problems, "; ".join(problems))


def main() -> int:
    # The fixtures print the same characters the tool does, and vba_lint.py
    # captures this output through a pipe, so pin the encoding here too.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    ap = argparse.ArgumentParser(description="Strict-mode fixtures for release_provenance.py")
    ap.add_argument("-v", "--verbose", action="store_true", help="Print each case as it runs")
    ap.add_argument("--json", metavar="PATH", help="Write a machine-readable result document")
    args = ap.parse_args()

    if not TOOL.exists():
        print(f"FAIL  tool not found: {TOOL}")
        return 1

    m = Matrix()

    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)

        # --- input completeness --------------------------------------------
        repo = make_repo(tmp / "a")
        m.rejects("missing --version", repo, valid_args(version=None))
        m.rejects("missing --asset", repo, valid_args(asset=None))
        m.rejects("missing --tag", repo, valid_args(tag=None))
        m.rejects("missing --excel", repo, valid_args(excel=None))
        m.rejects("missing --bitness", repo, valid_args(bitness=None))
        m.rejects("missing --cases", repo, valid_args(cases=None))
        m.rejects("missing --assertions", repo, valid_args(assertions=None))
        m.rejects("missing --failures", repo, valid_args(failures=None))
        m.rejects("blank --excel", repo, valid_args(excel="   "))

        # --- asset integrity -------------------------------------------------
        m.rejects("asset file does not exist", repo, valid_args(asset="NOT THERE.xlsm"))
        (repo / "a-directory").mkdir()
        m.rejects("asset is not a regular file", repo, valid_args(asset="a-directory"))

        # --- numeric domains -------------------------------------------------
        m.rejects("cases is zero", repo, valid_args(cases="0"))
        m.rejects("cases is negative", repo, valid_args(cases="-1"))
        m.rejects("assertions is zero", repo, valid_args(assertions="0"))
        m.rejects("failures is negative", repo, valid_args(failures="-1"))
        m.rejects("failures is non-zero", repo, valid_args(failures="1"))

        # --- release identity ------------------------------------------------
        m.rejects("malformed version", repo, valid_args(version="1.4"))
        m.rejects("malformed tag", repo, valid_args(tag="release-1.4.0"))
        m.rejects("version and tag disagree", repo, valid_args(version="9.9.9", tag="v1.4.0"))
        m.rejects("version has a leading zero", repo, valid_args(version="01.4.0", tag="v01.4.0"))
        m.rejects("tag has a leading zero", repo, valid_args(tag="v01.4.0"))

        # A tag that does not exist at all, distinguished from one that does.
        untagged = make_repo(tmp / "b", tag="")
        m.rejects("tag does not exist", untagged, valid_args())

        # --- repository state ------------------------------------------------
        ahead = make_repo(tmp / "c")
        (ahead / "src/classes/cPerformanceManager.cls").write_text(
            "' a later commit\n", encoding="utf-8")
        _git(ahead, "add", "-A")
        _git(ahead, "commit", "-q", "-m", "later than the tag")
        m.rejects("HEAD is not the tag target", ahead, valid_args())

        dirty = make_repo(tmp / "d")
        (dirty / "src/classes/cPerformanceManager.cls").write_text(
            "' uncommitted edit\n", encoding="utf-8")
        m.rejects("tracked file is modified", dirty, valid_args())

        missing_src = make_repo(tmp / "e", tag="")
        (missing_src / "src/modules/M_cPM_TIMEWASTERS.bas").unlink()
        _git(missing_src, "add", "-A")
        _git(missing_src, "commit", "-q", "-m", "drop a required source")
        _git(missing_src, "tag", "v1.4.0")
        m.rejects("required source is missing", missing_src, valid_args())

        # --- tag/source content mismatch --------------------------------------
        # Reachable through the CLI only when git is told to stop noticing a
        # modification. assume-unchanged does exactly that, and a stale
        # assume-unchanged bit is a real way to end up hashing content that is
        # not what the tag holds while `git status` reports nothing.
        content = make_repo(tmp / "f")
        _git(content, "update-index", "--assume-unchanged",
             "src/classes/cPerformanceManager.cls")
        (content / "src/classes/cPerformanceManager.cls").write_text(
            "' drifted from the tag\n", encoding="utf-8")
        m.rejects("source content differs from the tag", content, valid_args())

        # --- documentation drift ---------------------------------------------
        # The tool's help and RELEASING.md must show the same command. They
        # drifted apart once already, which is how a documented invocation
        # stops being one the tool would accept.
        releasing = ROOT / "RELEASING.md"
        problems = []
        if not releasing.exists():
            problems.append("RELEASING.md not found")
        else:
            import importlib.util
            spec = importlib.util.spec_from_file_location("_rp", TOOL)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            found = re.search(r"```bash\n(python tools/release_provenance\.py --version.*?)\n```",
                              releasing.read_text(encoding="utf-8"), re.S)
            if not found:
                problems.append("RELEASING.md has no provenance invocation block")
            elif found.group(1) not in (mod.__doc__ or ""):
                problems.append("the RELEASING.md invocation is not verbatim in the tool docstring")
        m.case("documented invocation matches the tool docstring", not problems,
               "; ".join(problems))

        # --- parser-only syntax errors, which are exit 2 ----------------------
        for name, extra in [
            ("unknown option is exit 2", ["--nonsense"]),
            ("non-integer count is exit 2", ["--cases", "eighty"]),
            ("invalid bitness choice is exit 2", ["--bitness", "128-bit"]),
        ]:
            proc = run(repo, *valid_args(), *extra)
            m.case(name, proc.returncode == 2, f"exit {proc.returncode}, expected 2")

        # --- the positive case ------------------------------------------------
        good = make_repo(tmp / "g")
        out_name = "release-manifest.json"
        proc = run(good, *valid_args(), "--out", out_name)
        problems = []
        if proc.returncode != 0:
            problems.append(f"exit {proc.returncode}, expected 0: {proc.stderr.strip()[:300]}")
        if PROVENANCE_MARKER not in proc.stdout:
            problems.append("no provenance block was printed")
        for marker in ("TODO", "not found", "not checked", "*not present*"):
            if marker in proc.stdout:
                problems.append(f"success output contains an incomplete marker: {marker!r}")
        m.case("valid invocation succeeds and emits a complete block", not problems,
               "; ".join(problems))

        manifest_path = good / out_name
        problems = []
        if not manifest_path.exists():
            problems.append("no manifest was written")
        else:
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            head = subprocess.run(["git", "rev-parse", "HEAD"], cwd=good, capture_output=True,
                                  text=True, encoding="utf-8").stdout.strip()
            if data.get("version") != "1.4.0":
                problems.append(f"version is {data.get('version')!r}")
            if data.get("tag") != "v1.4.0":
                problems.append(f"tag is {data.get('tag')!r}")
            if data.get("commit") != head:
                problems.append("commit does not match HEAD")
            if data.get("tag_problems") != []:
                problems.append(f"tag_problems is {data.get('tag_problems')!r}")
            if data["certification"]["failures"] != 0:
                problems.append("certification.failures is not 0")
            scope = data.get("scope", {})
            for key in ("source_files_match_tag", "head_equals_tag_target", "tracked_files_unmodified"):
                if scope.get(key) is not True:
                    problems.append(f"scope.{key} is {scope.get(key)!r}")
            if ASSET not in data.get("sha256", {}):
                problems.append("the asset digest is absent")
            if "TODO" in manifest_path.read_text(encoding="utf-8"):
                problems.append("the manifest contains a TODO sentinel")
        m.case("successful manifest is complete and identity-consistent", not problems,
               "; ".join(problems))

        # Success must not be achievable without --out either, since a valid
        # invocation may legitimately emit markdown alone.
        proc = run(good, *valid_args())
        m.case("valid invocation without --out succeeds",
               proc.returncode == 0 and PROVENANCE_MARKER in proc.stdout,
               f"exit {proc.returncode}")

    print("-" * 60)
    if args.json:
        Path(args.json).write_text(json.dumps({
            "tool": "test_release_provenance",
            "cases_run": m.run_count,
            "failures": m.failures,
            "passed": not m.failures,
        }, indent=2) + "\n", encoding="utf-8")

    if m.failures:
        print(f"{len(m.failures)} of {m.run_count} fixtures failed:")
        for f in m.failures:
            print(f"  - {f}")
        return 1

    print(f"all {m.run_count} fixtures passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
