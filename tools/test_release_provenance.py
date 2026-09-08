#!/usr/bin/env python3
"""Strict-mode fixture matrix for release_provenance.py."""

from __future__ import annotations

import argparse
import importlib.util
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

TOOL = Path(__file__).resolve().parent / "release_provenance.py"
ROOT = TOOL.parent.parent
PROVENANCE_MARKER = "## 🔐 Provenance"

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
    subprocess.run(
        ["git", *args], cwd=repo, check=True, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )


def make_repo(
    tmp: Path,
    tag: str = "v1.4.0",
    *,
    version: str = "1.4.0",
    annotated: bool = True,
    include_version: bool = True,
) -> Path:
    repo = tmp / "repo"
    (repo / "tools").mkdir(parents=True)
    shutil.copy2(TOOL, repo / "tools" / TOOL.name)

    for rel in SOURCES:
        path = repo / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(f"' fixture content for {rel}\n", encoding="utf-8")
    if include_version:
        (repo / "VERSION").write_text(version + "\n", encoding="utf-8")

    _git(repo, "init", "-q", "-b", "main")
    _git(repo, "config", "user.email", "fixture@example.invalid")
    _git(repo, "config", "user.name", "fixture")
    _git(repo, "add", "-A")
    _git(repo, "commit", "-q", "-m", "fixture baseline")
    if tag:
        if annotated:
            _git(repo, "tag", "-a", tag, "-m", f"fixture {tag}")
        else:
            _git(repo, "tag", tag)

    (repo / ASSET).write_bytes(b"fixture workbook bytes")
    return repo


def run(repo: Path, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(repo / "tools" / TOOL.name), *args], cwd=repo,
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )


def valid_args(**overrides) -> list[str]:
    merged = dict(VALID)
    for key, value in overrides.items():
        flag = "--" + key.replace("_", "-")
        if value is None:
            merged.pop(flag, None)
        else:
            merged[flag] = value
    result: list[str] = []
    for flag, value in merged.items():
        result += [flag, value]
    return result


class Matrix:
    def __init__(self) -> None:
        self.failures: list[str] = []
        self.run_count = 0

    def case(self, name: str, ok: bool, detail: str = "") -> None:
        self.run_count += 1
        if ok:
            print(f"ok    {name}")
            return
        self.failures.append(f"{name}: {detail}" if detail else name)
        print(f"FAIL  {name}")
        if detail:
            print(f"        {detail}")

    def rejects(
        self, name: str, repo: Path, args: list[str], *,
        out_name: str = "release-manifest.json", expected: str | None = None,
    ) -> None:
        out = repo / out_name
        sentinel = "PREVIOUS MANIFEST - MUST NOT BE TOUCHED\n"
        out.parent.mkdir(parents=True, exist_ok=True)
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
        if expected and expected not in proc.stderr:
            problems.append(f"stderr does not contain {expected!r}")
        self.case(name, not problems, "; ".join(problems))


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-v", "--verbose", action="store_true")
    ap.add_argument("--json", metavar="PATH")
    args = ap.parse_args()

    m = Matrix()
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
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

        m.rejects("asset file does not exist", repo, valid_args(asset="NOT THERE.xlsm"))
        (repo / "a-directory").mkdir()
        m.rejects("asset is not a regular file", repo, valid_args(asset="a-directory"))
        m.rejects("cases is zero", repo, valid_args(cases="0"))
        m.rejects("cases is negative", repo, valid_args(cases="-1"))
        m.rejects("assertions is zero", repo, valid_args(assertions="0"))
        m.rejects("failures is negative", repo, valid_args(failures="-1"))
        m.rejects("failures is non-zero", repo, valid_args(failures="1"))

        m.rejects("malformed version", repo, valid_args(version="1.4"))
        m.rejects("malformed tag", repo, valid_args(tag="release-1.4.0"))
        m.rejects("version and tag disagree", repo, valid_args(version="9.9.9", tag="v1.4.0"))
        m.rejects("version has a leading zero", repo, valid_args(version="01.4.0", tag="v01.4.0"))
        m.rejects("tag has a leading zero", repo, valid_args(tag="v01.4.0"))
        m.rejects("pre-release numeric identifier has a leading zero", repo,
                  valid_args(version="1.4.0-01", tag="v1.4.0-01"))
        m.rejects("empty pre-release identifier is invalid", repo,
                  valid_args(version="1.4.0-", tag="v1.4.0-"))

        untagged = make_repo(tmp / "b", tag="")
        m.rejects("tag does not exist", untagged, valid_args())
        lightweight = make_repo(tmp / "c", annotated=False)
        m.rejects("lightweight tag is rejected", lightweight, valid_args(), expected="annotated tag")

        ahead = make_repo(tmp / "d")
        (ahead / SOURCES[1]).write_text("' a later commit\n", encoding="utf-8")
        _git(ahead, "add", "-A")
        _git(ahead, "commit", "-q", "-m", "later than the tag")
        m.rejects("HEAD is not the tag target", ahead, valid_args())

        dirty = make_repo(tmp / "e")
        (dirty / SOURCES[1]).write_text("' uncommitted edit\n", encoding="utf-8")
        m.rejects("tracked file is modified", dirty, valid_args())

        missing_src = make_repo(tmp / "f", tag="")
        (missing_src / SOURCES[0]).unlink()
        _git(missing_src, "add", "-A")
        _git(missing_src, "commit", "-q", "-m", "drop a required source")
        _git(missing_src, "tag", "-a", "v1.4.0", "-m", "fixture")
        m.rejects("required source is missing", missing_src, valid_args())

        content = make_repo(tmp / "g")
        _git(content, "update-index", "--assume-unchanged", SOURCES[1])
        (content / SOURCES[1]).write_text("' drifted from the tag\n", encoding="utf-8")
        m.rejects("source content differs from the tag", content, valid_args())

        wrong_version = make_repo(tmp / "h", tag="v9.9.9", version="1.4.0")
        m.rejects("requested version disagrees with tagged VERSION", wrong_version,
                  valid_args(version="9.9.9", tag="v9.9.9"), expected="VERSION")
        missing_version = make_repo(tmp / "i", include_version=False)
        m.rejects("tagged VERSION is required", missing_version, valid_args(), expected="VERSION")

        alias_asset = make_repo(tmp / "j")
        original_asset = (alias_asset / ASSET).read_bytes()
        proc = run(alias_asset, *valid_args(), "--out", ASSET)
        m.case(
            "--out cannot alias the release asset",
            proc.returncode == 1 and PROVENANCE_MARKER not in proc.stdout
            and (alias_asset / ASSET).read_bytes() == original_asset
            and "aliases a protected release input" in proc.stderr,
            proc.stderr.strip()[:300],
        )

        alias_source = make_repo(tmp / "k")
        source_path = alias_source / SOURCES[1]
        original_source = source_path.read_bytes()
        proc = run(alias_source, *valid_args(), "--out", SOURCES[1])
        m.case(
            "--out cannot alias a protected source",
            proc.returncode == 1 and PROVENANCE_MARKER not in proc.stdout
            and source_path.read_bytes() == original_source
            and "aliases a protected release input" in proc.stderr,
            proc.stderr.strip()[:300],
        )

        write_failure = make_repo(tmp / "l")
        (write_failure / "manifest-target").mkdir()
        proc = run(write_failure, *valid_args(), "--out", "manifest-target")
        m.case(
            "manifest write failure emits no publishable Markdown",
            proc.returncode == 1 and PROVENANCE_MARKER not in proc.stdout
            and "manifest write failed" in proc.stderr,
            f"exit={proc.returncode}; stdout={proc.stdout[:80]!r}; stderr={proc.stderr[:200]!r}",
        )

        releasing = ROOT / "RELEASING.md"
        problems = []
        if not releasing.exists():
            problems.append("RELEASING.md not found")
        else:
            spec = importlib.util.spec_from_file_location("_rp", TOOL)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            found = re.search(
                r"```bash\n(python tools/release_provenance\.py --version.*?)\n```",
                releasing.read_text(encoding="utf-8"), re.S,
            )
            if not found:
                problems.append("RELEASING.md has no provenance invocation block")
            elif found.group(1) not in (mod.__doc__ or ""):
                problems.append("the RELEASING.md invocation is not verbatim in the tool docstring")
        m.case("documented invocation matches the tool docstring", not problems, "; ".join(problems))

        for name, extra in [
            ("unknown option is exit 2", ["--nonsense"]),
            ("non-integer count is exit 2", ["--cases", "eighty"]),
            ("invalid bitness choice is exit 2", ["--bitness", "128-bit"]),
        ]:
            proc = run(repo, *valid_args(), *extra)
            m.case(name, proc.returncode == 2, f"exit {proc.returncode}, expected 2")

        good = make_repo(tmp / "m")
        out_name = "release-manifest.json"
        proc = run(good, *valid_args(), "--out", out_name)
        problems = []
        if proc.returncode != 0:
            problems.append(f"exit {proc.returncode}: {proc.stderr.strip()[:300]}")
        if PROVENANCE_MARKER not in proc.stdout:
            problems.append("no provenance block was printed")
        for marker in ("TODO", "not found", "not checked", "*not present*"):
            if marker in proc.stdout:
                problems.append(f"success output contains {marker!r}")
        m.case("valid invocation succeeds and emits a complete block", not problems, "; ".join(problems))

        manifest_path = good / out_name
        problems = []
        if not manifest_path.exists():
            problems.append("no manifest was written")
        else:
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            head = subprocess.run(
                ["git", "rev-parse", "HEAD"], cwd=good, capture_output=True,
                text=True, encoding="utf-8",
            ).stdout.strip()
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
            for key in (
                "source_files_match_tag", "head_equals_tag_target", "tag_is_annotated",
                "version_matches_tagged_source", "tracked_files_unmodified",
            ):
                if scope.get(key) is not True:
                    problems.append(f"scope.{key} is {scope.get(key)!r}")
            if ASSET not in data.get("sha256", {}):
                problems.append("the asset digest is absent")
            if "TODO" in manifest_path.read_text(encoding="utf-8"):
                problems.append("the manifest contains TODO")
        m.case("successful manifest is complete and identity-consistent", not problems, "; ".join(problems))

        proc = run(good, *valid_args())
        m.case(
            "valid invocation without --out succeeds",
            proc.returncode == 0 and PROVENANCE_MARKER in proc.stdout,
            f"exit {proc.returncode}",
        )

        prerelease = make_repo(
            tmp / "n", tag="v1.4.1-rc.1", version="1.4.1-rc.1", annotated=True,
        )
        proc = run(
            prerelease,
            *valid_args(version="1.4.1-rc.1", tag="v1.4.1-rc.1"),
            "--out", "release-manifest.json",
        )
        m.case(
            "valid SemVer pre-release succeeds",
            proc.returncode == 0 and PROVENANCE_MARKER in proc.stdout
            and (prerelease / "release-manifest.json").exists(),
            f"exit={proc.returncode}; stderr={proc.stderr[:200]!r}",
        )

    print("-" * 60)
    if args.json:
        Path(args.json).write_text(
            json.dumps({
                "tool": "test_release_provenance",
                "cases_run": m.run_count,
                "failures": m.failures,
                "passed": not m.failures,
            }, indent=2) + "\n",
            encoding="utf-8",
        )

    if m.failures:
        print(f"{len(m.failures)} of {m.run_count} fixtures failed:")
        for failure in m.failures:
            print(f"  - {failure}")
        return 1

    print(f"all {m.run_count} fixtures passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
