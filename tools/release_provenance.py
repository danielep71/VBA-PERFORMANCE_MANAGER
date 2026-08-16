#!/usr/bin/env python3
"""
release_provenance.py — generate the provenance block for a GitHub Release.

The shipped source files live in git, so their integrity is already established
by the commit they came from. The demo workbook does not: it is a binary
uploaded to a Release, and a downloader currently has no way to check that what
they received is what was built.

This emits a markdown block recording:

  * the commit the release was cut from
  * a SHA-256 for every shipped file, including any Release asset
  * the environment the regression suite was certified on

Run it from the repository root:

    python3 tools/release_provenance.py --version 1.3.0 \\
        --asset "demo/PERFORMANCE MANAGER.xlsm" \\
        --excel "Microsoft 365 MSO, Version 2606, Build 16.0.20131.20152" \\
        --bitness 64-bit --cases 69 --assertions 447

Anything omitted is emitted as a TODO marker rather than silently left out, so an
incomplete block is visible in review rather than shipped.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from datetime import date
from pathlib import Path

# Recorded in the manifest so a reader knows which logic produced it.
TOOL_VERSION = "1.1.0"

ROOT = Path(__file__).resolve().parent.parent

# Files that constitute the component. The demo and test modules are listed
# separately because they are optional imports.
REQUIRED = [
    "src/modules/M_cPM_TIMEWASTERS.bas",
    "src/classes/cPerformanceManager.cls",
]
OPTIONAL = [
    "test/M_cPM_Test.bas",
    "demo/M_cPM_DEMO.bas",
    "demo/M_cPM_USAGE_EXAMPLES.bas",
    "demo/M_DEMO_BUILDER.bas",
]

TODO = "`TODO`"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def git(*args: str) -> str | None:
    try:
        out = subprocess.run(
            ["git", *args], cwd=ROOT, capture_output=True, text=True, check=True
        )
        return out.stdout.strip()
    except Exception:
        return None


def verify_against_tag(tag: str, paths: list[str]) -> list[str]:
    """Confirm each file on disk matches the blob recorded at a tag.

    The manifest hashes the working tree while claiming a commit. Those can
    disagree - a stray edit, a stale checkout, a file exported but not committed.
    Comparing against the tag turns the claim into a checked fact.

    This proves the published hashes are the tagged sources. It says nothing
    about the workbook, which no automated step produces.
    """
    # Distinguish a missing tag from missing files. git show fails identically
    # for both, and reporting "not present at <tag>" for every file when the tag
    # itself is unreachable sends the reader looking in entirely the wrong place.
    if git("rev-parse", "-q", "--verify", f"refs/tags/{tag}") is None:
        return [
            f"tag {tag} does not exist in this clone",
            "  If the release is still a draft, GitHub creates the tag on publish.",
            "  If it is published, run: git fetch --tags",
        ]

    problems: list[str] = []
    for rel in paths:
        p = ROOT / rel
        blob = git("show", f"{tag}:{rel}")
        if blob is None:
            problems.append(f"{rel}: not present at {tag}")
            continue
        if not p.exists():
            problems.append(f"{rel}: present at {tag} but missing on disk")
            continue
        on_disk = p.read_text(encoding="utf-8", errors="replace")
        if on_disk.replace("\r\n", "\n") != blob.replace("\r\n", "\n"):
            problems.append(f"{rel}: differs from {tag}")
    return problems


def rows(paths: list[str]) -> list[str]:
    out = []
    for rel in paths:
        p = ROOT / rel
        if not p.exists():
            out.append(f"| `{rel}` | *not present* |")
            continue
        out.append(f"| `{rel}` | `{sha256(p)}` |")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--version", required=True, help="Release version, e.g. 1.3.0")
    ap.add_argument("--asset", help="Path to a Release asset to hash, e.g. the demo workbook")
    ap.add_argument("--excel", help="Excel version and build the suite was certified on")
    ap.add_argument("--bitness", choices=["32-bit", "64-bit"], help="Office bitness")
    ap.add_argument("--cases", type=int, help="Regression cases run")
    ap.add_argument("--assertions", type=int, help="Assertions executed")
    ap.add_argument("--failures", type=int, default=0, help="Failures (default 0)")
    ap.add_argument("--tag", help="Verify every hashed source matches this tag's blob")
    ap.add_argument("--out", metavar="PATH", help="Also write the manifest as JSON")
    args = ap.parse_args()

    commit = git("rev-parse", "HEAD")
    dirty = git("status", "--porcelain")

    tag_problems: list[str] = []
    if args.tag:
        tag_problems = verify_against_tag(args.tag, REQUIRED + OPTIONAL)

    if dirty:
        print("WARNING: the working tree has uncommitted changes.", file=sys.stderr)
        print("         Hashes below will not match the commit they claim.", file=sys.stderr)
        print("", file=sys.stderr)

    excel = f"`{args.excel}`" if args.excel else TODO
    bitness = f"**{args.bitness}**" if args.bitness else TODO
    cases = str(args.cases) if args.cases is not None else TODO
    asserts = str(args.assertions) if args.assertions is not None else TODO

    out: list[str] = []
    out.append("## 🔐 Provenance")
    out.append("")
    out.append("| | |")
    out.append("|---|---|")
    out.append(f"| **Version** | v{args.version} |")
    out.append(f"| **Commit** | `{commit or TODO}` |")
    out.append(f"| **Built** | {date.today().isoformat()} |")
    out.append(f"| **Manifest tool** | `release_provenance.py {TOOL_VERSION}` |")
    out.append("")
    out.append("### Certification")
    out.append("")
    out.append("| | |")
    out.append("|---|---|")
    out.append(f"| **Regression suite** | {cases} cases · {asserts} assertions · "
               f"**{args.failures} failures** |")
    out.append(f"| **Excel** | {excel} |")
    out.append(f"| **Bitness** | {bitness} |")
    out.append("")
    if args.bitness == "64-bit":
        out.append("> [!NOTE]")
        out.append("> On 64-bit Office, backend 2 compiles to `GetTickCount64`, so")
        out.append("> `RolloverSeconds` is certified on its `Win64` branch only. The 32-bit")
        out.append("> wrap-correction branch is compiled out and was not exercised.")
        out.append("")
    elif args.bitness == "32-bit":
        out.append("> [!NOTE]")
        out.append("> On 32-bit Office, backend 2 compiles to `GetTickCount`, so the Win64")
        out.append("> branch of `RolloverSeconds` is compiled out and was not exercised.")
        out.append("")

    out.append("### SHA-256")
    out.append("")
    out.append("**Required files**")
    out.append("")
    out.append("| File | SHA-256 |")
    out.append("|---|---|")
    out.extend(rows(REQUIRED))
    out.append("")
    out.append("**Optional files**")
    out.append("")
    out.append("| File | SHA-256 |")
    out.append("|---|---|")
    out.extend(rows(OPTIONAL))

    if args.asset:
        p = Path(args.asset)
        if not p.is_absolute():
            p = ROOT / p
        out.append("")
        out.append("**Release assets**")
        out.append("")
        out.append("| Asset | SHA-256 |")
        out.append("|---|---|")
        if p.exists():
            out.append(f"| `{p.name}` | `{sha256(p)}` |")
        else:
            out.append(f"| `{p.name}` | *not found at {args.asset}* |")

    out.append("")
    if args.tag:
        out.append("### Source integrity")
        out.append("")
        if tag_problems:
            missing_tag = tag_problems and tag_problems[0].startswith("tag ")
            out.append("> [!CAUTION]")
            if missing_tag:
                out.append(f"> **`{args.tag}` could not be found, so nothing was verified.**")
                out.append(">")
                for p in tag_problems:
                    out.append(f"> {p.strip()}" if p.startswith("  ") else f"> - {p}")
            else:
                out.append(f"> The hashed sources do **not** match `{args.tag}`:")
                out.append(">")
                for p in tag_problems:
                    out.append(f"> - {p}")
        else:
            out.append(f"Every hashed source file matches its blob at `{args.tag}`.")
        out.append("")

    out.append("### What this establishes")
    out.append("")
    out.append("| Claim | Established by |")
    out.append("|---|---|")
    out.append("| The published source files are the tagged ones | "
               + ("comparison against the tag" if args.tag else "*not checked — rerun with `--tag`*") + " |")
    out.append("| A downloaded file is the one published here | its SHA-256 |")
    out.append("| The suite passed in the stated environment | the certification block, asserted by the releaser |")
    out.append("")
    out.append("> [!IMPORTANT]")
    out.append("> **The workbook digest does not prove the workbook was built from this")
    out.append("> source.** No automated step produces it: the modules are imported by hand")
    out.append("> and the file is saved, and the VBA editor reformats on import — stripping")
    out.append("> alignment, appending blank lines — so extracted source could never match")
    out.append("> the repository byte for byte even in principle.")
    out.append(">")
    out.append("> The digest establishes that a download is the file published here, which")
    out.append("> is a real and useful guarantee, and a different one.")
    out.append(">")
    out.append("> **The source files in this tag are authoritative.** The workbook is a")
    out.append("> convenience copy. Where they disagree, the source is right.")
    out.append("")
    out.append("<details>")
    out.append("<summary><strong>Verifying a download</strong></summary>")
    out.append("")
    out.append("PowerShell:")
    out.append("")
    out.append("```powershell")
    out.append('Get-FileHash -Algorithm SHA256 ".\\PERFORMANCE MANAGER.xlsm"')
    out.append("```")
    out.append("")
    out.append("The source files are text and are also in the repository at the commit")
    out.append("above, so `git show <commit>:<path>` gives the same bytes. The workbook is")
    out.append("only distributed as a Release asset, which is why its hash is published")
    out.append("here.")
    out.append("")
    out.append("</details>")

    text = "\n".join(out)
    print(text)

    if args.out:
        manifest = {
            "tool": "release_provenance.py",
            "tool_version": TOOL_VERSION,
            "version": args.version,
            "commit": commit,
            "built": date.today().isoformat(),
            "tag_verified": args.tag if (args.tag and not tag_problems) else None,
            "tag_problems": tag_problems,
            "certification": {
                "cases": args.cases,
                "assertions": args.assertions,
                "failures": args.failures,
                "excel": args.excel,
                "bitness": args.bitness,
            },
            "sha256": {
                rel: sha256(ROOT / rel)
                for rel in REQUIRED + OPTIONAL
                if (ROOT / rel).exists()
            },
            "scope": {
                "source_files_match_tag": bool(args.tag and not tag_problems),
                "workbook_built_from_source": False,
                "workbook_build_is_manual": True,
            },
        }
        if args.asset:
            ap_path = Path(args.asset)
            if not ap_path.is_absolute():
                ap_path = ROOT / ap_path
            if ap_path.exists():
                manifest["sha256"][ap_path.name] = sha256(ap_path)
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
        print(f"\nmanifest written to {out_path}", file=sys.stderr)

    if tag_problems:
        print(f"\n{len(tag_problems)} source file(s) do not match {args.tag}.", file=sys.stderr)
        return 1

    missing = text.count(TODO)
    if missing:
        print(f"\n{missing} field(s) left as TODO — supply them before publishing.",
              file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
