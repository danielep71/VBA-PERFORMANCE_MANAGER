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
import subprocess
import sys
from datetime import date
from pathlib import Path

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
    args = ap.parse_args()

    commit = git("rev-parse", "HEAD")
    dirty = git("status", "--porcelain")

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

    missing = text.count(TODO)
    if missing:
        print(f"\n{missing} field(s) left as TODO — supply them before publishing.",
              file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
