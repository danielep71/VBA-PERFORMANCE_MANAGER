"""
release_provenance.py - generate the provenance block for a GitHub Release.

The shipped source files live in git, so their integrity is already established
by the commit they came from. The demo workbook does not: it is a binary
uploaded to a Release, and a downloader otherwise has no way to check that what
they received is what was built.

This emits a markdown block recording:

  * the commit the release was cut from
  * a SHA-256 for every shipped file, including the Release asset
  * the environment the regression suite was certified on

There is one strict path. Every release-critical input is required and is
validated before any publishable markdown or JSON is produced. Incomplete
output is not supported: a manifest is an integrity claim, and a partial claim
can still be uploaded by accident.

Run it from a clean checkout whose HEAD is exactly the release tag:

    python tools/release_provenance.py --version 1.4.0 --tag v1.4.0 \\
        --asset "PERFORMANCE MANAGER.xlsm" \\
        --excel "Microsoft 365 MSO, Version 2607, Build 16.0.20228.20188" \\
        --bitness 64-bit --cases 80 --assertions 643 --failures 0 \\
        --out release-manifest.json

Exit codes:

  0  a validated release manifest was produced
  1  the command parsed but violates the release contract
  2  a command-line syntax error, reported by the argument parser
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

# Recorded in the manifest so a reader knows which logic produced it.
TOOL_VERSION = "2.0.0"

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



def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def git(*args: str) -> str | None:
    """Run a git command, returning None on any failure.

    The encoding is pinned deliberately. subprocess decodes with the locale
    codepage by default, which on Windows is cp1252, so UTF-8 output from git
    came back mangled and every comparison against a file read as UTF-8 failed
    on the first non-ASCII character.
    """
    try:
        out = subprocess.run(
            ["git", *args], cwd=ROOT, capture_output=True, text=True, check=True,
            encoding="utf-8", errors="replace"
        )
        return out.stdout.strip()
    except Exception:
        return None


def git_raw(*args: str) -> str | None:
    """Run a git command and return its output UNSTRIPPED.

    git() strips, which is right for a commit SHA and wrong for file content:
    it removes the trailing newlines a source file legitimately ends with, so
    every file compared against its blob appeared to differ by whitespace that
    was never there.
    """
    try:
        out = subprocess.run(
            ["git", *args], cwd=ROOT, capture_output=True, text=True, check=True,
            encoding="utf-8", errors="replace",
        )
        return out.stdout
    except Exception:
        return None


def git_error(*args: str) -> str:
    """Return git's own message for a failing command, for reporting."""
    try:
        out = subprocess.run(
            ["git", *args], cwd=ROOT, capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
        return (out.stderr or out.stdout).strip()
    except FileNotFoundError:
        return "git executable not found on PATH"
    except Exception as e:  # pragma: no cover - defensive
        return str(e)


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
            "  If it is published, fetch it: GitHub Desktop -> Fetch origin,",
            "  or git fetch --tags from a shell that can see git.",
        ]

    problems: list[str] = []
    for rel in paths:
        p = ROOT / rel
        blob = git_raw("show", f"{tag}:{rel}")
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


GIT_MISSING_HINT = (
    "Common causes:\n"
    "  - git is not on PATH. GitHub Desktop ships its own copy: open a shell\n"
    "    that can see it with Repository -> Open in Command Prompt.\n"
    "  - Windows refuses the folder as dubious ownership, which happens often\n"
    "    under OneDrive. Run the git config command git suggests above."
)


def git_usable() -> bool:
    """Whether git can read THIS repository.

    Checking `git --version` is not enough. It succeeds whenever the executable
    exists, including when git then refuses to operate on the repository - most
    commonly Windows' dubious-ownership guard, which trips on folders under
    OneDrive and fails every repository command.

    That produced a manifest with an unresolved commit and every tag reported
    as missing, pointing the reader at the tags rather than at the real cause.
    """
    return git("rev-parse", "--git-dir") is not None


def rows(paths: list[str]) -> list[str]:
    out = []
    for rel in paths:
        p = ROOT / rel
        if not p.exists():
            out.append(f"| `{rel}` | *not present* |")
            continue
        out.append(f"| `{rel}` | `{sha256(p)}` |")
    return out


VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")
TAG_RE = re.compile(r"^v\d+\.\d+\.\d+$")

EXIT_OK = 0
EXIT_CONTRACT = 1
# Exit 2 is argparse's, for genuine command-line syntax errors. It is never
# returned from this module: a parsed command that breaks the release contract
# is a contract failure, not a usage error, and the two must stay separable.


def resolve_tag_commit(tag: str) -> str | None:
    """The commit a tag points at, following an annotated tag to its target."""
    return git("rev-parse", "-q", "--verify", f"refs/tags/{tag}^{{commit}}")


def validate(args: argparse.Namespace) -> list[str]:
    """Every release-contract violation in one pass.

    Collected rather than raised one at a time, so a single run tells the
    releaser everything that is wrong. A partially diagnosed release is how a
    manifest ends up regenerated four times in a row during a publish.
    """
    problems: list[str] = []

    # --- presence -----------------------------------------------------------
    required = {
        "--version": args.version,
        "--tag": args.tag,
        "--asset": args.asset,
        "--excel": args.excel,
        "--bitness": args.bitness,
        "--cases": args.cases,
        "--assertions": args.assertions,
        "--failures": args.failures,
    }
    for flag, value in required.items():
        if value is None:
            problems.append(f"{flag} is required in release mode")

    # --- numeric domains ----------------------------------------------------
    if args.cases is not None and args.cases <= 0:
        problems.append(f"--cases must be a positive integer, got {args.cases}")
    if args.assertions is not None and args.assertions <= 0:
        problems.append(f"--assertions must be a positive integer, got {args.assertions}")
    if args.failures is not None and args.failures != 0:
        problems.append(
            f"--failures must be exactly 0 to publish, got {args.failures}. "
            "A suite with failures has no publishable manifest."
        )

    # --- release identity ---------------------------------------------------
    if args.version is not None and not VERSION_RE.match(args.version):
        problems.append(f"--version {args.version!r} is not a MAJOR.MINOR.PATCH version")
    if args.tag is not None and not TAG_RE.match(args.tag):
        problems.append(f"--tag {args.tag!r} is not a vMAJOR.MINOR.PATCH tag")
    if (args.version is not None and args.tag is not None
            and VERSION_RE.match(args.version) and TAG_RE.match(args.tag)
            and args.tag != f"v{args.version}"):
        problems.append(
            f"--version {args.version} and --tag {args.tag} disagree; "
            f"expected --tag v{args.version}"
        )

    # --- asset --------------------------------------------------------------
    if args.asset is not None:
        p = Path(args.asset)
        if not p.is_absolute():
            p = ROOT / p
        if not p.exists():
            problems.append(f"--asset not found at {args.asset}")
        elif not p.is_file():
            problems.append(f"--asset is not a regular file: {args.asset}")

    # --- repository state ---------------------------------------------------
    # Only tracked modifications count. Untracked files cannot change what a
    # hashed source contains, and the release asset and the manifest itself
    # legitimately sit untracked in the working tree during a publish - so
    # treating them as dirt would reject every real release invocation.
    dirty = git("status", "--porcelain", "--untracked-files=no")
    if dirty:
        paths = [line[3:] for line in dirty.splitlines()][:20]
        problems.append("tracked files are modified; hashing them would not match the commit claimed")
        problems.extend(f"  modified: {path}" for path in paths)

    head = git("rev-parse", "HEAD")
    if head is None:
        problems.append("HEAD could not be resolved")

    if args.tag is not None:
        target = resolve_tag_commit(args.tag)
        if target is None:
            problems.append(f"tag {args.tag} does not exist in this clone")
            problems.append("  If the release is still a draft, GitHub creates the tag on publish.")
            problems.append("  If it is published, fetch it: GitHub Desktop -> Fetch origin.")
        elif head is not None and head != target:
            problems.append(
                f"HEAD {head} is not the target of {args.tag} ({target}). "
                "Check out the tag before generating provenance."
            )
        else:
            problems.extend(verify_against_tag(args.tag, REQUIRED + OPTIONAL))

    # --- sources ------------------------------------------------------------
    for rel in REQUIRED:
        if not (ROOT / rel).exists():
            problems.append(f"required source missing on disk: {rel}")

    return problems


def build_markdown(args: argparse.Namespace, commit: str) -> str:
    out: list[str] = []
    out.append("## 🔐 Provenance")
    out.append("")
    out.append("| | |")
    out.append("|---|---|")
    out.append(f"| **Version** | v{args.version} |")
    out.append(f"| **Commit** | `{commit}` |")
    out.append(f"| **Tag** | `{args.tag}` |")
    out.append(f"| **Built** | {date.today().isoformat()} |")
    out.append(f"| **Manifest tool** | `release_provenance.py {TOOL_VERSION}` |")
    out.append("")
    out.append("### Certification")
    out.append("")
    out.append("| | |")
    out.append("|---|---|")
    out.append(f"| **Regression suite** | {args.cases} cases · {args.assertions} assertions · "
               f"**{args.failures} failures** |")
    out.append(f"| **Excel** | `{args.excel}` |")
    out.append(f"| **Bitness** | **{args.bitness}** |")
    out.append("")
    if args.bitness == "64-bit":
        out.append("> [!NOTE]")
        out.append("> On 64-bit Office, backend 2 compiles to `GetTickCount64`, so")
        out.append("> `RolloverSeconds` is certified on its `Win64` branch only. The 32-bit")
        out.append("> wrap-correction branch is compiled out and was not exercised.")
        out.append("")
    else:
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

    asset_path = Path(args.asset)
    if not asset_path.is_absolute():
        asset_path = ROOT / asset_path
    out.append("")
    out.append("**Release assets**")
    out.append("")
    out.append("| Asset | SHA-256 |")
    out.append("|---|---|")
    out.append(f"| `{asset_path.name}` | `{sha256(asset_path)}` |")

    out.append("")
    out.append("### Source integrity")
    out.append("")
    out.append(f"`HEAD` is the target of `{args.tag}`, no tracked file is modified, and every")
    out.append(f"hashed source file matches its blob at `{args.tag}`.")
    out.append("")
    out.append("### What this establishes")
    out.append("")
    out.append("| Claim | Established by |")
    out.append("|---|---|")
    out.append("| The published source files are the tagged ones | comparison against the tag |")
    out.append("| The manifest describes the tagged commit | `HEAD` equals the tag target |")
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
    return "\n".join(out)


def build_manifest(args: argparse.Namespace, commit: str, tag_target: str) -> dict:
    manifest = {
        "tool": "release_provenance.py",
        "tool_version": TOOL_VERSION,
        "version": args.version,
        "tag": args.tag,
        "commit": commit,
        "built": date.today().isoformat(),
        "tag_verified": args.tag,
        "tag_problems": [],
        "certification": {
            "cases": args.cases,
            "assertions": args.assertions,
            "failures": args.failures,
            "excel": args.excel,
            "bitness": args.bitness,
        },
        "sha256": {rel: sha256(ROOT / rel) for rel in REQUIRED + OPTIONAL
                   if (ROOT / rel).exists()},
        "scope": {
            "source_files_match_tag": True,
            "head_equals_tag_target": commit == tag_target,
            "tracked_files_unmodified": True,
            "workbook_built_from_source": False,
            "workbook_build_is_manual": True,
        },
    }
    asset_path = Path(args.asset)
    if not asset_path.is_absolute():
        asset_path = ROOT / asset_path
    manifest["sha256"][asset_path.name] = sha256(asset_path)
    return manifest


def write_atomically(path: Path, text: str) -> None:
    """Write through a temporary file in the destination directory, then replace.

    A half-written manifest is worse than none: it looks like evidence. The
    temporary file shares a directory with the destination so os.replace is a
    rename within one filesystem, which is atomic.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".manifest-", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def main() -> int:
    # The provenance block contains a padlock and em-dashes. A Windows console
    # takes those directly, but a redirected or piped stdout falls back to the
    # locale codepage - cp1252 here - and the whole block dies on the first
    # non-Latin-1 character. Release output is routinely redirected to a file,
    # so pin the encoding rather than depending on where stdout happens to go.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    ap = argparse.ArgumentParser(description=__doc__)
    # Nothing is argparse-required. Missing release inputs are a contract
    # failure diagnosed together with every other one and reported as exit 1;
    # argparse's exit 2 is reserved for genuine syntax errors.
    ap.add_argument("--version", help="Release version, e.g. 1.4.0")
    ap.add_argument("--tag", help="Release tag, e.g. v1.4.0; HEAD must be its target")
    ap.add_argument("--asset", help="Path to the Release asset to hash")
    ap.add_argument("--excel", help="Excel version and build the suite was certified on")
    ap.add_argument("--bitness", choices=["32-bit", "64-bit"], help="Office bitness")
    ap.add_argument("--cases", type=int, help="Regression cases run")
    ap.add_argument("--assertions", type=int, help="Assertions executed")
    ap.add_argument("--failures", type=int, help="Failures; must be supplied and equal 0")
    ap.add_argument("--out", metavar="PATH", help="Also write the manifest as JSON")
    args = ap.parse_args()

    if not git_usable():
        print("git cannot read this repository, so nothing could be verified.", file=sys.stderr)
        print("", file=sys.stderr)
        print("git said:", file=sys.stderr)
        for line in git_error("rev-parse", "--git-dir").splitlines():
            print(f"  {line}", file=sys.stderr)
        print("", file=sys.stderr)
        print(GIT_MISSING_HINT, file=sys.stderr)
        return EXIT_CONTRACT

    problems = validate(args)
    if problems:
        print(f"{len([p for p in problems if not p.startswith('  ')])} release-contract "
              f"problem(s); no provenance was generated.", file=sys.stderr)
        for p in problems:
            print(f"  {p}" if not p.startswith("  ") else f"  {p}", file=sys.stderr)
        return EXIT_CONTRACT

    # Past this point every gate has passed, so output is safe to produce.
    commit = git("rev-parse", "HEAD")
    tag_target = resolve_tag_commit(args.tag)

    print(build_markdown(args, commit))

    if args.out:
        write_atomically(Path(args.out),
                         json.dumps(build_manifest(args, commit, tag_target), indent=2) + "\n")
        print(f"\nmanifest written to {args.out}", file=sys.stderr)

    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
