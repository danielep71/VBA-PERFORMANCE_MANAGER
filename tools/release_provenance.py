"""
release_provenance.py - generate the provenance block for a GitHub Release.

The shipped source files live in git, so their integrity is already established
by the commit they came from. The demo workbook does not: it is a binary
uploaded to a Release, and a downloader otherwise has no way to check that what
they received is what was built.

This emits a markdown block recording:

  * the commit the release was cut from
  * a SHA-256 for every shipped file, including the Release package and each
    file inside it
  * the environment the regression suite was certified on

The Release asset is one ZIP package, VBA-Performance-Manager-vX.Y.Z.zip,
holding exactly the workbook, LICENSE and NOTICE at its root. The package is
verified before anything is emitted: its name must match the version, it
must be a sound ZIP with exactly those members, and LICENSE and NOTICE must be
byte-identical to their blobs at the release tag.

There is one strict path. Every release-critical input is required and is
validated before any publishable markdown or JSON is produced. Incomplete
output is not supported: a manifest is an integrity claim, and a partial claim
can still be uploaded by accident.

Run it from a clean checkout whose HEAD is exactly the annotated release tag.
This is the same command RELEASING.md documents, and a fixture keeps the two
identical:

python tools/release_provenance.py --version X.Y.Z --tag vX.Y.Z \\
    --asset VBA-Performance-Manager-vX.Y.Z.zip \\
    --excel "Microsoft 365 MSO, Version 2607, Build 16.0.20228.20188" \\
    --bitness 64-bit --cases 80 --assertions 643 --failures 0 \\
    --out release-manifest.json

Valid documented SemVer pre-releases use the same form, for example
`--version 1.4.1-rc.1 --tag v1.4.1-rc.1`.

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
import zipfile
from datetime import date
from pathlib import Path

TOOL_VERSION = "3.0.0"
ROOT = Path(__file__).resolve().parent.parent

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
VERSION_FILE = "VERSION"
LICENSE_FILES = ("LICENSE", "NOTICE")
WORKBOOK_SUFFIX = ".xlsm"


def package_name(version: str) -> str:
    """The one accepted Release asset name for a version."""
    return f"VBA-Performance-Manager-v{version}.zip"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def git(*args: str) -> str | None:
    """Run a git command, returning stripped stdout or None on failure."""
    try:
        out = subprocess.run(
            ["git", *args], cwd=ROOT, capture_output=True, text=True, check=True,
            encoding="utf-8", errors="replace",
        )
        return out.stdout.strip()
    except Exception:
        return None


def git_raw(*args: str) -> str | None:
    """Run a git command and return its output unstripped."""
    try:
        out = subprocess.run(
            ["git", *args], cwd=ROOT, capture_output=True, text=True, check=True,
            encoding="utf-8", errors="replace",
        )
        return out.stdout
    except Exception:
        return None


def git_bytes(*args: str) -> bytes | None:
    """Run a git command and return its raw stdout bytes, or None on failure."""
    try:
        return subprocess.run(
            ["git", *args], cwd=ROOT, capture_output=True, check=True,
        ).stdout
    except Exception:
        return None


def git_input(stdin: str, *args: str) -> str | None:
    """Run a git command with text on stdin, returning stdout or None on failure."""
    try:
        out = subprocess.run(
            ["git", *args], cwd=ROOT, input=stdin, capture_output=True, text=True,
            check=True, encoding="utf-8", errors="replace",
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
    except Exception as exc:  # pragma: no cover - defensive
        return str(exc)


def git_usable() -> bool:
    """Return whether git can identify the repository containing this tool."""
    return git("rev-parse", "--git-dir") is not None


def verify_against_tag(tag: str, paths: list[str]) -> list[str]:
    """Confirm each file on disk matches the blob recorded at a tag."""
    if git("rev-parse", "-q", "--verify", f"refs/tags/{tag}") is None:
        return [
            f"tag {tag} does not exist in this clone",
            "  If the release is still a draft, create the annotated tag before provenance.",
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


def tracked_content_problems() -> list[str]:
    """Verify every file tracked at HEAD against the working tree.

    git status trusts the index, so an assume-unchanged or skip-worktree entry
    hides an edited file from it. Those flags are rejected outright, and each
    tracked file is then re-hashed from disk through git's own attribute
    filters (including end-of-line conversion) and compared with its blob at
    HEAD, without consulting the index. Anything that cannot be verified fails.
    """
    problems: list[str] = []

    entries = git_raw("ls-files", "-v", "-z")
    if entries is None:
        return ["git ls-files -v failed, so index flags could not be checked"]
    flagged = []
    for entry in entries.split("\0"):
        if len(entry) < 3:
            continue
        tag, rel = entry[0], entry[2:]
        if tag in "Ss":
            flagged.append((rel, "skip-worktree"))
        if tag.islower():
            flagged.append((rel, "assume-unchanged"))
    if flagged:
        problems.append("index flags hide tracked files from git status; clear them before provenance")
        problems.extend(f"  {rel}: {flag} (git update-index --no-{flag} -- {rel})"
                        for rel, flag in flagged[:20])

    tree = git_raw("ls-tree", "-r", "-z", "--full-tree", "HEAD")
    if tree is None:
        return problems + ["git ls-tree failed, so tracked content could not be verified"]
    expected: dict[str, str] = {}
    unverifiable: list[str] = []
    for entry in tree.split("\0"):
        if not entry:
            continue
        meta, rel = entry.split("\t", 1)
        mode, kind, oid = meta.split()
        if kind != "blob" or mode not in ("100644", "100755") or "\n" in rel:
            unverifiable.append(f"  {rel}: tracked {kind} {mode} cannot be verified")
            continue
        expected[rel] = oid

    missing = [rel for rel in expected if not (ROOT / rel).is_file()]
    present = [rel for rel in expected if rel not in missing]
    actual: list[str] = []
    if present:
        hashed = git_input("\n".join(present) + "\n", "hash-object", "--stdin-paths")
        actual = hashed.split() if hashed is not None else []
        if len(actual) != len(present):
            return problems + ["git hash-object failed, so tracked content could not be verified"]
    differing = [rel for rel, oid in zip(present, actual) if oid != expected[rel]]

    if unverifiable:
        problems.append("tracked entries the content check does not support")
        problems.extend(unverifiable[:20])
    if missing or differing:
        problems.append("tracked files differ from HEAD on disk; hashing them would not match the commit claimed")
        problems.extend(f"  missing: {rel}" for rel in missing[:20])
        problems.extend(f"  differs: {rel}" for rel in differing[:20])
    return problems


def package_problems(asset: Path, version: str | None, tag: str | None) -> list[str]:
    """Verify the Release package before it is described as published.

    The package must be named for the release version, be a sound ZIP, and
    hold exactly LICENSE, NOTICE and one workbook at its root. LICENSE and
    NOTICE must be byte-identical to their blobs at the tag. Anything that
    cannot be read or compared fails closed. tag is None when the tag itself
    already failed validation, so the content comparison is skipped.
    """
    problems: list[str] = []
    if version is not None and asset.name != package_name(version):
        problems.append(
            f"--asset must be the Release package {package_name(version)!r}, "
            f"got {asset.name!r}"
        )
    try:
        with zipfile.ZipFile(asset) as package:
            bad = package.testzip()
            if bad is not None:
                return problems + [f"package member {bad!r} is corrupt (CRC mismatch)"]
            infos = package.infolist()
            names = [info.filename for info in infos]
            members = {info.filename: package.read(info) for info in infos}
    except (zipfile.BadZipFile, OSError, RuntimeError, ValueError) as exc:
        return problems + [f"--asset is not a readable ZIP package: {exc}"]

    for name in sorted({n for n in names if names.count(n) > 1}):
        problems.append(f"package lists {name!r} more than once")
    unsafe = [n for n in names if n.endswith("/") or "/" in n or "\\" in n or ":" in n]
    for name in unsafe:
        problems.append(f"package entry {name!r} is not a plain file at the package root")
    root = [n for n in dict.fromkeys(names) if n not in unsafe]
    workbooks = [n for n in root if n.lower().endswith(WORKBOOK_SUFFIX)]
    for required in LICENSE_FILES:
        if required not in root:
            problems.append(f"package is missing {required}")
    if not workbooks:
        problems.append(f"package has no {WORKBOOK_SUFFIX} workbook")
    elif len(workbooks) > 1:
        problems.append(f"package has {len(workbooks)} workbooks; exactly one is allowed")
    elif not members[workbooks[0]]:
        problems.append(f"package workbook {workbooks[0]!r} is empty")
    for name in root:
        if name not in LICENSE_FILES and name not in workbooks:
            problems.append(f"package contains unexpected file {name!r}")

    if tag is not None:
        for required in LICENSE_FILES:
            tagged = git_bytes("show", f"{tag}:{required}")
            if tagged is None:
                problems.append(f"{required} is missing from tagged source {tag}")
            elif required in members and members[required] != tagged:
                problems.append(f"package {required} differs from {required} at {tag}")
    return problems


def package_members(asset: Path) -> dict[str, str]:
    """SHA-256 of each file inside an already verified package."""
    with zipfile.ZipFile(asset) as package:
        return {info.filename: hashlib.sha256(package.read(info)).hexdigest()
                for info in package.infolist()}


GIT_MISSING_HINT = (
    "Common causes:\n"
    "  - git is not on PATH. GitHub Desktop ships its own copy: open a shell\n"
    "    that can see it with Repository -> Open in Command Prompt.\n"
    "  - Windows refuses the folder as dubious ownership, which happens often\n"
    "    under OneDrive. Run the git config command git suggests above."
)


def rows(paths: list[str]) -> list[str]:
    out = []
    for rel in paths:
        p = ROOT / rel
        if not p.exists():
            out.append(f"| `{rel}` | *not present* |")
            continue
        out.append(f"| `{rel}` | `{sha256(p)}` |")
    return out


# SemVer core numbers cannot contain leading zeroes. Numeric pre-release
# identifiers follow the same rule; alphanumeric identifiers may contain ASCII
# letters and hyphens. Build metadata is deliberately outside the repository's
# documented release identity contract.
_NUM = r"(?:0|[1-9]\d*)"
_PRE_ID = r"(?:0|[1-9]\d*|[0-9A-Za-z-]*[A-Za-z-][0-9A-Za-z-]*)"
_PRE = rf"(?:-{_PRE_ID}(?:\.{_PRE_ID})*)?"
VERSION_RE = re.compile(rf"^{_NUM}\.{_NUM}\.{_NUM}{_PRE}$")
TAG_RE = re.compile(rf"^v{_NUM}\.{_NUM}\.{_NUM}{_PRE}$")

EXIT_OK = 0
EXIT_CONTRACT = 1


def rooted_path(value: str) -> Path:
    """Resolve a user path consistently relative to the repository root."""
    path = Path(value)
    if not path.is_absolute():
        path = ROOT / path
    return path.resolve(strict=False)


def tag_object_type(tag: str) -> str | None:
    """Return the object type stored directly at refs/tags/<tag>."""
    return git("cat-file", "-t", f"refs/tags/{tag}")


def resolve_tag_commit(tag: str) -> str | None:
    """The commit an annotated tag ultimately points at."""
    return git("rev-parse", "-q", "--verify", f"refs/tags/{tag}^{{commit}}")


def tagged_version(tag: str) -> str | None:
    """Read the authoritative VERSION value from the tagged tree."""
    raw = git_raw("show", f"{tag}:{VERSION_FILE}")
    if raw is None:
        return None
    return raw.strip()


def protected_paths(asset: Path) -> set[Path]:
    paths = {asset, rooted_path(VERSION_FILE), Path(__file__).resolve()}
    paths.update(rooted_path(rel) for rel in REQUIRED + OPTIONAL)
    return paths


def validate(args: argparse.Namespace) -> list[str]:
    """Collect every release-contract violation in one pass."""
    problems: list[str] = []

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

    if args.excel is not None and not args.excel.strip():
        problems.append("--excel is blank; the certification environment must be recorded")

    if args.cases is not None and args.cases <= 0:
        problems.append(f"--cases must be a positive integer, got {args.cases}")
    if args.assertions is not None and args.assertions <= 0:
        problems.append(f"--assertions must be a positive integer, got {args.assertions}")
    if args.failures is not None and args.failures != 0:
        problems.append(
            f"--failures must be exactly 0 to publish, got {args.failures}. "
            "A suite with failures has no publishable manifest."
        )

    version_valid = args.version is not None and VERSION_RE.fullmatch(args.version) is not None
    tag_valid = args.tag is not None and TAG_RE.fullmatch(args.tag) is not None
    if args.version is not None and not version_valid:
        problems.append(f"--version {args.version!r} is not a supported SemVer release version")
    if args.tag is not None and not tag_valid:
        problems.append(f"--tag {args.tag!r} is not a supported v-prefixed SemVer release tag")
    if version_valid and tag_valid and args.tag != f"v{args.version}":
        problems.append(
            f"--version {args.version} and --tag {args.tag} disagree; "
            f"expected --tag v{args.version}"
        )

    asset_path: Path | None = None
    if args.asset is not None:
        asset_path = rooted_path(args.asset)
        if not asset_path.exists():
            problems.append(f"--asset not found at {args.asset}")
        elif not asset_path.is_file():
            problems.append(f"--asset is not a regular file: {args.asset}")

    if args.out is not None:
        out_path = rooted_path(args.out)
        tracked = git_raw("ls-files", "--cached", "-z")
        if tracked is None:
            problems.append("git ls-files failed, so --out could not be checked for tracked files")
        elif out_path in {rooted_path(rel) for rel in tracked.split("\0") if rel}:
            problems.append(
                f"--out {args.out!r} aliases a tracked file; "
                "the manifest must not overwrite repository content"
            )
        if asset_path is not None and out_path in protected_paths(asset_path):
            problems.append(
                f"--out {args.out!r} aliases a protected release input; "
                "the manifest must be written to a distinct path"
            )

    dirty = git("status", "--porcelain", "--untracked-files=no")
    if dirty is None:
        problems.append("git status failed, so the working tree could not be verified")
    elif dirty:
        paths = [line[3:] for line in dirty.splitlines()][:20]
        problems.append("tracked files are modified; hashing them would not match the commit claimed")
        problems.extend(f"  modified: {path}" for path in paths)
    problems.extend(tracked_content_problems())

    head = git("rev-parse", "HEAD")
    if head is None:
        problems.append("HEAD could not be resolved")

    if args.tag is not None:
        ref_type = tag_object_type(args.tag)
        if ref_type is None:
            problems.append(f"tag {args.tag} does not exist in this clone")
            problems.append("  Create/fetch the annotated release tag before generating provenance.")
        elif ref_type != "tag":
            problems.append(
                f"tag {args.tag} is a {ref_type} object, not an annotated tag; "
                "create the release tag with git tag -a"
            )
        else:
            target = resolve_tag_commit(args.tag)
            if target is None:
                problems.append(f"annotated tag {args.tag} could not be resolved to a commit")
            elif head is not None and head != target:
                problems.append(
                    f"HEAD {head} is not the target of {args.tag} ({target}). "
                    "Check out the tag before generating provenance."
                )
            else:
                problems.extend(verify_against_tag(args.tag, REQUIRED + OPTIONAL))

            if version_valid:
                source_version = tagged_version(args.tag)
                if source_version is None:
                    problems.append(f"{VERSION_FILE} is missing from tagged source {args.tag}")
                elif source_version != args.version:
                    problems.append(
                        f"--version {args.version} disagrees with {VERSION_FILE} at {args.tag} "
                        f"({source_version!r})"
                    )

    if asset_path is not None and asset_path.is_file():
        tag_ok = args.tag is not None and tag_object_type(args.tag) == "tag"
        problems.extend(package_problems(
            asset_path, args.version if version_valid else None,
            args.tag if tag_ok else None,
        ))

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

    asset_path = rooted_path(args.asset)
    out.append("")
    out.append("**Release assets**")
    out.append("")
    out.append("| Asset | SHA-256 |")
    out.append("|---|---|")
    out.append(f"| `{asset_path.name}` | `{sha256(asset_path)}` |")
    out.append("")
    out.append("**Package contents**")
    out.append("")
    out.append("| File | SHA-256 |")
    out.append("|---|---|")
    out.extend(f"| `{name}` | `{digest}` |" for name, digest in package_members(asset_path).items())

    out.append("")
    out.append("### Source integrity")
    out.append("")
    out.append(f"`HEAD` is the target of annotated tag `{args.tag}`, `{VERSION_FILE}` agrees with")
    out.append("the requested release identity, no tracked file is modified, and every")
    out.append(f"hashed source file matches its blob at `{args.tag}`. The package holds")
    out.append(f"exactly the workbook, `LICENSE` and `NOTICE`, and the latter two are")
    out.append(f"byte-identical to their blobs at `{args.tag}`.")
    out.append("")
    out.append("### What this establishes")
    out.append("")
    out.append("| Claim | Established by |")
    out.append("|---|---|")
    out.append("| The published source files are the tagged ones | comparison against the tag |")
    out.append("| The manifest describes the tagged commit | `HEAD` equals the annotated tag target |")
    out.append(f"| The release version is the tagged source version | `{VERSION_FILE}` comparison |")
    out.append("| A downloaded file is the one published here | its SHA-256 |")
    out.append("| The package carries the tagged LICENSE and NOTICE | package verification against the tag |")
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
    out.append(f'Get-FileHash -Algorithm SHA256 ".\\{asset_path.name}"')
    out.append(f'Expand-Archive ".\\{asset_path.name}" -DestinationPath ".\\{asset_path.stem}"')
    out.append(f'Get-FileHash -Algorithm SHA256 ".\\{asset_path.stem}\\*"')
    out.append("```")
    out.append("")
    out.append("Compare the package hash with the release asset table and each extracted")
    out.append("file with the package contents table. Keep `LICENSE` and `NOTICE` with the")
    out.append("workbook when redistributing it.")
    out.append("")
    out.append("The source files are text and are also in the repository at the commit")
    out.append("above, so `git show <commit>:<path>` gives the same bytes. The workbook is")
    out.append("only distributed inside the Release package, which is why its hash is")
    out.append("published here.")
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
            "tag_is_annotated": True,
            "version_matches_tagged_source": True,
            "tracked_files_unmodified": True,
            "package_contents_verified": True,
            "workbook_built_from_source": False,
            "workbook_build_is_manual": True,
        },
    }
    asset_path = rooted_path(args.asset)
    manifest["sha256"][asset_path.name] = sha256(asset_path)
    manifest["package"] = {"name": asset_path.name, "members": package_members(asset_path)}
    return manifest


def write_atomically(path: Path, text: str) -> None:
    """Write through a temporary file in the destination directory, then replace."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".manifest-", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(text)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--version", help="Release version, e.g. 1.4.0 or 1.4.1-rc.1")
    ap.add_argument("--tag", help="Annotated release tag, e.g. v1.4.0; HEAD must be its target")
    ap.add_argument("--asset", help="Path to the Release package, VBA-Performance-Manager-vX.Y.Z.zip")
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
        for problem in problems:
            print(f"  {problem}", file=sys.stderr)
        return EXIT_CONTRACT

    commit = git("rev-parse", "HEAD")
    tag_target = resolve_tag_commit(args.tag)
    markdown = build_markdown(args, commit)

    if args.out:
        out_path = rooted_path(args.out)
        try:
            write_atomically(
                out_path,
                json.dumps(build_manifest(args, commit, tag_target), indent=2) + "\n",
            )
        except (OSError, ValueError) as exc:
            print(f"manifest write failed; no provenance was emitted: {exc}", file=sys.stderr)
            return EXIT_CONTRACT
        print(f"manifest written to {args.out}", file=sys.stderr)

    print(markdown)
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
