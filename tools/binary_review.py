"""Tracked binary files pass only when reviewed explicitly (#84).

A tracked file is binary when .gitattributes unsets its ``text`` attribute
(the office-doc, opaque and vba-form-bin macros), or when it contains a NUL
byte in its first 8000 bytes, the same window git uses to sniff binaries. A
binary file passes only if:

- it is a UserForm .frx under src/, test/ or demo/ whose .frm is tracked
  beside it, the one approved path; or
- tools/binary-allowlist.json lists it with its current SHA-256 and a reason.

A committed workbook or a changed image therefore fails until someone updates
its allowlist entry on purpose. Entries for files that are gone or no longer
binary fail too, so the list cannot rot. The check never changes a file.
"""
from __future__ import annotations

import hashlib
import json
import re
import subprocess
from pathlib import Path, PurePosixPath
from typing import Any

ALLOWLIST = "tools/binary-allowlist.json"
SNIFF_BYTES = 8000
CHECK = "tracked binaries are reviewed"
SHA256 = re.compile(r"[0-9a-f]{64}")
FORM_ROOTS = ("src", "test", "demo")


def _git(root: Path, *args: str, stdin: str | None = None) -> str:
    return subprocess.run(["git", *args], cwd=root, input=stdin, text=True,
                          capture_output=True, check=True,
                          encoding="utf-8", errors="replace").stdout


def text_unset(root: Path, names: list[str]) -> set[str]:
    """The paths whose .gitattributes text attribute is unset (binary)."""
    if not names:
        return set()
    fields = _git(root, "check-attr", "-z", "--stdin", "text",
                  stdin="\0".join(names) + "\0").split("\0")
    return {fields[i] for i in range(0, len(fields) - 2, 3) if fields[i + 2] == "unset"}


def has_nul(path: Path) -> bool:
    with path.open("rb") as handle:
        return b"\0" in handle.read(SNIFF_BYTES)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def approved(name: str, tracked: set[str]) -> bool:
    """A UserForm binary is approved when its .frm is tracked beside it."""
    path = PurePosixPath(name)
    return (path.suffix.lower() == ".frx" and path.parts[0] in FORM_ROOTS
            and str(path.with_suffix(".frm")) in tracked)


def parse_allowlist(text: str) -> tuple[dict[str, str], list[str]]:
    """Return {path: sha256} and the problems found in the allowlist itself."""
    problems: list[str] = []
    try:
        data: Any = json.loads(text)
    except json.JSONDecodeError as exc:
        return {}, [f"{ALLOWLIST}: invalid JSON: {exc}"]
    if not isinstance(data, dict) or data.get("schema_version") != 1 \
            or not isinstance(data.get("files"), dict):
        return {}, [f"{ALLOWLIST}: expected {{\"schema_version\": 1, \"files\": {{...}}}}"]
    entries: dict[str, str] = {}
    for name, entry in data["files"].items():
        digest = entry.get("sha256") if isinstance(entry, dict) else None
        reason = entry.get("reason") if isinstance(entry, dict) else None
        if not isinstance(digest, str) or not SHA256.fullmatch(digest):
            problems.append(f"{ALLOWLIST}: {name}: sha256 must be 64 lowercase hex digits")
        elif not isinstance(reason, str) or not reason.strip():
            problems.append(f"{ALLOWLIST}: {name}: a reason is required")
        else:
            entries[name] = digest
    return entries, problems


def review(binaries: dict[str, str], tracked: set[str],
           allowlist: dict[str, str]) -> list[str]:
    """Check binaries ({path: sha256}) against the approved path and the allowlist."""
    problems: list[str] = []
    for name, digest in sorted(binaries.items()):
        if approved(name, tracked):
            continue
        expected = allowlist.get(name)
        if expected is None:
            problems.append(f"{name}: tracked binary is neither on an approved path nor in "
                            f"{ALLOWLIST}; review it and add an entry with its SHA-256")
        elif expected != digest:
            problems.append(f"{name}: SHA-256 {digest} differs from the allowlisted {expected}; "
                            "review the change and update the entry")
    for name in sorted(allowlist):
        if name not in tracked:
            problems.append(f"{ALLOWLIST}: {name} is not a tracked file; remove the entry")
        elif name not in binaries:
            problems.append(f"{ALLOWLIST}: {name} is not binary; remove the entry")
    return problems


def repository_problems(root: Path) -> list[str]:
    """Review every tracked binary in the repository at root."""
    try:
        tracked = sorted(n for n in _git(root, "ls-files", "-z").split("\0") if n)
        unset = text_unset(root, tracked)
    except (OSError, subprocess.CalledProcessError) as exc:
        return [f"git could not list the tracked files, so binaries were not reviewed: {exc}"]
    binaries: dict[str, str] = {}
    problems: list[str] = []
    for name in tracked:
        path = root / name
        if not path.is_file():
            continue
        if name in unset or has_nul(path):
            binaries[name] = sha256(path)
    allowlist_path = root / ALLOWLIST
    if not allowlist_path.exists():
        return [f"{ALLOWLIST} not found"]
    entries, allowlist_problems = parse_allowlist(allowlist_path.read_text(encoding="utf-8"))
    problems += allowlist_problems
    problems += review(binaries, set(tracked), entries)
    return problems
