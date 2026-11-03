"""The VBA components the static checks cover, derived from git (#82).

The checks used to name their files by hand, so a new component under src/,
test/ or demo/ escaped every check without a diagnostic. The component list
is now read from ``git ls-files``. Only the checks that apply to one role
still name a file, and that role assignment is explicit here and checked:

- every tracked .bas, .cls or .frm lies under src/, test/ or demo/;
- every role file is tracked and has the extension its role needs;
- release provenance hashes exactly the tracked components, so a new
  component cannot drop out of the release manifest either.
"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path, PurePosixPath

ROOTS = ("src", "test", "demo")
VBA_SUFFIXES = (".bas", ".cls", ".frm")

# Role files: the checks that inspect one specific component select it here.
ROLES = {
    "class": "src/classes/cPerformanceManager.cls",
    "module": "src/modules/M_cPM_TIMEWASTERS.bas",
    "test": "test/M_cPM_Test.bas",
}
ROLE_SUFFIX = {"class": ".cls", "module": ".bas", "test": ".bas"}

VB_NAME = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.M)


def is_vba(path: str) -> bool:
    return PurePosixPath(path).suffix.lower() in VBA_SUFFIXES


def tracked_vba(root: Path) -> list[str]:
    """Every tracked VBA component, as sorted POSIX paths relative to root.

    Raises RuntimeError when git cannot list the files, so the gate fails
    instead of checking an empty set.
    """
    try:
        listing = subprocess.run(
            ["git", "ls-files", "-z"], cwd=root, capture_output=True, check=True,
        ).stdout.decode("utf-8", errors="replace")
    except (OSError, subprocess.CalledProcessError) as exc:
        raise RuntimeError(f"git ls-files failed, so the VBA components are unknown: {exc}") from exc
    return sorted(path for path in listing.split("\0") if path and is_vba(path))


def under(paths: list[str], *roots: str) -> list[str]:
    """The paths that lie under any of the given top-level directories."""
    return [path for path in paths if PurePosixPath(path).parts[0] in roots]


def module_name(text: str) -> str | None:
    """The component's Attribute VB_Name, or None when the export has none."""
    match = VB_NAME.search(text)
    return match.group(1) if match else None


def inventory_problems(tracked: list[str], provenance: list[str],
                       roles: dict[str, str] = ROLES) -> list[str]:
    """Check the derived inventory; an empty list means every component is covered."""
    problems: list[str] = []
    covered = set(under(tracked, *ROOTS))
    for path in tracked:
        if path not in covered:
            problems.append(f"{path}: VBA component outside {', '.join(r + '/' for r in ROOTS)}; "
                            "no check covers it")
    for role, path in sorted(roles.items()):
        if path not in covered:
            problems.append(f"{role} role file {path} is not a tracked component")
        elif PurePosixPath(path).suffix.lower() != ROLE_SUFFIX.get(role, ""):
            problems.append(f"{role} role file {path} must be a {ROLE_SUFFIX.get(role)} export")
    if len(set(roles.values())) != len(roles):
        problems.append("two roles name the same file")
    hashed = set(provenance)
    for path in sorted(covered - hashed):
        problems.append(f"{path}: tracked component missing from release provenance "
                        "(REQUIRED or OPTIONAL in tools/release_provenance.py)")
    for path in sorted(hashed - covered):
        problems.append(f"{path}: listed by release provenance but not a tracked component")
    return problems
