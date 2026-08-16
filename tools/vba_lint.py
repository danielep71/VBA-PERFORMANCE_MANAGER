#!/usr/bin/env python3
"""
vba_lint.py — static consistency checks for the Class Performance Manager sources.

These checks exist because every one of them corresponds to a defect that
actually reached the repository during v1.2.0 development. None of them require
Excel, so they can run on a hosted runner and gate every push.

    python3 tools/vba_lint.py

Exit code 0 when all checks pass, 1 otherwise.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

ROOT = Path(__file__).resolve().parent.parent

CLASS_FILE = ROOT / "src" / "classes" / "cPerformanceManager.cls"
MODULE_FILE = ROOT / "src" / "modules" / "M_cPM_TIMEWASTERS.bas"
TEST_FILE = ROOT / "test" / "M_cPM_Test.bas"
README_FILE = ROOT / "README.md"

VBA_SOURCES = [CLASS_FILE, MODULE_FILE, TEST_FILE]
ALL_SOURCES = VBA_SOURCES + [
    ROOT / "demo" / "M_cPM_DEMO.bas",
    ROOT / "demo" / "M_cPM_USAGE_EXAMPLES.bas",
    ROOT / "demo" / "M_DEMO_BUILDER.bas",
]

# Names VBA reserves. Using one as a variable produces a bare "Syntax error"
# with no indication of the cause, which is how Dim Empty() reached the repo.
RESERVED_WORDS = {
    "empty", "null", "nothing", "true", "false", "error", "date", "time",
    "string", "name", "line", "input", "print", "type", "base", "stop",
    "next", "loop", "step", "then", "else", "case", "each", "exit",
    "resume", "new", "me", "and", "or", "not", "mod", "is", "like",
    "to", "as", "byval", "byref", "optional", "call", "dim", "redim",
    "erase", "static", "public", "private", "friend", "const", "declare",
    "sub", "function", "property", "end", "select", "with", "do", "while",
    "until", "for", "if", "option", "set", "let", "get", "on", "goto",
}

PROC_START = re.compile(
    r"^(?:Public|Private|Friend)(?:\s+Static)?\s+"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.M,
)
PROC_END = re.compile(r"^End\s+(?:Sub|Function|Property)\s*$")


# --------------------------------------------------------------------------- #
# Result plumbing
# --------------------------------------------------------------------------- #

@dataclass
class Report:
    failures: list[str] = field(default_factory=list)
    checks_run: int = 0

    def check(self, name: str, problems: list[str]) -> None:
        self.checks_run += 1
        if problems:
            self.failures.append(name)
            print(f"FAIL  {name}")
            for p in problems[:20]:
                print(f"        {p}")
            if len(problems) > 20:
                print(f"        ... and {len(problems) - 20} more")
        else:
            print(f"ok    {name}")


def strip_strings(text: str) -> str:
    """Blank out string literals.

    Assertion messages routinely contain procedure names, so identifier scans
    must not treat string content as code.
    """
    return re.sub(r'"[^"\n]*"', '""', text)


def read(path: Path) -> str:
    """Read a source file, normalising line endings for analysis only."""
    return path.read_text(encoding="utf-8", errors="replace").replace("\r\n", "\n")


def code_lines(text: str) -> list[tuple[int, str]]:
    """Numbered lines with comment-only lines removed."""
    out = []
    for i, line in enumerate(text.split("\n"), 1):
        if line.lstrip().startswith("'"):
            continue
        out.append((i, line))
    return out


# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #

def check_conflict_markers(rep: Report) -> None:
    """A conflicted README was published to the repository homepage."""
    problems = []
    for path in ALL_SOURCES + [README_FILE, ROOT / "CHANGELOG.md"]:
        if not path.exists():
            continue
        for i, line in enumerate(read(path).split("\n"), 1):
            if line.startswith(("<<<<<<<", ">>>>>>>")) or line.rstrip() == "=======":
                problems.append(f"{path.name}:{i}: merge conflict marker")
    rep.check("no merge conflict markers", problems)


def check_procedure_balance(rep: Report) -> None:
    """A dropped declaration line left an orphaned End Sub."""
    problems = []
    for path in ALL_SOURCES:
        if not path.exists():
            continue
        depth, current = 0, None
        for i, line in code_lines(read(path)):
            m = PROC_START.match(line)
            if m:
                if depth:
                    problems.append(f"{path.name}:{i}: '{m.group(1)}' starts inside '{current}'")
                depth, current = 1, m.group(1)
            elif PROC_END.match(line):
                if not depth:
                    problems.append(f"{path.name}:{i}: End with no open procedure")
                depth = 0
        if depth:
            problems.append(f"{path.name}: '{current}' is never closed")
    rep.check("procedure blocks balanced", problems)


def check_block_pairs(rep: Report) -> None:
    """Unbalanced Select/With/If/Type blocks."""
    pairs = [
        (r"^\s*Select\s+Case\b", r"^\s*End\s+Select\b", "Select Case"),
        (r"^\s*With\s+\S", r"^\s*End\s+With\b", "With"),
        (r"^\s*(?:Public|Private)\s+Type\s+\w+", r"^\s*End\s+Type\b", "Type"),
        (r"^\s*(?:Public|Private)\s+Enum\s+\w+", r"^\s*End\s+Enum\b", "Enum"),
        (r"^\s*#If\b", r"^\s*#End\s+If\b", "#If"),
    ]
    problems = []
    for path in ALL_SOURCES:
        if not path.exists():
            continue
        text = "\n".join(l for _, l in code_lines(read(path)))
        for open_re, close_re, label in pairs:
            o = len(re.findall(open_re, text, re.M))
            c = len(re.findall(close_re, text, re.M))
            if o != c:
                problems.append(f"{path.name}: {label} {o} open / {c} close")
    rep.check("block pairs balanced", problems)


def check_reserved_words(rep: Report) -> None:
    """Dim Empty() compiled to a bare 'Syntax error' with no explanation."""
    problems = []
    # "Private Type Foo" declares a type, not a variable called Type. Skip the
    # keywords that can legally follow a declaration modifier.
    not_a_variable = {
        "type", "declare", "enum", "sub", "function", "property",
        "const", "withevents", "static", "ptrsafe",
    }
    dim_re = re.compile(r"^\s*(?:Dim|Private|Public|Const|Static)\s+(\w+)", re.I)
    for path in ALL_SOURCES:
        if not path.exists():
            continue
        for i, line in code_lines(read(path)):
            m = dim_re.match(line)
            if not m:
                continue
            name = m.group(1)
            if name.lower() in not_a_variable:
                continue
            if name.lower() in RESERVED_WORDS:
                problems.append(f"{path.name}:{i}: '{name}' is a VBA reserved word")
    rep.check("no reserved words as identifiers", problems)


def check_error_sources(rep: Report) -> None:
    """Error sources named M_cPM_ReportHelpers and M_cPM_RegressionTests,
    neither of which has ever existed."""
    module_names = {
        CLASS_FILE.name: "cPerformanceManager",
        MODULE_FILE.name: "M_cPM_TimeWasters",
        TEST_FILE.name: "M_cPM_Test",
    }
    src_re = re.compile(r'"([A-Za-z_]\w*)\.([A-Za-z_]\w*)"')
    problems = []
    for path in VBA_SOURCES:
        if not path.exists():
            continue
        expected_module = module_names[path.name]
        current = None
        for i, line in enumerate(read(path).split("\n"), 1):
            m = PROC_START.match(line)
            if m:
                current = m.group(1)
                continue
            if line.lstrip().startswith("'"):
                continue
            for sm in src_re.finditer(line):
                mod, proc = sm.group(1), sm.group(2)
                if proc != current:
                    continue
                if mod.lower() != expected_module.lower():
                    problems.append(
                        f"{path.name}:{i}: source says '{mod}.{proc}' "
                        f"but this is {expected_module}"
                    )
    rep.check("error sources name the right module", problems)


def check_no_bare_error_numbers(rep: Report) -> None:
    """All raised error numbers must be named constants.

    Scoped to the shipped component. The regression harness raises three of its
    own bare offsets (2400-2402); those are reported separately as a note rather
    than a failure, because they are not part of the distributed surface.
    """
    problems = []
    for path in (CLASS_FILE, MODULE_FILE):
        if not path.exists():
            continue
        for i, line in code_lines(read(path)):
            if re.search(r"Err\.Raise\s+vbObjectError\s*\+\s*\d+", line):
                problems.append(f"{path.name}:{i}: bare vbObjectError offset in Err.Raise")
    rep.check("no bare error numbers (shipped source)", problems)

    if TEST_FILE.exists():
        harness = [
            f"{TEST_FILE.name}:{i}"
            for i, line in code_lines(read(TEST_FILE))
            if re.search(r"Err\.Raise\s+vbObjectError\s*\+\s*\d+", line)
        ]
        if harness:
            print(f"note  test harness has {len(harness)} bare error numbers "
                  f"({', '.join(harness)}) - not gated")


def check_undefined_callees(rep: Report) -> None:
    """Renaming a private helper without updating its callers.

    Two things must be excluded or this produces noise: string literals, because
    assertion messages contain procedure names, and member access, because
    cPM.Stats_Min resolves against the class rather than the calling file.
    """
    problems = []
    prefixes = ("Elapsed_", "Method_", "Checkpoint_", "Stats_", "Timestamp_",
                "QPC_", "Start_", "Read_", "SystemTime_", "PM_TW_")
    call_re = re.compile(r"(?<![.\w])((?:%s)\w+)\b" % "|".join(prefixes))

    for path in VBA_SOURCES:
        if not path.exists():
            continue
        raw = read(path)
        code = strip_strings("\n".join(l for _, l in code_lines(raw)))

        defined = {m.group(1) for m in PROC_START.finditer(raw)}
        defined |= {m.group(1) for m in
                    re.finditer(r"^\s*(?:Public|Private)\s+Enum\s+(\w+)", raw, re.M)}

        called = {m.group(1) for m in call_re.finditer(code)}
        # PM_TW_* live in the companion module; cross-file calls are expected.
        local = {c for c in called if not c.startswith("PM_TW_")}

        for name in sorted(local - defined):
            problems.append(f"{path.name}: calls '{name}' which is not defined here")

    rep.check("no undefined local callees", problems)


def check_test_wiring(rep: Report) -> None:
    """A test can be defined but never called, or called but never defined."""
    if not TEST_FILE.exists():
        rep.check("test wiring", [f"{TEST_FILE} not found"])
        return
    text = read(TEST_FILE)
    defined = set(re.findall(r"^Private Sub (Test_[A-Za-z_0-9]+)\(\)", text, re.M))
    defined -= {d for d in defined if d.startswith("Test_Assert_")}
    called = set(re.findall(r"^\s+(Test_[A-Za-z_0-9]+)\s*$", text, re.M))
    problems = []
    for name in sorted(defined - called):
        problems.append(f"defined but never called: {name}")
    for name in sorted(called - defined):
        problems.append(f"called but never defined: {name}")
    rep.check("every test is defined and called", problems)


def check_total_steps(rep: Report) -> None:
    """TotalSteps drifted from 41 to 52 to 62 without being updated."""
    if not TEST_FILE.exists():
        rep.check("TotalSteps matches case count", [f"{TEST_FILE} not found"])
        return
    text = read(TEST_FILE)
    m = re.search(r"TotalSteps\s+As Long\s*=\s*(\d+)", text)
    if not m:
        rep.check("TotalSteps matches case count", ["TotalSteps constant not found"])
        return
    declared = int(m.group(1))
    called = set(re.findall(r"^\s+(Test_[A-Za-z_0-9]+)\s*$", text, re.M))
    problems = []
    if declared != len(called):
        problems.append(f"TotalSteps = {declared} but {len(called)} cases are invoked")
    rep.check("TotalSteps matches case count", problems)


def check_version_consistency(rep: Report) -> None:
    """The class and module shipped as 1.2.1 while everything else said 1.2.0."""
    versions: dict[str, str] = {}
    if CLASS_FILE.exists():
        m = re.search(r"VERSION:\s*(\d+\.\d+\.\d+)", read(CLASS_FILE))
        if m:
            versions[CLASS_FILE.name] = m.group(1)
    for path in (MODULE_FILE, TEST_FILE):
        if path.exists():
            m = re.search(r"^' VERSION\n'\s+(\d+\.\d+\.\d+)", read(path), re.M)
            if m:
                versions[path.name] = m.group(1)
    if README_FILE.exists():
        m = re.search(r"Version-(\d+\.\d+\.\d+)", read(README_FILE))
        if m:
            versions["README.md"] = m.group(1)

    problems = []
    distinct = set(versions.values())
    if len(distinct) > 1:
        for name, v in sorted(versions.items()):
            problems.append(f"{name}: {v}")
    elif not versions:
        problems.append("no version stamp found in any file")
    rep.check("version stamps agree", problems)


def check_api_declarations(rep: Report) -> None:
    """Native timing APIs must be called from exactly one place each."""
    if not CLASS_FILE.exists():
        rep.check("native APIs have a single call site", [f"{CLASS_FILE} not found"])
        return
    text = "\n".join(l for _, l in code_lines(read(CLASS_FILE)))
    problems = []
    for api in ("QueryPerformanceCounter", "timeGetSystemTime"):
        # Declarations read "Name Lib \"kernel32\" (...)", so "Name(" only ever
        # matches a genuine call site.
        calls = len(re.findall(rf"\b{api}\(", text))
        if calls != 1:
            problems.append(f"{api}: {calls} call sites, expected exactly 1")
    rep.check("native APIs have a single call site", problems)


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #

def main() -> int:
    missing = [p for p in VBA_SOURCES if not p.exists()]
    if missing:
        for p in missing:
            print(f"FAIL  source not found: {p}")
        return 1

    print("VBA static checks")
    print("-" * 60)

    rep = Report()
    check_conflict_markers(rep)
    check_procedure_balance(rep)
    check_block_pairs(rep)
    check_reserved_words(rep)
    check_error_sources(rep)
    check_no_bare_error_numbers(rep)
    check_undefined_callees(rep)
    check_test_wiring(rep)
    check_total_steps(rep)
    check_version_consistency(rep)
    check_api_declarations(rep)

    print("-" * 60)
    if rep.failures:
        print(f"{len(rep.failures)} of {rep.checks_run} checks failed:")
        for f in rep.failures:
            print(f"  - {f}")
        return 1

    print(f"all {rep.checks_run} checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
