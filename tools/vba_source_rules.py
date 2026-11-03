"""Per-component VBA source rules for the static gate (#83).

Source-only: none of these compiles or runs VBA. Each check reads the tracked
exports as bytes and reports ``path:line: problem``.

- SOURCE_CHECK: the component's ``Attribute VB_Name`` matches its file name
  (VBA names and Windows file names are both case-insensitive, so case is not
  compared); the file is ASCII; ``Option Explicit`` is declared before the
  first procedure and ``Option Base`` is not used; no line holds a tab or
  trailing white space; every string literal closes on its line.
- PTRSAFE_CHECK: in each conditional-compilation profile that
  tools/vba_compile_safety.py models, every reachable ``Declare`` carries
  ``PtrSafe`` under VBA7, and none carries it before VBA7, where the keyword
  does not compile.
- VISIBILITY_CHECK: every procedure and every module-level declaration states
  Public, Private or Friend. ``Dim`` and ``Global`` at module level, and a
  declaration with no keyword, fail.
- PUBLIC_NAME_CHECK: no public name (procedure, variable, constant, Declare,
  Type, Enum or Enum member) is defined in two standard modules in any
  profile. All of the project's standard modules share one namespace, so an
  unqualified reference to such a name is an "Ambiguous name detected" error.
"""
from __future__ import annotations

import re
from collections.abc import Iterator
from pathlib import PurePosixPath

from vba_compile_safety import IDENT, PROFILES, comma_parts, lines

SOURCE_CHECK = "VBA components follow the source rules"
PTRSAFE_CHECK = "Declare statements follow each profile's PtrSafe rule"
VISIBILITY_CHECK = "procedures and module declarations state their visibility"
PUBLIC_NAME_CHECK = "public names are unique across standard modules"

HEADER = re.compile(r"^(?:VERSION\s|Object\s*=|Attribute\s)", re.I)
VB_NAME = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.I | re.M)
PROC = re.compile(rf"^(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?"
                  rf"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+({IDENT})", re.I)
PROC_END = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.I)
BLOCK = re.compile(rf"^(?:(Public|Private)\s+)?(Type|Enum)\s+({IDENT})", re.I)
BLOCK_END = re.compile(r"^End\s+(?:Type|Enum)\b", re.I)
DECLARE = re.compile(rf"^(?:(Public|Private)\s+)?Declare\s+(PtrSafe\s+)?(?:Sub|Function)\s+({IDENT})",
                     re.I)
MODULE_DECLARATION = re.compile(
    r"^(?:(Public|Private|Global|Dim|Friend)\s+)?(Const|Declare|Event|WithEvents)?\b", re.I)
SKIPPED_STATEMENT = re.compile(r"^(?:Option|Implements|Attribute|DefBool|DefByte|DefInt|DefLng|"
                               r"DefLngLng|DefLngPtr|DefCur|DefSng|DefDbl|DefDec|DefDate|DefStr|"
                               r"DefObj|DefVar)\b", re.I)
OPTION_EXPLICIT = re.compile(r"^Option\s+Explicit\b", re.I)
OPTION_BASE = re.compile(r"^Option\s+Base\b", re.I)


# --------------------------------------------------------------------------- #
# Reading a component
# --------------------------------------------------------------------------- #

def header_length(physical: list[str]) -> int:
    """How many leading lines are the export header the VBE writes."""
    depth = 0
    for index, line in enumerate(physical):
        word = line.strip().split(" ", 1)[0].upper()
        if word == "BEGIN":
            depth += 1
        elif word == "END" and depth:
            depth -= 1
        elif not depth and not HEADER.match(line):
            return index
    return len(physical)


def physical_lines(text: str) -> list[str]:
    """The lines of a CRLF (or LF) export, without their line ends."""
    found = text.replace("\r\n", "\n").split("\n")
    if found and found[-1] == "":
        found.pop()
    return found


def body_text(text: str) -> str:
    """The text with its header lines blanked, so line numbers are kept."""
    physical = physical_lines(text)
    split = header_length(physical)
    return "\n".join([""] * split + physical[split:]) + "\n"


def code_part(line: str) -> tuple[str, bool]:
    """(the line up to its comment, whether a string literal is left open)."""
    in_literal = False
    out: list[str] = []
    for index, char in enumerate(line):
        if char == '"':
            in_literal = not in_literal
        elif not in_literal and char == "'":
            break
        elif not in_literal and re.match(r"Rem(?:\s|$)", line[index:], re.I) \
                and re.search(r"(?:^|:)\s*$", "".join(out)):
            break       # Rem starts a comment only at the start of a statement
        out.append(char)
    return "".join(out), in_literal


def statements(text: str, vba7: bool, win64: bool) -> Iterator[tuple[int, str]]:
    """(first line, statement) for each active statement of one profile."""
    for number, logical in lines(body_text(text), vba7, win64):
        for statement in re.split(r":(?!=)", logical):
            statement = statement.strip()
            if statement:
                yield number, statement


# --------------------------------------------------------------------------- #
# Source rules
# --------------------------------------------------------------------------- #

def source_problems(path: str, data: bytes) -> list[str]:
    """The source-rule problems of one export, read as bytes."""
    problems: list[str] = []
    text = data.decode("latin-1")
    stem = PurePosixPath(path).stem
    name = VB_NAME.search(text)
    if not name:
        problems.append(f"{path}: no Attribute VB_Name")
    elif name.group(1).lower() != stem.lower():
        problems.append(f"{path}: VB_Name {name.group(1)!r} does not match the file name {stem!r}")
    physical = physical_lines(text)
    split = header_length(physical)
    explicit: int | None = None          # line of Option Explicit
    first_procedure: int | None = None
    comment_continues = False
    for number, line in enumerate(physical, 1):
        if any(ord(c) > 127 for c in line):
            problems.append(f"{path}:{number}: non-ASCII character")
        if "\t" in line:
            problems.append(f"{path}:{number}: tab character")
        if line != line.rstrip(" \t"):
            problems.append(f"{path}:{number}: trailing white space")
        if number <= split:
            continue
        if comment_continues:
            # A comment that ends in " _" continues onto this line.
            comment_continues = line.rstrip().endswith(" _")
            continue
        code, open_literal = code_part(line)
        comment_continues = code != line and line.rstrip().endswith(" _")
        if open_literal:
            problems.append(f"{path}:{number}: string literal is not closed")
        statement = code.strip()
        if OPTION_BASE.match(statement):
            problems.append(f"{path}:{number}: Option Base is not used; arrays state their bounds")
        if OPTION_EXPLICIT.match(statement) and explicit is None:
            explicit = number
        if PROC.match(statement) and first_procedure is None:
            first_procedure = number
    if explicit is None:
        problems.append(f"{path}: no Option Explicit")
    elif first_procedure is not None and explicit > first_procedure:
        problems.append(f"{path}:{explicit}: Option Explicit must precede the first procedure")
    return problems


def ptrsafe_problems(path: str, text: str) -> list[str]:
    problems: set[str] = set()
    for profile, vba7, win64 in PROFILES:
        try:
            for number, statement in statements(text, vba7, win64):
                declare = DECLARE.match(statement)
                if not declare:
                    continue
                if vba7 and not declare.group(2):
                    problems.add(f"{path}:{number}: [{profile}] Declare {declare.group(3)} "
                                 "needs PtrSafe")
                if not vba7 and declare.group(2):
                    problems.add(f"{path}:{number}: [{profile}] Declare {declare.group(3)} "
                                 "uses PtrSafe, which does not compile before VBA7")
        except ValueError as exc:
            problems.add(f"{path}: [{profile}] unsupported source: {exc}")
    return sorted(problems, key=_order)


def visibility_problems(path: str, text: str) -> list[str]:
    problems: set[str] = set()
    for profile, vba7, win64 in PROFILES:
        try:
            problems |= set(_visibility(path, text, vba7, win64))
        except ValueError as exc:
            problems.add(f"{path}: [{profile}] unsupported source: {exc}")
    return sorted(problems, key=_order)


def _visibility(path: str, text: str, vba7: bool, win64: bool) -> Iterator[str]:
    in_procedure = in_block = False
    for number, statement in statements(text, vba7, win64):
        procedure = PROC.match(statement)
        if procedure:
            in_procedure = True
            keyword = (procedure.group(1) or "").lower()
            if keyword in ("", "global"):
                yield (f"{path}:{number}: procedure {procedure.group(2)} needs Public, Private "
                       "or Friend")
            continue
        if PROC_END.match(statement):
            in_procedure = False
            continue
        if in_procedure:
            continue
        if in_block:
            in_block = not BLOCK_END.match(statement)
            continue
        block = BLOCK.match(statement)
        if block:
            in_block = True
            if not block.group(1):
                yield f"{path}:{number}: {block.group(2)} {block.group(3)} needs Public or Private"
            continue
        if SKIPPED_STATEMENT.match(statement) or statement.startswith("#"):
            continue
        keyword = (MODULE_DECLARATION.match(statement) or _never()).group(1)
        word = (keyword or "").lower()
        if word == "dim":
            yield f"{path}:{number}: module-level Dim; use Private or Public"
        elif word == "global":
            yield f"{path}:{number}: Global; use Public"
        elif not word:
            yield f"{path}:{number}: module-level declaration needs Public or Private"


def _never() -> re.Match[str]:
    raise AssertionError("MODULE_DECLARATION matches every statement")


def public_names(text: str, vba7: bool, win64: bool) -> dict[str, tuple[int, str]]:
    """The project-wide names a standard module defines: lower name -> (line, name)."""
    names: dict[str, tuple[int, str]] = {}
    in_procedure = False
    block: str | None = None          # "public enum", "private enum" or "type"
    for number, statement in statements(text, vba7, win64):
        procedure = PROC.match(statement)
        if procedure:
            in_procedure = True
            if (procedure.group(1) or "public").lower() in ("public", "global"):
                names.setdefault(procedure.group(2).lower(), (number, procedure.group(2)))
            continue
        if PROC_END.match(statement):
            in_procedure = False
            continue
        if in_procedure:
            continue
        if block:
            if BLOCK_END.match(statement):
                block = None
            elif block == "public enum":
                member = re.match(IDENT, statement)
                if member:
                    names.setdefault(member.group(0).lower(), (number, member.group(0)))
            continue
        found = BLOCK.match(statement)
        if found:
            public = (found.group(1) or "public").lower() == "public"
            block = f"{'public' if public else 'private'} {found.group(2).lower()}"
            if public:
                names.setdefault(found.group(3).lower(), (number, found.group(3)))
            continue
        for name in declared_public(statement):
            names.setdefault(name.lower(), (number, name))
    return names


def declared_public(statement: str) -> list[str]:
    """The public names a module-level Declare, variable or Const statement defines."""
    declare = DECLARE.match(statement)
    if declare:
        return [declare.group(3)] if (declare.group(1) or "public").lower() == "public" else []
    variable = re.match(r"(Public|Global)\s+(?:Const\s+)?(.+)", statement, re.I)
    if not variable or re.match(r"(?:Declare|Type|Enum|Event)\b", variable.group(2), re.I):
        return []
    found = []
    for part in comma_parts(variable.group(2)):
        name = re.match(IDENT, re.sub(r"^WithEvents\s+", "", part, flags=re.I))
        if name:
            found.append(name.group(0))
    return found


def public_name_problems(sources: dict[str, str]) -> list[str]:
    """Public names defined in more than one standard module, in any profile."""
    problems: set[str] = set()
    modules = {path: text for path, text in sources.items() if path.lower().endswith(".bas")}
    for profile, vba7, win64 in PROFILES:
        owners: dict[str, list[tuple[str, int, str]]] = {}
        for path, text in sorted(modules.items()):
            try:
                found = public_names(text, vba7, win64)
            except ValueError as exc:
                problems.add(f"{path}: [{profile}] unsupported source: {exc}")
                continue
            for key, (number, name) in found.items():
                owners.setdefault(key, []).append((path, number, name))
        for places in owners.values():
            if len(places) > 1:
                where = ", ".join(f"{path}:{number}" for path, number, _ in places)
                problems.add(f"public name {places[0][2]} is defined in {len(places)} standard "
                             f"modules: {where}")
    return sorted(problems)


def _order(problem: str) -> tuple[str, int, str]:
    match = re.match(r"(.+?):(\d+):", problem)
    return (match.group(1), int(match.group(2)), problem) if match else (problem, 0, problem)


# --------------------------------------------------------------------------- #
# Repository and fixtures
# --------------------------------------------------------------------------- #

def repository_problems(sources: dict[str, bytes]) -> dict[str, list[str]]:
    """Problems per check for the tracked exports (path -> bytes)."""
    texts = {path: data.decode("latin-1") for path, data in sources.items()}
    result: dict[str, list[str]] = {SOURCE_CHECK: [], PTRSAFE_CHECK: [], VISIBILITY_CHECK: []}
    for path in sorted(sources):
        result[SOURCE_CHECK] += source_problems(path, sources[path])
        result[PTRSAFE_CHECK] += ptrsafe_problems(path, texts[path])
        result[VISIBILITY_CHECK] += visibility_problems(path, texts[path])
    result[PUBLIC_NAME_CHECK] = public_name_problems(texts)
    return result


GOOD_MODULE = (
    'Attribute VB_Name = "M_Good"\r\n'
    "    Option Explicit\r\n"
    "    Option Private Module\r\n"
    "    Private Const Greeting As String = \"it's \"\"quoted\"\"\"   'comment with \" inside\r\n"
    "    #If VBA7 Then\r\n"
    "        Private Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\r\n"
    "    #Else\r\n"
    "        Private Declare Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\r\n"
    "    #End If\r\n"
    "    Public Enum GoodKind\r\n"
    "        GoodFirst = 1\r\n"
    "    End Enum\r\n"
    "Public Sub Run()\r\n"
    "    Dim text As String: text = Greeting\r\n"
    "End Sub\r\n"
)

# (check, description, path, source) -> each must produce a problem for that check.
BAD_FIXTURES: tuple[tuple[str, str, str, str], ...] = (
    (SOURCE_CHECK, "VB_Name differs", "M_Good.bas",
     GOOD_MODULE.replace('"M_Good"', '"M_Other"')),
    (SOURCE_CHECK, "non-ASCII", "M_Good.bas", GOOD_MODULE.replace("comment with", "comment é")),
    (SOURCE_CHECK, "no Option Explicit", "M_Good.bas",
     GOOD_MODULE.replace("    Option Explicit\r\n", "")),
    (SOURCE_CHECK, "Option Base", "M_Good.bas",
     GOOD_MODULE.replace("Option Private Module", "Option Base 1")),
    (SOURCE_CHECK, "tab", "M_Good.bas", GOOD_MODULE.replace("    Dim text", "\tDim text")),
    (SOURCE_CHECK, "trailing white space", "M_Good.bas",
     GOOD_MODULE.replace("End Sub\r\n", "End Sub  \r\n")),
    (SOURCE_CHECK, "open literal", "M_Good.bas",
     GOOD_MODULE.replace("text = Greeting", 'text = "open')),
    (PTRSAFE_CHECK, "VBA7 Declare without PtrSafe", "M_Good.bas",
     GOOD_MODULE.replace("Declare PtrSafe Sub", "Declare Sub")),
    (PTRSAFE_CHECK, "pre-VBA7 Declare with PtrSafe", "M_Good.bas",
     GOOD_MODULE.replace("        Private Declare Sub", "        Private Declare PtrSafe Sub")),
    (VISIBILITY_CHECK, "procedure without visibility", "M_Good.bas",
     GOOD_MODULE.replace("Public Sub Run()", "Sub Run()")),
    (VISIBILITY_CHECK, "module-level Dim", "M_Good.bas",
     GOOD_MODULE.replace("    Private Const", "    Dim Counter As Long\r\n    Private Const")),
    (VISIBILITY_CHECK, "Const without visibility", "M_Good.bas",
     GOOD_MODULE.replace("    Private Const", "    Const")),
    (VISIBILITY_CHECK, "Enum without visibility", "M_Good.bas",
     GOOD_MODULE.replace("    Public Enum", "    Enum")),
)


def fixture_problems() -> dict[str, list[str]]:
    """Each check must accept the good module and reject its own bad fixtures."""
    found: dict[str, list[str]] = {SOURCE_CHECK: [], PTRSAFE_CHECK: [], VISIBILITY_CHECK: [],
                                   PUBLIC_NAME_CHECK: []}
    good = repository_problems({"M_Good.bas": GOOD_MODULE.encode("latin-1")})
    for check, problems in good.items():
        found[check] += [f"fixture: the good module fails: {p}" for p in problems]
    for check, description, path, source in BAD_FIXTURES:
        if not repository_problems({path: source.encode("latin-1")})[check]:
            found[check].append(f"fixture: {description} was not reported")
    twin = GOOD_MODULE.replace('"M_Good"', '"M_Twin"').replace("GoodKind", "TwinKind") \
        .replace("GoodFirst", "TwinFirst")
    pair = {"M_Good.bas": GOOD_MODULE.encode("latin-1"), "M_Twin.bas": twin.encode("latin-1")}
    if not repository_problems(pair)[PUBLIC_NAME_CHECK]:
        found[PUBLIC_NAME_CHECK].append("fixture: Run defined in two modules was not reported")
    return found
