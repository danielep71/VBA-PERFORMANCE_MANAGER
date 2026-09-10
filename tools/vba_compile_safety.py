"""Bounded, source-only VBA compile-safety analysis; not a VBA compiler.

Analyse each supported Windows conditional-compilation profile separately so a
declaration or label in an inactive branch cannot satisfy an active reference.
Unknown conditional syntax fails closed instead of silently skipping a branch.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

LABEL_CHECK = "error-handler labels resolve within their procedure"
ASSIGN_CHECK = "assigned identifiers are declared"
PROFILES = (("pre-VBA7", False, False), ("VBA7-Win32", True, False),
            ("VBA7-Win64", True, True))
IDENT = r"[A-Za-z_][A-Za-z_0-9]*"
PROC = re.compile(rf"^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
                  rf"(Sub|Function|Property\s+(?:Get|Let|Set))\s+({IDENT})"
                  r"[$%&!#@]?\s*(.*)$", re.I)


def mask(line: str) -> str:
    """Remove comments and mask escaped strings without losing separators."""
    out, i = [], 0
    while i < len(line):
        c = line[i]
        if c == "'":
            break
        if c == '"':
            out.append('""')
            i += 1
            while i < len(line):
                if line[i] == '"':
                    i += 1
                    if i < len(line) and line[i] == '"':
                        i += 1
                        continue
                    break
                i += 1
            continue
        # Rem is a comment only at a statement start (including inline If).
        if re.match(r"Rem(?:\s|$)", line[i:], re.I) and (
            not "".join(out).strip() or re.search(r"(?:[:]|\bThen|\bElse)\s*$", "".join(out), re.I)
        ):
            break
        out.append(c)
        i += 1
    return "".join(out)


def conditional(expr: str, env: dict[str, bool]) -> bool:
    tokens = re.findall(r"\w+|\S", expr.lower())
    pos = 0

    def atom():
        nonlocal pos
        if pos >= len(tokens):
            raise ValueError("incomplete conditional expression")
        t = tokens[pos]
        pos += 1
        if t == "not":
            return not atom()
        if t == "(":
            result = disjunction()
            if pos >= len(tokens) or tokens[pos] != ")":
                raise ValueError("unbalanced conditional expression")
            pos += 1
            return result
        if t in env:
            return env[t]
        raise ValueError(f"unsupported conditional token {t!r}")

    def conjunction():
        nonlocal pos
        value = atom()
        while pos < len(tokens) and tokens[pos] == "and":
            pos += 1
            right = atom()
            value = value and right
        return value

    def disjunction():
        nonlocal pos
        value = conjunction()
        while pos < len(tokens) and tokens[pos] in ("or", "xor"):
            op = tokens[pos]
            pos += 1
            right = conjunction()
            value = (value != right) if op == "xor" else (value or right)
        return value

    result = disjunction()
    if pos != len(tokens):
        raise ValueError("unsupported conditional expression")
    return result


def lines(text: str, vba7: bool, win64: bool):
    env = dict(vba7=vba7, win64=win64, win32=True, mac=False,
               true=True, false=False)
    stack, active, pending, first = [], True, "", 0
    for number, raw in enumerate(text.splitlines(), 1):
        code = mask(raw).strip()
        if code.startswith("#"):
            m = re.fullmatch(r"#(If|ElseIf)\s+(.+?)\s+Then", code, re.I)
            if m:
                value = conditional(m[2], env)
                if m[1].lower() == "if":
                    stack.append([active, value])
                    active = active and value
                else:
                    if not stack:
                        raise ValueError(f"line {number}: orphan #ElseIf")
                    parent, taken = stack[-1]
                    active = parent and not taken and value
                    stack[-1][1] = taken or value
            elif re.fullmatch(r"#Else", code, re.I) and stack:
                parent, taken = stack[-1]
                active = parent and not taken
                stack[-1][1] = True
            elif re.fullmatch(r"#End\s+If", code, re.I) and stack:
                active = stack.pop()[0]
            else:
                raise ValueError(f"line {number}: unsupported conditional directive {code}")
            continue
        if not active or not code:
            continue
        if not pending:
            first = number
        if re.search(r"\s_\s*$", code):
            pending += re.sub(r"\s_\s*$", " ", code)
            continue
        yield first, pending + code
        pending = ""
    if pending or stack:
        raise ValueError("unterminated continuation or conditional block")


def comma_parts(text: str):
    depth, start = 0, 0
    for i, c in enumerate(text):
        depth += (c == "(") - (c == ")")
        if c == "," and depth == 0:
            yield text[start:i].strip()
            start = i + 1
    yield text[start:].strip()


def declared(text: str) -> set[str]:
    result = set()
    for part in comma_parts(text):
        part = re.sub(r"^(?:(?:Optional|ByVal|ByRef|ParamArray|WithEvents|Const)\s+)+", "", part, flags=re.I)
        m = re.match(IDENT, part)
        if m:
            result.add(m[0].lower())
    return result


@dataclass
class Procedure:
    name: str
    kind: str
    locals: set[str] = field(default_factory=set)
    labels: set[str] = field(default_factory=set)
    handlers: list[tuple[int, str]] = field(default_factory=list)
    targets: list[tuple[int, str]] = field(default_factory=list)


def parse(text: str, vba7: bool, win64: bool, *, object_module: bool = False):
    module, exported, procedures = set(), set(), []
    current, block = None, None
    for number, logical in lines(text, vba7, win64):
        # := is a named argument, never a statement separator.
        parts = re.split(r":(?!=)", logical)
        for index, statement in enumerate(parts):
            s = statement.strip()
            if not s:
                continue
            if current and index < len(parts) - 1 and re.fullmatch(IDENT + r"|\d+", s):
                current.labels.add(s.lower())
                continue
            m = PROC.match(s)
            if m:
                if current:
                    raise ValueError(f"line {number}: nested procedure")
                kind = m[1].lower()
                current = Procedure(m[2], kind)
                # Property Let/Set names are assignable members of their own
                # module. Keep them module-scoped: object members must never
                # become project-wide declarations merely because they are
                # Public in a class or UserForm export.
                if kind in ("property let", "property set"):
                    module.add(m[2].lower())
                signature = m[3]
                if signature.startswith("("):
                    end = signature.rfind(")")
                    if end < 0:
                        raise ValueError(f"line {number}: incomplete signature")
                    current.locals |= declared(signature[1:end])
                procedures.append(current)
                continue
            if re.fullmatch(r"End\s+(Sub|Function|Property)", s, re.I):
                current = None
                continue
            if not current:
                m = re.match(r"(?:(Public|Private)\s+)?(Type|Enum)\b", s, re.I)
                if m:
                    block = (m[2].lower(), (m[1] or "public").lower())
                    continue
                if re.match(r"End\s+(Type|Enum)\b", s, re.I):
                    block = None
                    continue
                if block:
                    if block[0] == "enum":
                        names = declared(s)
                        module |= names
                        if block[1] == "public":
                            exported |= names
                    continue
            m = re.match(r"(Dim|Static|Public|Private|Global|Const)\s+(.+)", s, re.I)
            if m and not re.match(r"(?:Declare|Type|Enum|Event)\b", m[2], re.I):
                names = declared(m[2])
                if current:
                    current.locals |= names
                else:
                    module |= names
                    if not object_module and m[1].lower() in ("public", "global"):
                        exported |= names
                continue
            if not current:
                continue
            # Numbered lines may precede ordinary statements.
            m = re.match(r"(\d+)\s+(.*)", s)
            if m:
                current.labels.add(m[1])
                s = m[2]
            for h in re.finditer(r"\bOn\s+Error\s+GoTo\s+(-?\d+|" + IDENT + r")\b", s, re.I):
                if h[1] not in ("0", "-1"):
                    current.handlers.append((number, h[1]))
            # Single-line If/Else arms are independently executable statements.
            arms = re.split(r"\bThen\b|\bElse\b", s, flags=re.I)
            for arm in arms:
                arm = arm.strip()
                if re.match(r"(?:If|ElseIf)\b", arm, re.I):
                    continue
                arm = re.sub(r"^(?:Let|Set)\s+", "", arm, flags=re.I)
                loop = re.match(r"For\s+(?:Each\s+)?(" + IDENT + r")\b", arm, re.I)
                redim = re.match(r"ReDim\s+(?:Preserve\s+)?(.+)", arm, re.I)
                if loop:
                    current.targets.append((number, loop[1]))
                elif redim:
                    for part in comma_parts(redim[1]):
                        m = re.match(r"(" + IDENT + r")\s*\(", part)
                        if m:
                            current.targets.append((number, m[1]))
                else:
                    m = re.match(r"(" + IDENT + r")[$%&!#@]?\s*(.*)", arm)
                    if m:
                        tail = m[2]
                        if tail.startswith("("):
                            depth, end = 0, None
                            for i, c in enumerate(tail):
                                depth += (c == "(") - (c == ")")
                                if depth == 0:
                                    end = i
                                    break
                            tail = tail[end + 1:].lstrip() if end is not None else ""
                        if tail.startswith("="):
                            current.targets.append((number, m[1]))
    if current:
        raise ValueError("unterminated procedure")
    return module, exported, procedures


def analyse(sources: dict[str, str]) -> tuple[list[str], list[str]]:
    handlers, assignments = set(), set()
    for profile, vba7, win64 in PROFILES:
        parsed, public = {}, set()
        for path, text in sources.items():
            try:
                object_module = path.lower().endswith((".cls", ".frm"))
                parsed[path] = parse(text, vba7, win64, object_module=object_module)
                public |= parsed[path][1]
            except ValueError as exc:
                problem = f"{path}: [{profile}] unsupported source: {exc}"
                handlers.add(problem)
                assignments.add(problem)
        for path, (module, _, procedures) in parsed.items():
            for proc in procedures:
                allowed = public | module | proc.locals
                if proc.kind in ("function", "property get"):
                    allowed |= {proc.name.lower()}
                for number, label in proc.handlers:
                    if label.lower() not in proc.labels:
                        handlers.add(f"{path}:{number}: {proc.name}: [{profile}] missing error-handler label '{label}'")
                for number, name in proc.targets:
                    if name.lower() not in allowed:
                        assignments.add(f"{path}:{number}: {proc.name}: [{profile}] undeclared assignment '{name}'")
    return sorted(handlers), sorted(assignments)
