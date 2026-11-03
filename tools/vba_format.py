"""Check or apply the layout of the tracked VBA exports (#83).

The formatter applies only whitespace rules, measured from this repository's
own source, that cannot change what VBA compiles:

1. CRLF line ends, one final line end, and no blank lines at the end;
2. no trailing white space;
3. four spaces for each tab in a line's indentation;
4. procedure declarations, their End statements and line labels start in
   column 0 (continuation lines are left alone);
5. one or two blank lines between a procedure's End statement and what follows:
   none becomes one, more than two become two.

It never changes a string literal or a comment's text, and it keeps the export
header the VBE writes (VERSION, BEGIN ... END, Object and Attribute lines) byte
for byte. After --write, re-import the changed components into the VBE.

    python3 tools/vba_format.py --check   # fail on any difference, write nothing
    python3 tools/vba_format.py --write   # rewrite the files that differ

Without paths it covers every tracked export under src/, test/ and demo/. The
static gate runs the same check as "VBA source layout is formatted".
"""
from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from pathlib import Path

from vba_inventory import ROOTS, tracked_vba, under
from vba_source_rules import PROC, PROC_END, code_part, header_length, physical_lines

ROOT = Path(__file__).resolve().parent.parent
CHECK = "VBA source layout is formatted"
INDENT = "    "
LABEL = re.compile(r"^[A-Za-z_][A-Za-z_0-9]*:$")
MAX_BLANK_AFTER_END = 2


def _outer(code: str) -> bool:
    """Whether a statement belongs in column 0."""
    return bool(PROC.match(code) or PROC_END.match(code) or LABEL.match(code))


def format_body(physical: Sequence[str]) -> list[str]:
    out: list[str] = []
    continued = False
    after_end: int | None = None      # blank lines seen since a procedure's End
    for raw in physical:
        body = raw.rstrip(" \t").lstrip(" \t")
        indent = raw[:len(raw) - len(raw.lstrip(" \t"))].replace("\t", INDENT)
        line = indent + body if body else ""
        if not line:
            if after_end is None:
                out.append("")
            else:
                after_end += 1
            continue
        if after_end is not None:
            out.extend([""] * min(max(after_end, 1), MAX_BLANK_AFTER_END))
            after_end = None
        code = code_part(line)[0].strip()
        if not continued and _outer(code):
            line = line.lstrip()
            if PROC_END.match(code):
                after_end = 0
        out.append(line)
        continued = code.endswith(" _") or code == "_"
    while out and not out[-1]:
        out.pop()
    return out


def format_text(text: str) -> str:
    """The formatted form of one export; idempotent."""
    physical = physical_lines(text.replace("\r\n", "\n").replace("\r", "\n"))
    split = header_length(physical)
    return "".join(f"{line}\r\n" for line in [*physical[:split], *format_body(physical[split:])])


def read(path: Path) -> str:
    # Latin-1 maps every byte to one character, so any byte round-trips.
    return path.read_bytes().decode("latin-1")


def repository_files(root: Path = ROOT) -> list[str]:
    return under(tracked_vba(root), *ROOTS)


def problems(root: Path = ROOT, paths: Sequence[str] | None = None) -> list[str]:
    """One line per export whose layout differs from its formatted form."""
    found = []
    for rel in paths if paths is not None else repository_files(root):
        text = read(root / rel)
        if format_text(text) != text:
            found.append(f"{rel}: layout differs; run python3 tools/vba_format.py --write "
                         "and re-import the component")
    return found


def fixture_problems() -> list[str]:
    """The formatter must fix each rule, keep a literal, and be idempotent."""
    messy = ("Attribute VB_Name = \"M_Fixture\"  \r\n"
             "    Option Explicit \r\n"
             "\tPrivate Const Spaced As String = \"a  \"\r\n"
             "  Public Sub First()\r\n"
             "  Done:\r\n"
             "  End Sub\r\n"
             "Public Sub Second( _\r\n"
             "    ByVal value As Long)\r\n"
             "End Sub\r\n\r\n\r\n\r\n\r\n")
    clean = ("Attribute VB_Name = \"M_Fixture\"  \r\n"
             "    Option Explicit\r\n"
             "    Private Const Spaced As String = \"a  \"\r\n"
             "Public Sub First()\r\n"
             "Done:\r\n"
             "End Sub\r\n"
             "\r\n"
             "Public Sub Second( _\r\n"
             "    ByVal value As Long)\r\n"
             "End Sub\r\n")
    found = []
    if format_text(messy) != clean:
        found.append("fixture: the messy module did not format to the expected layout")
    if format_text(clean) != clean:
        found.append("fixture: formatting is not idempotent")
    if format_text(clean.replace("\r\n", "\n")) != clean:
        found.append("fixture: LF line ends were not converted to CRLF")
    return found


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true", help="fail on any difference; write nothing")
    mode.add_argument("--write", action="store_true", help="rewrite the files that differ")
    parser.add_argument("paths", nargs="*", help="exports to format (default: all tracked)")
    args = parser.parse_args(argv)
    paths = args.paths or repository_files()
    differing = problems(ROOT, paths)
    if args.write:
        for line in differing:
            rel = line.split(":", 1)[0]
            (ROOT / rel).write_bytes(format_text(read(ROOT / rel)).encode("latin-1"))
            print(f"formatted {rel}")
        print(f"{len(differing)} of {len(paths)} file(s) rewritten")
        return 0
    print("\n".join(differing) or f"{len(paths)} file(s) formatted")
    return 1 if differing else 0


if __name__ == "__main__":
    sys.exit(main())
