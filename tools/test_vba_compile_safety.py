"""Local deterministic fixtures for #45; also enforced by the ordinary gate."""
from __future__ import annotations

from vba_compile_safety import analyse


def procedure(body: str, header: str = "Public Sub Test()") -> str:
    return header + "\n" + body + "\nEnd Sub\n"


# Each expectation is a list of diagnostic substrings. An empty list requires
# that check to pass. Thus every single-defect case proves check independence.
CASES = [
    ("missing label", {"a.bas": procedure("On Error GoTo Lost")}, ["a.bas:2: Test:", "'Lost'"], []),
    ("local label", {"a.bas": procedure("On Error GoTo Clean\nClean:\nExit Sub")}, [], []),
    ("other procedure label", {"a.bas": procedure("On Error GoTo Clean") + procedure("Clean:", "Sub Other()")}, ["missing error-handler"], []),
    ("other module label", {"a.bas": procedure("On Error GoTo Clean"), "b.bas": procedure("Clean:")}, ["missing error-handler"], []),
    ("special error forms", {"a.bas": procedure("On Error GoTo 0\nOn Error GoTo -1\nOn Error Resume Next")}, [], []),
    ("numeric label", {"a.bas": procedure("On Error GoTo 100\n100 Exit Sub")}, [], []),
    ("case insensitive label", {"a.bas": procedure("On Error GoTo CLEAN\nclean:")}, [], []),
    ("missing assignment", {"a.bas": procedure("Missing = 1")}, [], ["a.bas:2: Test:", "'Missing'"]),
    ("local declarations", {"a.bas": procedure("Dim A As Long, B As Long\nStatic C As Long\nA=1: B=2: C=3")}, [], []),
    ("multiline declarations", {"a.bas": procedure("Dim A As Long, _\n B As Long\nA=1: B=2")}, [], []),
    ("multiline parameters", {"a.bas": procedure("A=1: B=2", "Private Static Sub Test( _\n ByVal A As Long, _\n Optional ByRef B As Long = 0)")}, [], []),
    ("array parameter commas", {"a.bas": procedure("A(1)=2", "Sub Test(ByRef A() As Long, Optional B As Long = 0)")}, [], []),
    ("function result", {"a.bas": "Function F() As Long\nF=1\nEnd Function"}, [], []),
    ("property get result", {"a.cls": "Public Property Get Value() As Long\nValue=1\nEnd Property"}, [], []),
    ("other function result", {"a.bas": "Function F() As Long\nG=1\nEnd Function\nFunction G() As Long\nG=1\nEnd Function"}, [], ["'G'"]),
    ("sub name not result", {"a.bas": procedure("Test=1")}, [], ["'Test'"]),
    ("qualified targets", {"a.bas": procedure("Obj.Value=1\nObj.Items(1)=2\nWith Obj\n.Value=3\n!Field=4\nEnd With")}, [], []),
    ("let set", {"a.bas": procedure("Dim A As Long, B As Object\nLet A=1\nSet B=Nothing")}, [], []),
    ("undeclared set", {"a.bas": procedure("Set Missing=Nothing")}, [], ["'Missing'"]),
    ("declared arrays", {"a.bas": procedure("Dim A() As Double, B() As Long\nReDim A(1 To 2), B(1 To 3)\nReDim Preserve A(1 To 4)\nA(1)=2")}, [], []),
    ("undeclared array", {"a.bas": procedure("Missing(1, F(2))=3")}, [], ["'Missing'"]),
    ("undeclared redim", {"a.bas": procedure("ReDim Preserve Missing(1 To 4)")}, [], ["'Missing'"]),
    ("loop variables", {"a.bas": procedure("Dim I As Long, Item As Variant\nFor I=1 To 3\nNext I\nFor Each Item In Items\nNext Item")}, [], []),
    ("undeclared for", {"a.bas": procedure("For Missing=1 To 3\nNext Missing")}, [], ["'Missing'"]),
    ("undeclared for each", {"a.bas": procedure("For Each Missing In Items\nNext Missing")}, [], ["'Missing'"]),
    ("inline branches", {"a.bas": procedure("Dim A As Long\nIf A=0 Then A=1 Else Missing=2")}, [], ["'Missing'"]),
    ("inline handler", {"a.bas": procedure("If True Then On Error GoTo Missing")}, ["'Missing'"], []),
    ("comments and escaped strings", {"a.bas": procedure("Dim S As String\nS = \"a'\"\"b: Then Missing=1 On Error GoTo Bad\" ' Missing=1\n' On Error GoTo Bad\nRem Missing=1: On Error GoTo Bad\nIf True Then Rem Missing=2")}, [], []),
    ("named arguments", {"a.bas": procedure("Call Work(Name:=1, Other:=2)")}, [], []),
    ("private module variable", {"a.bas": "Private A As Long\n" + procedure("A=1")}, [], []),
    ("public other module variable", {"a.bas": procedure("Shared=1"), "b.bas": "Public Shared As Long"}, [], []),
    ("private other module variable", {"a.bas": procedure("Hidden=1"), "b.bas": "Private Hidden As Long"}, [], ["'Hidden'"]),
    ("class member not global", {"a.bas": procedure("Value=1"), "b.cls": "VERSION 1.0 CLASS\nPublic Value As Long"}, [], ["'Value'"]),
    ("type field not local", {"a.bas": "Private Type Record\nField As Long\nEnd Type\n" + procedure("Field=1")}, [], ["'Field'"]),
    ("conditional declaration", {"a.bas": "#If VBA7 Then\nPrivate A As Long\n#Else\nPrivate A As Long\n#End If\n" + procedure("A=1")}, [], []),
    ("inactive declaration", {"a.bas": "#If Win64 Then\nPrivate A As Long\n#End If\n" + procedure("A=1")}, [], ["[VBA7-Win32]", "'A'"]),
    ("inactive handler", {"a.bas": procedure("On Error GoTo Clean\n#If Win64 Then\nClean:\n#End If")}, ["[VBA7-Win32]", "'Clean'"], []),
    ("conditional procedure signatures", {"a.bas": "#If VBA7 And Win64 Then\nSub Test(A As LongLong)\n#ElseIf VBA7 Then\nSub Test(A As Long)\n#Else\nSub Test(A As Long)\n#End If\nA=1\nEnd Sub"}, [], []),
    ("unknown conditional fails closed", {"a.bas": "#If UNKNOWN Then\n" + procedure("Missing=1") + "#End If"}, ["unsupported conditional"], ["unsupported conditional"]),
    ("label and assignment same line", {"a.bas": procedure("Dim A As Long\nOn Error GoTo Clean\nClean: A=1")}, [], []),
    ("enum declarations not executable", {"a.bas": "Public Enum Choice\nFirst=1\nSecond=2\nEnd Enum\n" + procedure("Dim A As Long\nA=First")}, [], []),
]


def fixture_problems() -> tuple[list[str], list[str]]:
    failures = ([], [])
    for name, sources, expected_h, expected_a in CASES:
        actual = analyse(sources)
        for index, expected in enumerate((expected_h, expected_a)):
            text = "\n".join(actual[index])
            if (not expected and actual[index]) or any(s not in text for s in expected):
                failures[index].append(f"fixture {name!r}: expected {expected!r}, got {actual[index]!r}")
    return failures


def integration_problems() -> list[str]:
    """Exercise the real CLI/JSON in an isolated copy, never the working tree."""
    import json
    import shutil
    import subprocess
    import sys
    import tempfile
    from pathlib import Path
    from vba_compile_safety import LABEL_CHECK, ASSIGN_CHECK

    failures = []
    root = Path(__file__).resolve().parent.parent
    with tempfile.TemporaryDirectory(prefix="cpm-compile-fixtures-") as temp:
        dest = Path(temp) / "repo"
        shutil.copytree(root, dest, ignore=shutil.ignore_patterns(".git", "__pycache__"))
        # The copied tree deliberately has no Git identity: this verifies CLI
        # wiring, not release-history or exact-SHA evidence.
        path = dest / "demo" / "CompileSafetyFixture.bas"
        for name, body, expected in (
            ("valid", "Dim A As Long\nA=1", None),
            ("handler", "On Error GoTo Missing", LABEL_CHECK),
            ("assignment", "Missing=1", ASSIGN_CHECK),
        ):
            path.write_text(procedure(body), encoding="utf-8")
            output = dest / "result.json"
            run = subprocess.run([sys.executable, str(dest / "tools/vba_lint.py"),
                                  "--json", str(output)], capture_output=True,
                                 text=True, timeout=120, cwd=dest)
            if not output.exists():
                failures.append(f"CLI {name}: no JSON: {run.stderr}")
                continue
            result = json.loads(output.read_text(encoding="utf-8"))
            failed = [r["check"] for r in result["results"] if not r["passed"]]
            if run.returncode != int(expected is not None) or failed != ([expected] if expected else []):
                failures.append(f"CLI {name}: exit={run.returncode}, failed={failed}")
            names = [r["check"] for r in result["results"]]
            if any(names.count(check) != 1 for check in (LABEL_CHECK, ASSIGN_CHECK)):
                failures.append(f"CLI {name}: both checks must be reported exactly once")
            if result["checks_run"] != len(names):
                failures.append(f"CLI {name}: reported check count disagrees with results")
    return failures


if __name__ == "__main__":
    import sys
    failures = fixture_problems()
    for problem in failures[0] + failures[1]:
        print(problem)
    print(f"{len(CASES)} fixtures; {len(failures[0]) + len(failures[1])} failures")
    integration = integration_problems() if "--integration" in sys.argv else []
    for problem in integration:
        print(problem)
    if "--integration" in sys.argv:
        print(f"3 CLI/JSON integration scenarios; {len(integration)} failures")
    raise SystemExit(bool(failures[0] or failures[1] or integration))
