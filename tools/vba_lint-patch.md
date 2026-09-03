# `vba_lint.py` — two insertions for #51

Nothing else in the file changes. Active gate goes from 12 checks to 13.

## 1. New check function

Insert immediately **before** `def write_json(` (or anywhere among the other
`check_*` functions; placement is cosmetic):

```python
def check_release_provenance_fixtures(rep: Report) -> None:
    """Delegate the strict provenance matrix to its own module.

    The matrix asserts process exit codes against a copy of the release tool in
    throwaway repositories, so it cannot be an inline fixture like the changelog
    ones. Running it from here keeps it enforced by the same gate rather than by
    whoever remembers to run it.
    """
    harness = ROOT / "tools" / "test_release_provenance.py"
    if not harness.exists():
        rep.check("release provenance strict fixtures",
                  [f"harness not found: {harness}"])
        return

    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "provenance-fixtures.json"
        proc = subprocess.run(
            [sys.executable, str(harness), "--json", str(out)],
            cwd=str(ROOT), capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
        problems: list[str] = []
        if out.exists():
            problems = json.loads(out.read_text(encoding="utf-8")).get("failures", [])
        # A crash before the result document is written must still fail the gate.
        if proc.returncode != 0 and not problems:
            problems = [f"harness exited {proc.returncode}"]
            tail = (proc.stderr or proc.stdout).strip().splitlines()[-5:]
            problems.extend(f"  {line}" for line in tail)

    rep.check("release provenance strict fixtures", problems)
```

## 2. Register it in `main()`

Add one line after the existing final check:

```python
    check_changelog_released_sections_frozen(rep)
    check_release_provenance_fixtures(rep)          # <-- new
```

## 3. Imports

`json`, `subprocess`, `sys` and `Path` are already imported. Add if absent:

```python
import tempfile
```
