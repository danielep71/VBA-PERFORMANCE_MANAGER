"""The static gate's checks, in the order tools/vba_lint.py runs them (#89).

Each entry pairs the name the gate reports with the summary the README lists.
vba_lint.py fails when the checks it ran differ from this list, so the list
cannot drift from the gate. tools/readme_facts.py renders the README's
static-check count and numbered list from it.
"""
from __future__ import annotations

CHECKS: tuple[tuple[str, str], ...] = (
    ("no merge conflict markers", "no merge-conflict markers"),
    ("procedure blocks balanced", "balanced procedure blocks"),
    ("block pairs balanced", "balanced control blocks"),
    ("no reserved words as identifiers", "no reserved words used as identifiers"),
    ("error sources name the right module", "error sources name the correct module"),
    ("no bare error numbers (shipped source)",
     "no bare error numbers in shipped production source"),
    ("no undefined local callees", "no undefined local callees"),
    ("every test is defined and called", "every test is defined and called"),
    ("TotalSteps matches case count", "`TotalSteps` matches the executed case count"),
    ("version stamps agree", "version stamps agree"),
    ("native APIs have a single call site", "native APIs have a single call site"),
    ("released changelog sections frozen",
     "released CHANGELOG sections remain frozen against their tags"),
    ("release provenance strict fixtures", "strict release-provenance fixtures"),
    ("tracked VBA components are inventoried", "tracked VBA components are inventoried"),
    ("error-handler labels resolve within their procedure",
     "error-handler labels resolve within their procedure"),
    ("assigned identifiers are declared", "assigned identifiers are declared"),
    ("workflow actions use immutable pins", "workflow actions use immutable pins"),
    ("tracked binaries are reviewed", "tracked binaries are reviewed"),
    ("Markdown links and anchors resolve", "Markdown links and anchors resolve"),
    ("README facts match their sources", "README facts match their sources"),
)

NAMES: tuple[str, ...] = tuple(name for name, _ in CHECKS)
