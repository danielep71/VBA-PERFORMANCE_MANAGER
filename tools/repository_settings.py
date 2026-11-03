"""Preview, apply and verify the repository settings baseline (#86).

tools/repository-settings.json records the settings this repository relies
on: general options and merge methods, the security features, and the
rulesets with their required checks. This script compares the live
repository with that baseline through the GitHub CLI (``gh api``), using the
maintainer's own authentication.

- ``--preview`` (the default) lists every difference and the API calls that
  ``--apply`` would make, and changes nothing;
- ``--apply`` makes those calls, then verifies. GitHub enables some security
  features asynchronously, so the verify re-reads the settings a bounded
  number of times, with a pause between attempts, before it reports a
  difference;
- ``--verify`` changes nothing, reads once and exits 1 on any difference.

Dependabot security updates go through their own endpoint,
``/automated-security-fixes``, after vulnerability alerts (#103).

``--evidence DIR`` writes a dated JSON record of the run, with the live
snapshot, the differences and the result, for the maintainer to keep. A
setting the token cannot read is a difference, never a pass.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import time
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
BASELINE = ROOT / "tools" / "repository-settings.json"
RULESET_KEYS = ("name", "target", "enforcement", "conditions", "bypass_actors", "rules")
# Settings the general PATCH carries in security_and_analysis.
ANALYSIS = ("secret_scanning", "secret_scanning_push_protection")
DEPENDABOT = "dependabot_security_updates"
ATTEMPTS = 5          # reads after an apply before a difference is reported
PAUSE = 3.0           # seconds between those reads
UNREADABLE = "unreadable"

# call(method, endpoint, payload) -> (HTTP status, parsed body or None)
Call = Callable[[str, str, "Any"], "tuple[int, Any]"]


def gh_call(method: str, endpoint: str, payload: Any = None) -> tuple[int, Any]:
    """Call the REST API through the authenticated GitHub CLI."""
    command = ["gh", "api", "--include", "--method", method,
               "-H", "Accept: application/vnd.github+json", endpoint]
    if payload is not None:
        command += ["--input", "-"]
    run = subprocess.run(command, input=None if payload is None else json.dumps(payload),
                         capture_output=True, text=True)
    head, _, body = run.stdout.replace("\r\n", "\n").partition("\n\n")
    try:
        status = int(head.split()[1])
    except (IndexError, ValueError):
        status = 0
    try:
        parsed = json.loads(body) if body.strip() else None
    except json.JSONDecodeError:
        parsed = body
    return status, parsed


def load_baseline(text: str) -> tuple[dict[str, Any], list[str]]:
    try:
        data: Any = json.loads(text)
    except json.JSONDecodeError as exc:
        return {}, [f"repository-settings.json: invalid JSON: {exc}"]
    problems = []
    if not isinstance(data, dict) or data.get("schema_version") != 1:
        return {}, ['repository-settings.json: expected "schema_version": 1']
    for key, kind in (("repository", str), ("general", dict), ("security", dict), ("rulesets", list)):
        if not isinstance(data.get(key), kind):
            problems.append(f"repository-settings.json: {key} must be a {kind.__name__}")
    for ruleset in data.get("rulesets") or []:
        if not isinstance(ruleset, dict) or set(ruleset) != set(RULESET_KEYS):
            problems.append("repository-settings.json: each ruleset needs exactly "
                            + ", ".join(RULESET_KEYS))
    return data, problems


def _normal_rules(rules: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(({"type": r["type"], **({"parameters": r["parameters"]} if r.get("parameters")
                                          else {})} for r in rules), key=lambda r: str(r["type"]))


def snapshot(call: Call, repo: str) -> dict[str, Any]:
    """Read the live settings the baseline covers."""
    status, info = call("GET", f"repos/{repo}", None)
    if status != 200 or not isinstance(info, dict):
        raise RuntimeError(f"cannot read repos/{repo}: HTTP {status}")
    live: dict[str, Any] = {"general": info, "security": {}, "rulesets": {}}
    analysis = info.get("security_and_analysis") or {}
    for key in ANALYSIS:
        state = (analysis.get(key) or {}).get("status")
        live["security"][key] = UNREADABLE if state is None else state == "enabled"
    # Dependabot security updates: the dedicated endpoint, then the summary in
    # security_and_analysis, else unreadable (#103).
    status, fixes = call("GET", f"repos/{repo}/automated-security-fixes", None)
    if status == 200 and isinstance(fixes, dict) and "enabled" in fixes:
        live["security"][DEPENDABOT] = bool(fixes["enabled"])
    else:
        state = (analysis.get(DEPENDABOT) or {}).get("status")
        live["security"][DEPENDABOT] = UNREADABLE if state is None else state == "enabled"
    status, _ = call("GET", f"repos/{repo}/vulnerability-alerts", None)
    live["security"]["vulnerability_alerts"] = {204: True, 404: False}.get(status, UNREADABLE)
    status, body = call("GET", f"repos/{repo}/private-vulnerability-reporting", None)
    live["security"]["private_vulnerability_reporting"] = (
        bool(body.get("enabled")) if status == 200 and isinstance(body, dict) else UNREADABLE)
    status, listing = call("GET", f"repos/{repo}/rulesets", None)
    if status != 200 or not isinstance(listing, list):
        raise RuntimeError(f"cannot list rulesets: HTTP {status}")
    for item in listing:
        status, ruleset = call("GET", f"repos/{repo}/rulesets/{item['id']}", None)
        if status == 200 and isinstance(ruleset, dict):
            live["rulesets"][ruleset["name"]] = {"id": ruleset["id"],
                                                  **{k: ruleset.get(k) for k in RULESET_KEYS}}
    return live


def differences(baseline: dict[str, Any], live: dict[str, Any]) -> list[tuple[str, Any, Any]]:
    """(setting, expected, actual) for every setting that does not match."""
    found: list[tuple[str, Any, Any]] = []
    for key, expected in baseline["general"].items():
        if live["general"].get(key) != expected:
            found.append((f"general.{key}", expected, live["general"].get(key)))
    for key, expected in baseline["security"].items():
        if live["security"].get(key) != expected:
            found.append((f"security.{key}", expected, live["security"].get(key)))
    for ruleset in baseline["rulesets"]:
        current = live["rulesets"].get(ruleset["name"])
        if current is None:
            found.append((f"ruleset {ruleset['name']}", "present", "missing"))
            continue
        for key in RULESET_KEYS:
            want, have = ruleset[key], current.get(key)
            if key == "rules":
                want, have = _normal_rules(want), _normal_rules(have or [])
            if want != have:
                found.append((f"ruleset {ruleset['name']}.{key}", want, have))
    return found


def planned_calls(baseline: dict[str, Any], live: dict[str, Any],
                  found: list[tuple[str, Any, Any]]) -> list[tuple[str, str, Any]]:
    """The API calls that bring the live settings to the baseline, in a safe order."""
    repo = baseline["repository"]
    changed = {setting for setting, _, _ in found}
    calls: list[tuple[str, str, Any]] = []
    security = baseline["security"]
    # Dependabot security updates need vulnerability alerts: enable alerts
    # first, and disable the updates first.
    alerts: list[tuple[str, str, Any]] = [
        ("PUT" if security.get("vulnerability_alerts") else "DELETE",
         f"repos/{repo}/vulnerability-alerts", None)
    ] if "security.vulnerability_alerts" in changed else []
    fixes: list[tuple[str, str, Any]] = [
        ("PUT" if security.get(DEPENDABOT) else "DELETE",
         f"repos/{repo}/automated-security-fixes", None)
    ] if f"security.{DEPENDABOT}" in changed else []
    calls += alerts + fixes if security.get(DEPENDABOT) else fixes + alerts
    general = {key: value for key, value in baseline["general"].items()
               if f"general.{key}" in changed and key != "default_branch"}
    analysis = {key: {"status": "enabled" if security[key] else "disabled"}
                for key in ANALYSIS if f"security.{key}" in changed and key in security}
    if general or analysis:
        payload: dict[str, Any] = dict(general)
        if analysis:
            payload["security_and_analysis"] = analysis
        calls.append(("PATCH", f"repos/{repo}", payload))
    if "security.private_vulnerability_reporting" in changed:
        method = "PUT" if security["private_vulnerability_reporting"] else "DELETE"
        calls.append((method, f"repos/{repo}/private-vulnerability-reporting", None))
    for ruleset in baseline["rulesets"]:
        if any(s.startswith(f"ruleset {ruleset['name']}") for s in changed):
            current = live["rulesets"].get(ruleset["name"])
            if current is None:
                calls.append(("POST", f"repos/{repo}/rulesets", ruleset))
            else:
                calls.append(("PUT", f"repos/{repo}/rulesets/{current['id']}", ruleset))
    return calls


def write_evidence(directory: Path, record: dict[str, Any]) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    stamp = record["checked_at"].replace(":", "").replace("-", "")
    path = directory / f"repository-settings-{stamp}.json"
    path.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return path


def run(mode: str, call: Call = gh_call, baseline_text: str | None = None,
        evidence: Path | None = None, now: datetime | None = None,
        attempts: int = ATTEMPTS, pause: float = PAUSE,
        wait: Callable[[float], None] = time.sleep) -> tuple[int, list[str]]:
    """Run one mode; return (exit status, output lines)."""
    text = baseline_text if baseline_text is not None else BASELINE.read_text(encoding="utf-8")
    baseline, problems = load_baseline(text)
    if problems:
        return 1, problems
    repo = baseline["repository"]
    out: list[str] = []
    live = snapshot(call, repo)
    found = differences(baseline, live)
    calls = planned_calls(baseline, live, found)
    for setting, expected, actual in found:
        out.append(f"differs: {setting}: expected {json.dumps(expected)}, live {json.dumps(actual)}")
    if mode == "preview":
        out += [f"would call: {method} {endpoint}" for method, endpoint, _ in calls]
    if mode == "apply":
        for method, endpoint, payload in calls:
            status, body = call(method, endpoint, payload)
            out.append(f"called: {method} {endpoint} -> HTTP {status}")
            if status >= 300:
                out.append(f"  failed: {body}")
        for attempt in range(1, attempts + 1):
            live = snapshot(call, repo)
            found = differences(baseline, live)
            if not found or attempt == attempts:
                break
            out.append(f"re-read {attempt} of {attempts}: {len(found)} setting(s) still "
                       f"differ; waiting {pause:g} s")
            wait(pause)
        out.append(f"verified with {attempt} read(s)")
        out += [f"still differs: {s}: expected {json.dumps(e)}, live {json.dumps(a)}"
                for s, e, a in found]
    status = 1 if (found and mode in ("apply", "verify")) else 0
    out.append("settings match the baseline" if not found else
               f"{len(found)} setting(s) differ from the baseline")
    if evidence is not None:
        moment = (now or datetime.now(timezone.utc)).strftime("%Y-%m-%dT%H:%M:%SZ")
        record = {
            "checked_at": moment, "mode": mode, "repository": repo,
            "baseline_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
            "differences": [{"setting": s, "expected": e, "live": a} for s, e, a in found],
            "result": "match" if not found else "differs",
            "live": {"general": {k: live["general"].get(k) for k in baseline["general"]},
                     "security": live["security"], "rulesets": live["rulesets"]},
        }
        out.append(f"evidence written to {write_evidence(evidence, record)}")
    return status, out


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--preview", action="store_true", help="show differences (default)")
    mode.add_argument("--apply", action="store_true", help="apply the baseline, then verify")
    mode.add_argument("--verify", action="store_true", help="exit 1 on any difference")
    parser.add_argument("--evidence", type=Path, metavar="DIR", help="write a dated JSON record")
    args = parser.parse_args(argv)
    selected = "apply" if args.apply else "verify" if args.verify else "preview"
    try:
        status, out = run(selected, evidence=args.evidence)
    except RuntimeError as exc:
        print(f"error: {exc}")
        return 1
    print("\n".join(out))
    return status


if __name__ == "__main__":
    sys.exit(main())
