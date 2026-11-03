"""Keep the repository's issue labels in tools/labels.json (#86).

The file holds every label's name, colour and description. Two modes run
against the live repository through the GitHub REST API, with the token in
GITHUB_TOKEN (or GH_TOKEN) and the repository in GITHUB_REPOSITORY:

- ``--apply`` (the labels-sync workflow, on the default branch) creates every
  missing label and updates any whose colour or description differs. A label
  that exists only on GitHub is reported, never deleted: removing a label
  strips it from every issue, so that stays a deliberate manual step.
- ``--check`` (the labels-drift workflow) changes nothing and fails on any
  difference, including labels that exist only on GitHub.

Without a mode the file itself is validated, which needs no network.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.parse
import urllib.request
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
LABELS_FILE = ROOT / "tools" / "labels.json"
COLOR = re.compile(r"[0-9a-f]{6}")
MAX_DESCRIPTION = 100   # GitHub's limit

# send(method, path, payload) -> parsed JSON; the default talks to the REST API.
Send = Callable[[str, str, "dict[str, Any] | None"], Any]


@dataclass(frozen=True)
class Label:
    name: str
    color: str
    description: str


def parse(text: str) -> tuple[list[Label], list[str]]:
    """Return the declared labels and every problem in the file itself."""
    try:
        data: Any = json.loads(text)
    except json.JSONDecodeError as exc:
        return [], [f"labels.json: invalid JSON: {exc}"]
    if not isinstance(data, dict) or data.get("schema_version") != 1 \
            or not isinstance(data.get("labels"), list):
        return [], ['labels.json: expected {"schema_version": 1, "labels": [...]}']
    labels: list[Label] = []
    problems: list[str] = []
    seen: set[str] = set()
    for index, entry in enumerate(data["labels"]):
        if not isinstance(entry, dict) or set(entry) != {"name", "color", "description"}:
            problems.append(f"labels.json: entry {index} needs exactly name, color and description")
            continue
        name, color, description = entry["name"], entry["color"], entry["description"]
        if not isinstance(name, str) or not name.strip():
            problems.append(f"labels.json: entry {index} has no name")
            continue
        if name.lower() in seen:
            problems.append(f"labels.json: {name} is declared twice")
        seen.add(name.lower())
        if not isinstance(color, str) or not COLOR.fullmatch(color):
            problems.append(f"labels.json: {name}: color must be six lowercase hex digits")
        if not isinstance(description, str) or not description.strip():
            problems.append(f"labels.json: {name}: a description is required")
        elif len(description) > MAX_DESCRIPTION:
            problems.append(f"labels.json: {name}: description exceeds {MAX_DESCRIPTION} characters")
        labels.append(Label(name, str(color), str(description)))
    return labels, problems


def differences(declared: list[Label], live: list[Label]) -> tuple[list[Label], list[Label], list[str]]:
    """Return (to create, to update, live-only names) to bring live to declared."""
    by_name = {label.name: label for label in live}
    create = [label for label in declared if label.name not in by_name]
    update = [label for label in declared
              if label.name in by_name and by_name[label.name] != label]
    declared_names = {label.name for label in declared}
    extra = sorted(label.name for label in live if label.name not in declared_names)
    return create, update, extra


def _rest(token: str) -> Send:
    def send(method: str, path: str, payload: dict[str, Any] | None) -> Any:
        request = urllib.request.Request(
            "https://api.github.com" + path, method=method,
            data=None if payload is None else json.dumps(payload).encode(),
            headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json",
                     "X-GitHub-Api-Version": "2022-11-28", "User-Agent": "labels-sync"})
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read()
        return json.loads(body) if body else None
    return send


def live_labels(send: Send, repo: str) -> list[Label]:
    labels: list[Label] = []
    page = 1
    while True:
        batch = send("GET", f"/repos/{repo}/labels?per_page=100&page={page}", None)
        labels += [Label(item["name"], item["color"].lower(), item.get("description") or "")
                   for item in batch]
        if len(batch) < 100:
            return labels
        page += 1


def apply(declared: list[Label], send: Send, repo: str) -> list[str]:
    """Create and update labels; return a report line per action."""
    create, update, extra = differences(declared, live_labels(send, repo))
    report: list[str] = []
    for label in create:
        send("POST", f"/repos/{repo}/labels",
             {"name": label.name, "color": label.color, "description": label.description})
        report.append(f"created {label.name}")
    for label in update:
        send("PATCH", f"/repos/{repo}/labels/{urllib.parse.quote(label.name, safe='')}",
             {"color": label.color, "description": label.description})
        report.append(f"updated {label.name}")
    report += [f"not declared, left in place: {name}" for name in extra]
    return report


def drift(declared: list[Label], send: Send, repo: str) -> list[str]:
    """Return every difference between the file and the live labels."""
    create, update, extra = differences(declared, live_labels(send, repo))
    return ([f"missing on GitHub: {label.name}" for label in create]
            + [f"differs on GitHub: {label.name}" for label in update]
            + [f"on GitHub but not in tools/labels.json: {name}" for name in extra])


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--apply", action="store_true", help="create and update the live labels")
    mode.add_argument("--check", action="store_true", help="fail on any live difference")
    args = parser.parse_args(argv)

    declared, problems = parse(LABELS_FILE.read_text(encoding="utf-8"))
    if problems:
        print("\n".join(problems))
        return 1
    if not (args.apply or args.check):
        print(f"{len(declared)} labels declared; the file is valid")
        return 0
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    repo = os.environ.get("GITHUB_REPOSITORY")
    if not token or not repo:
        print("GITHUB_TOKEN (or GH_TOKEN) and GITHUB_REPOSITORY are required")
        return 1
    if args.apply:
        print("\n".join(apply(declared, _rest(token), repo)) or "labels already match")
        return 0
    found = drift(declared, _rest(token), repo)
    print("\n".join(found) or "labels match tools/labels.json")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
