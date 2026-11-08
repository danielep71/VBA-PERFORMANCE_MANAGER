"""Audit open issues for the metadata issue forms cannot enforce (#86).

Read-only: it never changes an issue. Every open issue must have:

- danielep71 among its assignees;
- exactly one priority label: P1, P2 or P3;
- an open milestone;
- a target date (YYYY-MM-DD) on or before the milestone's due date, given
  either in the issue's Schedule section (``- Target date: 2026-11-27``, with
  an optional ``- Start date:`` that must not be later) or through the issue
  form's "Target date" field;
- a Dependencies section (a ``## Dependencies`` heading or the form field)
  whose lines are ``None``, an existing ``#number``, or ``External
  condition: ...``.

A failed run is the triage signal; each finding is resolved by hand. The
issue-triage workflow runs it on issue and milestone events and weekly, with
GITHUB_TOKEN; without a token it reads the public repository anonymously.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import urllib.request
from collections.abc import Callable
from typing import Any

ASSIGNEE = "danielep71"
PRIORITIES = ("P1", "P2", "P3")
PLANNED_DATE = re.compile(r"^[ \t]*[-*][ \t]*Target (?:end )?date:[ \t]*(.*?)[ \t]*$", re.I | re.M)
START_DATE = re.compile(r"^[ \t]*[-*][ \t]*Start date:[ \t]*(\d{4}-\d{2}-\d{2})[ \t]*$", re.I | re.M)
FORM_DATE = re.compile(r"^###[ \t]+Target (?:end )?date[ \t]*\n+[ \t]*(.*?)[ \t]*$", re.I | re.M)
HEADING = re.compile(r"^(#{2,3})[ \t]+(.+?)[ \t]*$", re.M)
DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
REFERENCE = re.compile(r"(?<![\w/])#(\d+)\b")
EXTERNAL = re.compile(r"^external condition:\s*\S", re.I)
NO_RESPONSE = "_no response_"   # what an issue form renders for an empty field

Get = Callable[[str], Any]


def normalize_line_endings(body: str) -> str:
    """Use LF for parsing browser-edited bodies without changing their content."""
    return body.replace("\r\n", "\n").replace("\r", "\n")


def target_date(body: str) -> tuple[str | None, str | None]:
    """Return (date text, problem). The planning line wins over the form field."""
    body = normalize_line_endings(body)
    for pattern in (PLANNED_DATE, FORM_DATE):
        match = pattern.search(body)
        if match and match.group(1).strip().lower() not in ("", NO_RESPONSE):
            value = match.group(1).strip()
            if not DATE.fullmatch(value):
                return None, f"target date {value!r} is not YYYY-MM-DD"
            try:
                dt.date.fromisoformat(value)
            except ValueError:
                return None, f"target date {value} is not a valid date"
            return value, None
    return None, "no target date"


def dependency_lines(body: str) -> list[str] | None:
    """The non-empty lines of the Dependencies section, or None when it is absent."""
    body = normalize_line_endings(body)
    headings = list(HEADING.finditer(body))
    for index, heading in enumerate(headings):
        if heading.group(2).strip().lower() == "dependencies":
            end = headings[index + 1].start() if index + 1 < len(headings) else len(body)
            text = body[heading.end():end]
            lines = [line.strip().lstrip("-*").strip() for line in text.splitlines()]
            return [line for line in lines if line and line.lower() != NO_RESPONSE]
    return None


def issue_problems(issue: dict[str, Any], known: set[int]) -> list[str]:
    """Every rule the issue breaks; known holds every issue and PR number."""
    problems: list[str] = []
    body = normalize_line_endings(issue.get("body") or "")
    if ASSIGNEE not in [a["login"] for a in issue.get("assignees") or []]:
        problems.append(f"not assigned to {ASSIGNEE}")
    priorities = [label["name"] for label in issue.get("labels") or []
                  if label["name"] in PRIORITIES]
    if len(priorities) != 1:
        problems.append(f"needs exactly one priority label, has {priorities or 'none'}")
    milestone = issue.get("milestone")
    if not milestone:
        problems.append("no milestone")
    elif milestone.get("state") != "open":
        problems.append(f"milestone {milestone.get('title')} is closed")
    date, problem = target_date(body)
    if problem:
        problems.append(problem)
    elif date:
        if milestone and milestone.get("due_on") and date > milestone["due_on"][:10]:
            problems.append(f"target date {date} is after the milestone due date "
                            f"{milestone['due_on'][:10]}")
        start = START_DATE.search(body)
        if start and start.group(1) > date:
            problems.append(f"start date {start.group(1)} is after the target date {date}")
    return problems + dependency_problems(body, known)


def dependency_problems(body: str, known: set[int]) -> list[str]:
    """Check the Dependencies section; known holds every issue and PR number."""
    lines = dependency_lines(body)
    if lines is None:
        return ["no Dependencies section"]
    if not lines:
        return ["the Dependencies section is empty; write None if there are none"]
    problems: list[str] = []
    for line in lines:
        if line.lower() == "none" or EXTERNAL.match(line):
            continue
        numbers = [int(n) for n in REFERENCE.findall(line)]
        if not numbers:
            problems.append(f"dependency line {line!r} names no #issue, None or "
                            "External condition")
        problems += [f"dependency #{n} does not exist" for n in numbers if n not in known]
    return problems


def audit(issues: list[dict[str, Any]]) -> list[str]:
    """Findings for every open issue in a list of issues and pull requests."""
    known = {item["number"] for item in issues}
    findings: list[str] = []
    for issue in sorted(issues, key=lambda item: item["number"]):
        if issue.get("pull_request") or issue.get("state") != "open":
            continue
        findings += [f"#{issue['number']}: {p}" for p in issue_problems(issue, known)]
    return findings


def _get(token: str | None) -> Get:
    def get(path: str) -> Any:
        headers = {"Accept": "application/vnd.github+json", "User-Agent": "issue-triage",
                   "X-GitHub-Api-Version": "2022-11-28"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        request = urllib.request.Request("https://api.github.com" + path, headers=headers)
        with urllib.request.urlopen(request, timeout=30) as response:
            return json.loads(response.read())
    return get


def fetch(get: Get, repo: str) -> list[dict[str, Any]]:
    """Every issue and pull request of the repository, open and closed."""
    items: list[dict[str, Any]] = []
    page = 1
    while True:
        batch = get(f"/repos/{repo}/issues?state=all&per_page=100&page={page}")
        items += batch
        if len(batch) < 100:
            return items
        page += 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--repo", default=os.environ.get("GITHUB_REPOSITORY"),
                        help="owner/name; defaults to GITHUB_REPOSITORY")
    args = parser.parse_args(argv)
    if not args.repo:
        print("--repo or GITHUB_REPOSITORY is required")
        return 1
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    findings = audit(fetch(_get(token), args.repo))
    print("\n".join(findings) or "every open issue passes triage")
    print(f"{len(findings)} finding(s)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
