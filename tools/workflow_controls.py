"""Hardening rules for the GitHub workflows, read with a real YAML parser.

tools/workflow_pins.py proves that every action reference is immutable. This
module checks the rest of the workflow contract (#80), which needs the
structure of each job and step and therefore a YAML parser:

- the workflow-level token is read-only, and only the jobs listed in
  WRITE_PERMISSIONS may write, each only what it is listed for;
- every job has a positive integer timeout-minutes;
- every job runs on a fixed runner image, never a moving ``*-latest`` label;
- every actions/checkout step sets ``persist-credentials: false``;
- every actions/setup-python step reads the one exact interpreter version in
  .python-version, except the minimum-Python job, which pins the documented
  floor exactly;
- that minimum-Python job exists, and the README states the same floor.

PyYAML is pinned with hashes in tools/requirements-ci.txt. A duplicate mapping
key is an error: PyYAML would keep the last value while GitHub may read
another, so a rule could pass on a value GitHub does not use.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import yaml

ROOT = Path(__file__).resolve().parent.parent
WORKFLOW_DIR = ".github/workflows"
PYTHON_VERSION_FILE = ".python-version"
README = "README.md"

# The only jobs that may write, and what each may write.
WRITE_PERMISSIONS = {
    ("labels-sync.yml", "sync"): {"issues"},
    ("pr-milestones.yml", "clear"): {"pull-requests"},
    ("scorecard.yml", "analysis"): {"id-token", "security-events"},
}
MINIMUM_JOB = ("static-checks.yml", "minimum-python")
FIXED_RUNNER = re.compile(r"(ubuntu|windows|macos)-\d+(\.\d+)?(-arm)?")
EXACT_VERSION = re.compile(r"\d+\.\d+\.\d+")
FLOOR_SENTENCE = re.compile(r"Python (\d+\.\d+) or later")


class _StrictLoader(yaml.SafeLoader):
    """A safe loader whose mappings fail on a repeated key."""


def _unique_mapping(loader: _StrictLoader, node: yaml.MappingNode, deep: bool = False) -> Any:
    loader.flatten_mapping(node)
    seen: set[Any] = set()
    for key_node, _ in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if key in seen:
            raise yaml.constructor.ConstructorError(
                "while constructing a mapping", node.start_mark,
                f"found duplicate key {key!r}", key_node.start_mark)
        seen.add(key)
    return loader.construct_mapping(node, deep=deep)


_StrictLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _unique_mapping)


def load(text: str) -> Any:
    """Parse one YAML document; a duplicate key raises yaml.YAMLError."""
    document = yaml.load(text, Loader=_StrictLoader)
    # YAML 1.1 reads the bare key "on" as True.
    if isinstance(document, dict) and True in document:
        document["on"] = document.pop(True)
    return document


def _writes(permissions: Any) -> set[str]:
    if permissions == "write-all":
        return {"write-all"}
    if isinstance(permissions, dict):
        return {scope for scope, level in permissions.items() if level == "write"}
    return set()


def _steps(job: dict[str, Any]) -> list[dict[str, Any]]:
    return [step for step in job.get("steps") or [] if isinstance(step, dict)]


def _action(step: dict[str, Any]) -> str:
    return str(step.get("uses", "")).split("@", 1)[0]


def _job_problems(name: str, job_id: str, job: Any) -> tuple[list[str], str | None]:
    """Check one job; also return the exact minimum Python it pins, if any."""
    where = f"{name}: job {job_id}"
    if not isinstance(job, dict):
        return [f"{where}: not a mapping"], None
    problems: list[str] = []
    minimum: str | None = None
    allowed = WRITE_PERMISSIONS.get((name, job_id), set())
    extra = _writes(job.get("permissions")) - allowed
    if extra:
        problems.append(f"{where}: may write {sorted(extra)}, which is not allowed")

    timeout = job.get("timeout-minutes")
    if not isinstance(timeout, int) or isinstance(timeout, bool) or timeout <= 0:
        problems.append(f"{where}: timeout-minutes must be a positive integer")

    runner = job.get("runs-on")
    if not isinstance(runner, str) or not FIXED_RUNNER.fullmatch(runner):
        problems.append(f"{where}: runs-on must be a fixed runner image such as "
                        f"ubuntu-24.04, got {runner!r}")

    for step in _steps(job):
        options = step.get("with") or {}
        if _action(step) == "actions/checkout" and options.get("persist-credentials") is not False:
            problems.append(f"{where}: actions/checkout must set persist-credentials: false")
        if _action(step) != "actions/setup-python":
            continue
        if (name, job_id) == MINIMUM_JOB:
            version = str(options.get("python-version", ""))
            if not EXACT_VERSION.fullmatch(version) or "python-version-file" in options:
                problems.append(f"{where}: must pin the minimum Python as an exact "
                                f"python-version, got {version!r}")
            else:
                minimum = version
        elif options.get("python-version-file") != PYTHON_VERSION_FILE \
                or "python-version" in options:
            problems.append(f"{where}: actions/setup-python must read "
                            f"python-version-file: {PYTHON_VERSION_FILE}")
    return problems, minimum


def check(workflows: dict[str, str], python_version: str | None,
          readme: str | None) -> list[str]:
    """Return every rule violation; an empty list means the workflows comply.

    workflows maps a workflow file name to its text. python_version is the
    content of .python-version and readme the README text, or None if absent.
    """
    problems: list[str] = []

    pinned = (python_version or "").strip()
    if not EXACT_VERSION.fullmatch(pinned):
        problems.append(f"{PYTHON_VERSION_FILE} must hold one exact version such as 3.13.18, "
                        f"got {pinned!r}")

    minimum: str | None = None
    for name, text in sorted(workflows.items()):
        try:
            document = load(text)
        except yaml.YAMLError as exc:
            problems.append(f"{name}: invalid YAML: {exc}".splitlines()[0])
            continue
        if not isinstance(document, dict) or not isinstance(document.get("jobs"), dict):
            problems.append(f"{name}: no jobs mapping")
            continue

        if "permissions" not in document:
            problems.append(f"{name}: workflow-level permissions are not explicit")
        elif _writes(document["permissions"]):
            problems.append(f"{name}: the workflow-level token may write "
                            f"{sorted(_writes(document['permissions']))}; grant writes per job")

        for job_id, job in document["jobs"].items():
            job_problems, job_minimum = _job_problems(name, job_id, job)
            problems += job_problems
            minimum = job_minimum or minimum

    if MINIMUM_JOB[0] in workflows and minimum is None:
        problems.append(f"{MINIMUM_JOB[0]}: no {MINIMUM_JOB[1]} job pins the minimum Python")
    if minimum is not None:
        floor = ".".join(minimum.split(".")[:2])
        stated = FLOOR_SENTENCE.findall(readme or "")
        if stated != [floor]:
            problems.append(f"{README} must state the minimum once as 'Python {floor} or later'; "
                            f"found {stated}")
    return problems


def repository_problems(root: Path = ROOT) -> list[str]:
    """Check this repository's workflows, version file and README."""
    workflows = {path.name: path.read_text(encoding="utf-8")
                 for path in sorted((root / WORKFLOW_DIR).glob("*.y*ml"))}
    version_path = root / PYTHON_VERSION_FILE
    readme_path = root / README
    return check(
        workflows,
        version_path.read_text(encoding="utf-8") if version_path.exists() else None,
        readme_path.read_text(encoding="utf-8") if readme_path.exists() else None,
    )
