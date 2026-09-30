"""Immutable action-reference policy for GitHub workflows; not a YAML parser.

Every ``uses:`` reference reachable from ``.github/workflows/*.yml|*.yaml`` is
classified and checked:

- remote action or reusable workflow: ``owner/repo[/path]@<40-hex commit SHA>``
  followed on the same line by a version comment such as ``# v4.4.0``;
- Docker action: ``docker://image[:tag]@sha256:<64 hex>`` with the same comment;
- local action (``./dir``): followed recursively into ``dir/action.yml`` or
  ``dir/action.yaml`` with cycle protection; the reference itself is trusted
  because it is reviewed in this repository;
- local reusable workflow (``./.github/workflows/x.yml``): must exist, and is
  checked as a workflow in its own right.

Local action metadata is also checked for ``image: docker://...`` references.

The scanner understands the YAML a workflow can use to spell a ``uses`` key:
plain and quoted keys, block and flow mappings, comments and block scalars. It
fails closed on constructs it does not model (explicit ``?`` keys, escaped
quoted keys, aliases, tags, anchors, and values that do not start on the key's
line) instead of silently skipping them.
"""
from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass

CHECK = "workflow actions use immutable pins"
WORKFLOW_DIR = ".github/workflows"

SHA = re.compile(r"[0-9a-f]{40}")
VERSION_COMMENT = re.compile(r"v?\d+(?:\.\d+){0,3}(?:[-+][0-9A-Za-z.-]+)?")
REMOTE = re.compile(r"([A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)/"
                    r"([A-Za-z0-9._-]+)((?:/[^@\s/]+)*)@(.+)")
DOCKER = re.compile(r"docker://[^@\s]+@sha256:[0-9a-f]{64}")
BLOCK_SCALAR = re.compile(r"[|>](?:[1-9][+-]?|[+-][1-9]?)?")


@dataclass
class Reference:
    path: str
    line: int
    key: str
    value: str | None
    comment: str
    kind: str = ""


def _problem(ref: Reference, message: str) -> str:
    return f"{ref.path}:{ref.line}: {message}"


def _read_quoted(line: str, i: int) -> tuple[str | None, int, bool]:
    """Read a quoted scalar starting at ``line[i]``.

    Returns (content, index after the closing quote, escaped). Content is None
    when the scalar does not close on this line.
    """
    quote, out, j = line[i], [], i + 1
    escaped = False
    while j < len(line):
        c = line[j]
        if quote == "'" and c == "'":
            if j + 1 < len(line) and line[j + 1] == "'":
                out.append("'")
                j += 2
                continue
            return "".join(out), j + 1, escaped
        if quote == '"' and c == "\\":
            escaped = True
            out.append(line[j:j + 2])
            j += 2
            continue
        if quote == '"' and c == '"':
            return "".join(out), j + 1, escaped
        out.append(c)
        j += 1
    return None, len(line), escaped


def scan(path: str, text: str, keys: tuple[str, ...] = ("uses",)) -> tuple[list[Reference], list[str]]:
    """Find every mapping key in ``keys`` and return its same-line scalar."""
    refs: list[Reference] = []
    problems: list[str] = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    block_indent: int | None = None   # column of the node that opened a block scalar
    flow_depth = 0                    # open { or [ carried across lines
    open_quote: str | None = None     # quote left open by a multi-line scalar

    def key_follows(k: int) -> bool:
        return k < len(line) and line[k] == ":" and (
            k + 1 == len(line) or line[k + 1] in " \t"
            or (flow_depth > 0 and line[k + 1] in ",{}[]"))

    for number, line in enumerate(lines, 1):
        stripped = line.lstrip(" ")
        if block_indent is not None:
            if not stripped or len(line) - len(stripped) > block_indent:
                continue
            block_indent = None

        i = 0
        if open_quote:
            close = line.find(open_quote)
            while open_quote == "'" and close != -1 and line[close:close + 2] == "''":
                close = line.find("'", close + 2)
            if close == -1:
                continue
            i, open_quote = close + 1, None

        while i < len(line):
            c = line[i]
            if c in " \t":
                i += 1
                continue
            if c == "#" and (i == 0 or line[i - 1] in " \t"):
                break
            if c in "-?" and (i + 1 == len(line) or line[i + 1] in " \t"):
                if c == "?":
                    problems.append(f"{path}:{number}: explicit '?' mapping keys are not supported")
                rest = line[i + 1:].split(" #", 1)[0].strip()
                if c == "-" and not flow_depth and BLOCK_SCALAR.fullmatch(rest):
                    block_indent = i
                    break
                i += 1
                continue
            if c in "{[":
                flow_depth += 1
                i += 1
                continue
            if c in "}]":
                flow_depth = max(0, flow_depth - 1)
                i += 1
                continue
            if c == "," and flow_depth:
                i += 1
                continue

            # One scalar token, quoted or plain.
            start, quoted, escaped = i, c in "\"'", False
            if quoted:
                token, i, escaped = _read_quoted(line, i)
                if token is None:
                    open_quote = c
                    break
            else:
                j = i
                while j < len(line) and not (
                        (flow_depth and line[j] in ",{}[]")
                        or key_follows(j)
                        or (line[j] == "#" and line[j - 1] in " \t")):
                    j += 1
                token = line[i:j].rstrip()
                i = max(i + len(token), i + 1)

            k = i
            while k < len(line) and line[k] in " \t":
                k += 1
            if not key_follows(k):
                continue   # a value or free text

            # An escape sequence could spell a checked key; do not decode it.
            if quoted and escaped:
                problems.append(f"{path}:{number}: escaped quoted key is not supported")
            i = k + 1
            if token not in keys:
                rest = line[i:].split(" #", 1)[0].strip()
                if not flow_depth and BLOCK_SCALAR.fullmatch(rest):
                    block_indent = start
                    break
                continue

            ref = Reference(path, number, token, None, "")
            refs.append(ref)
            while i < len(line) and line[i] in " \t":
                i += 1
            if i == len(line) or line[i] == "#":
                problems.append(_problem(ref, f"'{token}' value must start on the key's line"))
                break
            v = line[i]
            if v in "*&!|>":
                problems.append(_problem(ref, f"'{token}' value uses an unsupported YAML "
                                              "alias, anchor, tag or block scalar"))
                break
            if v in "\"'":
                content, i, escaped = _read_quoted(line, i)
                if content is None or escaped:
                    problems.append(_problem(ref, f"'{token}' value is an unsupported "
                                                  "multi-line or escaped quoted scalar"))
                    if content is None:
                        open_quote = v
                    break
                ref.value = content
            else:
                j = i
                while j < len(line) and line[j] not in " \t" and not (
                        flow_depth and line[j] in ",}]"):
                    j += 1
                ref.value, i = line[i:j], j

            # Only a comment or a flow separator may follow the value.
            while i < len(line) and line[i] in " \t":
                i += 1
            if i < len(line) and line[i] != "#" and not (flow_depth and line[i] in ",}]"):
                problems.append(_problem(ref, f"'{token}' value has trailing content"))
            hash_at = line.find("#", i)
            while hash_at > 0 and line[hash_at - 1] not in " \t":
                hash_at = line.find("#", hash_at + 1)
            if hash_at != -1:
                ref.comment = line[hash_at + 1:].strip()
    return refs, problems


def classify(ref: Reference) -> tuple[str, str | None]:
    """Return (kind, problem) for one reference value."""
    value = ref.value or ""
    if value.startswith("./"):
        if re.fullmatch(r"\./\.github/workflows/[^/]+\.ya?ml", value):
            return "local reusable workflow", None
        return "local action", None
    if value.startswith("docker://"):
        if not DOCKER.fullmatch(value):
            return "docker", "Docker reference is not pinned by a complete @sha256: digest (64 hex)"
    else:
        match = REMOTE.fullmatch(value)
        if not match:
            return "invalid", "reference is not owner/repo[/path]@ref, ./local or docker://"
        subpath, ref_part = match.group(3), match.group(4)
        kind = ("remote reusable workflow"
                if re.fullmatch(r"/\.github/workflows/[^/]+\.ya?ml", subpath)
                else "remote action")
        if not SHA.fullmatch(ref_part):
            return kind, "ref is not a full 40-hex commit SHA"
    kind = "docker" if value.startswith("docker://") else kind
    if not VERSION_COMMENT.fullmatch(ref.comment):
        return kind, "immutable pin lacks a trailing version comment (e.g. # v1.2.3)"
    return kind, None


def analyse(files: dict[str, str]) -> tuple[list[str], list[Reference]]:
    """Check every workflow in ``files`` (POSIX paths relative to the root)."""
    problems: list[str] = []
    inventory: list[Reference] = []
    pending = sorted(p for p in files
                     if posixpath.dirname(p) == WORKFLOW_DIR
                     and p.endswith((".yml", ".yaml")))
    visited: set[str] = set()

    while pending:
        path = pending.pop(0)
        if path in visited:
            continue
        visited.add(path)
        is_action = posixpath.basename(path) in ("action.yml", "action.yaml")
        refs, scan_problems = scan(path, files[path],
                                   ("uses", "image") if is_action else ("uses",))
        problems += scan_problems
        for ref in refs:
            if ref.value is None:
                continue
            if ref.key == "image":
                if not ref.value.startswith("docker://"):
                    ref.kind = "local Dockerfile"
                    inventory.append(ref)
                    continue
            kind, problem = classify(ref)
            ref.kind = kind
            inventory.append(ref)
            if problem:
                problems.append(_problem(ref, f"{problem}: {ref.value}"))
                continue
            if kind.startswith("local"):
                target = posixpath.normpath(ref.value[2:])
                if target == ".." or target.startswith("../"):
                    problems.append(_problem(ref, f"local reference leaves the repository: {ref.value}"))
                elif kind == "local reusable workflow":
                    if target not in files:
                        problems.append(_problem(ref, f"local reusable workflow not found: {ref.value}"))
                    else:
                        pending.append(target)
                else:
                    base = "" if target == "." else target + "/"
                    found = [base + n for n in ("action.yml", "action.yaml") if base + n in files]
                    if not found:
                        problems.append(_problem(ref, f"local action has no action.yml or action.yaml: {ref.value}"))
                    pending += found
    return problems, inventory


def repository_files(root) -> dict[str, str]:
    """Collect workflows and every action metadata file under ``root``."""
    files: dict[str, str] = {}
    for p in sorted(root.rglob("*")):
        rel = p.relative_to(root).as_posix()
        if rel.startswith(".git/") or not p.is_file():
            continue
        if (posixpath.dirname(rel) == WORKFLOW_DIR and p.suffix in (".yml", ".yaml")) or \
                p.name in ("action.yml", "action.yaml"):
            files[rel] = p.read_text(encoding="utf-8")
    return files
