"""Deterministic fixtures for the #46 workflow pin policy; also enforced by the gate."""
from __future__ import annotations

from workflow_pins import analyse

SHA = "11d5960a326750d5838078e36cf38b85af677262"
DIGEST = "sha256:" + "0123456789abcdef" * 4
WF = ".github/workflows/ci.yml"


def workflow(*steps: str) -> str:
    body = "\n".join("      " + s for s in steps)
    return "on: push\njobs:\n  build:\n    runs-on: ubuntu-latest\n    steps:\n" + body + "\n"


def only(*steps: str, **extra: str) -> dict[str, str]:
    files = {WF: workflow(*steps)}
    files.update({k.replace("__", "/").replace("_dot_", "."): v for k, v in extra.items()})
    return files


# (name, files, expected diagnostic substrings, expected kinds). An empty
# substring list requires the policy to pass for that fixture.
CASES = [
    # Accepted references.
    ("full SHA with version", only(f"- uses: actions/checkout@{SHA} # v4.4.0"), [], ["remote action"]),
    ("major-only version comment", only(f"- uses: actions/checkout@{SHA} # v4"), [], ["remote action"]),
    ("quoted key and value", only(f"- \"uses\": 'actions/checkout@{SHA}' # v4.4.0"), [], ["remote action"]),
    ("single-quoted key", only(f"- 'uses': \"actions/checkout@{SHA}\" # v4"), [], ["remote action"]),
    ("flow mapping", only(f"- {{name: x, uses: actions/checkout@{SHA}, with: {{a: 1}}}} # v4"), [], ["remote action"]),
    ("action subpath", only(f"- uses: github/codeql-action/init@{SHA} # v3.1.0"), [], ["remote action"]),
    ("docker digest", only(f"- uses: docker://alpine:3.20@{DIGEST} # 3.20"), [], ["docker"]),
    ("local action followed", only("- uses: ./.github/actions/setup",
                                   _dot_github__actions__setup__action_dot_yml=
                                   f"runs:\n  using: composite\n  steps:\n    - uses: actions/setup-python@{SHA} # v5.6.0\n"),
     [], ["local action", "remote action"]),
    ("local action.yaml", only("- uses: ./tools/act",
                               tools__act__action_dot_yaml="runs:\n  using: composite\n  steps: []\n"),
     [], ["local action"]),
    ("local reusable workflow", {WF: "jobs:\n  call:\n    uses: ./.github/workflows/reuse.yaml\n",
                                 ".github/workflows/reuse.yaml": "on: workflow_call\njobs: {}\n"},
     [], ["local reusable workflow"]),
    ("remote reusable workflow", {WF: f"jobs:\n  call:\n    uses: org/repo/.github/workflows/r.yml@{SHA} # v1.0.0\n"},
     [], ["remote reusable workflow"]),
    ("cycle terminates", only("- uses: ./a",
                              a__action_dot_yml="runs:\n  steps:\n    - uses: ./b\n",
                              b__action_dot_yml="runs:\n  steps:\n    - uses: ./a\n"),
     [], ["local action", "local action", "local action"]),
    ("uses inside block scalar", only("- run: |", "    echo 'uses: owner/a@main'", "  shell: bash"), [], []),
    ("uses inside sequence block scalar", {WF: "x:\n  - |\n    uses: owner/a@main\n"}, [], []),
    ("uses inside quoted value", only("- name: 'uses: owner/a@main'"), [], []),
    ("comment only", only("# uses: owner/a@main", f"- uses: actions/checkout@{SHA} # v4"), [], ["remote action"]),
    ("yaml extension", {".github/workflows/ci.yaml": workflow("- uses: owner/a@v1")}, ["ci.yaml:6", "40-hex"], None),
    ("docker image in local action", only("- uses: ./d",
                                          d__action_dot_yml=f"runs:\n  using: docker\n  image: docker://alpine@{DIGEST} # 3.20\n"),
     [], ["local action", "docker"]),
    ("dockerfile in local action", only("- uses: ./d", d__action_dot_yml="runs:\n  using: docker\n  image: Dockerfile\n"),
     [], ["local action", "local Dockerfile"]),

    # Rejected references.
    ("moving tag", only("- uses: actions/checkout@v4"), ["ci.yml:6", "40-hex"], None),
    ("branch", only("- uses: actions/checkout@main # v4"), ["40-hex"], None),
    ("short SHA", only("- uses: actions/checkout@11d5960 # v4"), ["40-hex"], None),
    ("SHA with trailing junk", only(f"- uses: actions/checkout@{SHA}x # v4"), ["40-hex"], None),
    ("41-hex SHA", only(f"- uses: actions/checkout@{SHA}a # v4"), ["40-hex"], None),
    ("uppercase SHA", only(f"- uses: actions/checkout@{SHA.upper()} # v4"), ["40-hex"], None),
    ("no ref", only("- uses: actions/checkout # v4"), ["not owner/repo"], None),
    ("missing version comment", only(f"- uses: actions/checkout@{SHA}"), ["version comment"], None),
    ("non-version comment", only(f"- uses: actions/checkout@{SHA} # pinned"), ["version comment"], None),
    ("quoted key moving tag", only("- \"uses\": actions/checkout@v4"), ["40-hex"], None),
    ("flow mapping moving tag", only("- {name: x, uses: actions/checkout@v4}"), ["40-hex"], None),
    ("multi-line flow mapping", only("- {name: x,", "   'uses': \"actions/checkout@main\"}"), ["ci.yml:7", "40-hex"], None),
    ("docker tag", only("- uses: docker://alpine:3.20 # 3.20"), ["@sha256:"], None),
    ("docker truncated digest", only("- uses: docker://alpine@sha256:" + "a" * 63 + " # 3"), ["@sha256:"], None),
    ("docker digest trailing junk", only(f"- uses: docker://alpine@{DIGEST}z # 3"), ["@sha256:"], None),
    ("docker digest missing comment", only(f"- uses: docker://alpine@{DIGEST}"), ["version comment"], None),
    ("local action external step", only("- uses: ./.github/actions/setup",
                                        _dot_github__actions__setup__action_dot_yml=
                                        "runs:\n  steps:\n    - uses: actions/setup-python@v5\n"),
     [".github/actions/setup/action.yml:3", "40-hex"], None),
    ("nested local action", only("- uses: ./a",
                                 a__action_dot_yml="runs:\n  steps:\n    - uses: ./b\n",
                                 b__action_dot_yaml="runs:\n  steps:\n    - uses: owner/x@v1\n"),
     ["b/action.yaml:3", "40-hex"], None),
    ("missing local action", only("- uses: ./nowhere"), ["no action.yml"], None),
    ("local action leaves repository", only("- uses: ./../outside"), ["leaves the repository"], None),
    ("missing local reusable workflow", {WF: "jobs:\n  call:\n    uses: ./.github/workflows/gone.yml\n"},
     ["not found"], None),
    ("remote reusable workflow tag", {WF: "jobs:\n  call:\n    uses: org/repo/.github/workflows/r.yml@v1\n"},
     ["40-hex"], None),
    ("docker image tag in local action", only("- uses: ./d", d__action_dot_yml="runs:\n  image: docker://alpine:3\n"),
     ["d/action.yml:2", "@sha256:"], None),
    ("value on next line", only("- uses:", "    actions/checkout@v4"), ["must start on the key's line"], None),
    ("alias value", only("- uses: *checkout"), ["alias"], None),
    ("tagged value", only("- uses: !!str actions/checkout@v4"), ["tag"], None),
    ("block scalar value", only("- uses: >-", "    actions/checkout@v4"), ["block scalar"], None),
    ("escaped quoted key", only("- \"u\\x73es\": actions/checkout@v4"), ["escaped quoted key"], None),
    ("escaped quoted value", only("- uses: \"actions/checkout@v\\x34\""), ["escaped quoted scalar"], None),
    ("explicit key", only("- ? uses", "  : actions/checkout@v4"), ["explicit '?'"], None),
    ("trailing content", only(f"- uses: actions/checkout@{SHA} extra # v4"), ["trailing content"], None),
]


def fixture_problems() -> list[str]:
    failures: list[str] = []
    for name, files, expected, kinds in CASES:
        problems, inventory = analyse(files)
        text = "\n".join(problems)
        if (not expected and problems) or (expected and not problems) \
                or any(s not in text for s in expected):
            failures.append(f"fixture {name!r}: expected {expected!r}, got {problems!r}")
        if kinds is not None and [r.kind for r in inventory] != kinds:
            failures.append(f"fixture {name!r}: expected kinds {kinds!r}, "
                            f"got {[r.kind for r in inventory]!r}")
    return failures


if __name__ == "__main__":
    failures = fixture_problems()
    for problem in failures:
        print(problem)
    print(f"{len(CASES)} fixtures; {len(failures)} failures")
    raise SystemExit(bool(failures))
