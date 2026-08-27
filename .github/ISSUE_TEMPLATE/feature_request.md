---
name: "✨ Feature request"
about: "Propose a capability or a change to the public surface"
title: "[Feature] "
labels: "enhancement"
assignees: ""
---

<!--
Thank you for proposing an improvement to VBA Performance Manager.

Start with the user problem and observable outcome. A clear problem can have
several valid designs, and maintainers may choose a different implementation.

Before submitting:
- Search open and closed issues, the README, and the wiki for related work.
- Keep one request focused on one user outcome.
- Remove client data, workbook internals, credentials, and other sensitive text.

Keep the headings below. Replace the prompts and examples with your proposal.
-->

## 🎯 Problem and user story

<!--
Describe the limitation in today's behavior. Who experiences it, under what
conditions, and what practical cost does it create?

Suggested form: As a <user>, I need <capability>, so that <outcome>.
-->


## 🧭 Desired outcome

<!--
Describe what success looks like from the caller's perspective without requiring
one particular implementation. Include concrete examples or observable results.
-->


## ✅ Acceptance criteria

<!--
List verifiable outcomes. Cover the normal path, important boundary cases, and
failure/cleanup behavior. Avoid implementation tasks such as "add a module."
-->

- [ ] <!-- Observable success criterion -->
- [ ] <!-- Boundary or failure-path criterion -->
- [ ] <!-- Compatibility or cleanup criterion -->


## 🔧 Current workaround and cost

<!--
What do users do today? Explain extra code, reliability risk, performance cost,
maintenance burden, or missing evidence. Write "No known workaround" if true.
-->


## 🧩 Candidate API or design

<!--
Optional. Show how the capability might feel to a VBA caller. Identify proposed
names, defaults, return values, error behavior, and status/diagnostic outputs.
Public members become compatibility commitments, so prefer the smallest surface
that expresses the user outcome.
-->

```vb
' Optional illustrative API — maintainers may choose another design.
Dim PM As cPerformanceManager
Set PM = New cPerformanceManager

' Example usage here.
```


## 📐 Scope and affected boundaries

<!-- Check every area that may be affected and explain the checked items below. -->

- [ ] Core timing and elapsed reads
- [ ] Procedure measurement, baseline, or overhead samples
- [ ] Statistics or sample-vector semantics
- [ ] Checkpoints, laps, splits, or reporting
- [ ] TimeWasters and shared Excel application state
- [ ] Pause/resume or timer resolution
- [ ] `Application.Run` target handling or trust boundary
- [ ] Native Windows APIs, VBA7 declarations, or Office bitness
- [ ] Public class API or error/status contract
- [ ] Source files, import order, or dependency graph
- [ ] Documentation, demos, tests, lint, or release tooling only

**Why these areas are affected**

<!-- Map the requested outcome to the selected boundaries. -->


## ⚖️ Compatibility and release impact

**Public behavior**

- [ ] Internal or documentation-only; no public behavior change
- [ ] Backward-compatible additive functionality
- [ ] Backward-compatible bug fix to documented or intended behavior
- [ ] Changes an existing default, result, error, or status semantic
- [ ] Removes or renames a public API, or otherwise requires caller changes

**Deployment**

- [ ] No new runtime source file
- [ ] Adds or changes required runtime files/import order
- [ ] No new external dependency
- [ ] Adds a dependency or host prerequisite
- [ ] Must support both 32-bit and 64-bit Office
- [ ] Intentionally targets a narrower Office/Windows environment

<!-- Explain migration impact, deprecation needs, and the oldest intended host. -->


<details>
<summary>SemVer guidance</summary>

- **Patch**: backward-compatible bug fixes.
- **Minor**: backward-compatible functionality.
- **Major**: incompatible public API or behavior changes.

A correctness fix may change defective behavior without automatically requiring a
major release. The contract, caller impact, migration path, and regression risk
still need explicit review. Maintainers assign the final release and milestone.

</details>


## 🏗️ Design constraints and invariants

<!--
Explain how the proposal preserves or intentionally changes applicable contracts:

- cPerformanceManager remains the public facade; helper-module internals remain
  private unless a public commitment is deliberate.
- A timer session binds to its active backend until an explicit, observable
  transition; requested and active methods remain distinguishable.
- A measurement vector has documented backend/failure semantics and does not
  silently mix incomparable samples.
- Strict and non-strict behavior, fallbacks, named errors, and diagnostic outputs
  remain explicit and testable.
- Nested TimeWasters sessions restore process-wide Excel state exactly once and
  do not leak state across workbooks.
- Error paths restore timer resolution and Excel application state.
- Application.Run procedure names are trusted executable input and must not be
  broadened into an unsafe discovery/execution surface.
- Source-first deployment remains inspectable; compiled workbooks are optional
  regression hosts, not runtime dependencies.

Write "Not applicable" for this section only when the request is documentation or
tooling with no runtime effect.
-->


## 🧪 Verification plan

<!--
Describe evidence that would make the change reviewable. Static CI does not run
Excel, so runtime changes normally need focused Excel checks plus the regression
suite. Include exact assertions, not only "add tests."
-->

- **Normal path:** <!-- Expected values/statuses and how to observe them -->
- **Boundary cases:** <!-- Zero/negative values, nesting, rollover, empty samples, etc. -->
- **Failure and fallback paths:** <!-- Strict/non-strict, native read failure, invalid target, etc. -->
- **Cleanup:** <!-- Excel state, timer resolution, active sessions, error exit -->
- **Regression suite:** <!-- Cases to add/change; expected totals if known -->
- **32/64-bit coverage:** <!-- Required when declarations or pointer-sized values change -->
- **Documentation/release evidence:** <!-- README, wiki, API inventory, changelog, migration notes -->


## 🔀 Alternatives and non-goals

**Alternatives considered**

<!-- Include smaller API, documentation-only, or caller-side solutions and why they fall short. -->


**Explicit non-goals**

<!-- Prevent the request from expanding into adjacent work. -->


## 📚 Examples, prior art, and additional context

<!--
Link related issues or discussions. Include short, anonymized examples and explain
which behavior is worth adopting. Do not paste proprietary workbook content.
-->


## ✅ Requester checklist

- [ ] I searched open and closed issues, the README, and the wiki.
- [ ] I described a user problem and observable outcome, not only an implementation.
- [ ] I supplied testable acceptance criteria.
- [ ] I considered compatibility, cleanup, and deployment impact.
- [ ] I identified the relevant runtime boundaries and invariants.
- [ ] I removed credentials, client data, and other sensitive information.
- [ ] I understand maintainers may solve the problem differently and assign the release.
