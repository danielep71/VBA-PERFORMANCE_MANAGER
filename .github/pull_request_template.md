<!--
Thank you for contributing to VBA-PERFORMANCE_MANAGER.

Complete every section that applies. Use "N/A — reason" where a section does
not apply; do not manufacture evidence and do not describe a branch run as if it
certified a different commit.

The hosted `VBA source consistency` workflow runs automatically, but it performs
static text analysis only. It does not import or compile VBA, start Excel, call
Windows APIs, exercise Office-bitness branches, run the regression suite, or
validate a release workbook.
-->

## 🎯 Summary

<!--
In one or two paragraphs, explain:
- the problem or requirement;
- the behavior that changes;
- why this approach was chosen.

Describe observable behavior, not only files or procedures edited.
-->



### Linked issues

<!-- Use a closing keyword only when this PR fully resolves the issue. -->

Closes #

Related to #

---

## 🧭 Change classification

<!-- Select every category that applies. -->

- [ ] Bug fix
- [ ] Enhancement
- [ ] Refactor with no intended behavior change
- [ ] Public API or contract change
- [ ] Regression-test or fault-injection change
- [ ] Documentation only
- [ ] Repository tooling / workflow / governance
- [ ] Release preparation or artifact packaging

### Affected technical boundaries

- [ ] Timing backend, rollover, native read, or QPC behavior
- [ ] Measurement harness, baseline, overhead, or sample retention
- [ ] Statistics or sample-domain behavior
- [ ] Checkpoints or report output
- [ ] Excel Application state / shared TW ownership
- [ ] `timeBeginPeriod` / `timeEndPeriod` lifecycle
- [ ] `Application.Run` qualification or trusted-input boundary
- [ ] Office bitness or conditional compilation
- [ ] Public enum, named error, default, or failure semantics
- [ ] Required source files / import dependency graph
- [ ] None of the above

---

## 📐 Scope and contract impact

### In scope

<!-- What this PR deliberately changes. -->

- <!-- item -->

### Out of scope

<!-- What a reviewer might reasonably expect, but this PR deliberately defers. -->

- <!-- item -->

### Compatibility

<!-- Select one and explain any user-visible change below. -->

- [ ] No supported public API name, signature, enum value, error value, or default changes
- [ ] Additive, source-compatible public surface change
- [ ] Existing behavior changes to correct a defect
- [ ] Breaking API or deployment change
- [ ] Documentation / tooling only; runtime compatibility is unchanged

**Compatibility or migration notes:**

<!--
Include any caller action, import-order change, host requirement, support-policy
change, or reason an existing workbook could observe different behavior.
-->



### Production source package

- [ ] The current two-file runtime remains complete:
  `M_cPM_TIMEWASTERS.bas` + `cPerformanceManager.cls`
- [ ] Required source files changed; `INSTALLATION.md`, inventories, and release guidance were updated
- [ ] Not applicable — no production source/package change

---

## 🔧 Implementation notes

<!--
Summarize the design, invariants, important control flow, and non-obvious trade-offs.
Call out strict/non-strict differences, cleanup behavior, process-global state,
and failure evidence where relevant.
-->



---

## ✅ Verification

### Static source analysis

<!--
The workflow publishes `vba-lint-results.json`. The commit recorded in that
artifact must equal the PR HEAD SHA claimed here.
-->

| Evidence | Result |
|---|---|
| Exact PR HEAD SHA | <!-- Paste the full 40-character SHA --> |
| `VBA source consistency` | <!-- PASS / FAIL / NOT RUN --> |
| Checks passed | <!-- e.g. 12 / 12 --> |
| Workflow run / artifact URL | <!-- direct GitHub Actions URL --> |
| Working tree used for any local check | <!-- clean / dirty; explain if dirty --> |

### Excel / VBA execution

<!--
Choose one applicability statement. For any production VBA, test-harness, demo,
or package change, run against the exact PR HEAD source or explain why it was not
possible. A static pass is not a substitute.
-->

- [ ] Required and completed against the exact PR HEAD
- [ ] Required but not completed — reason and release consequence documented below
- [ ] Not required — documentation/repository-only change with no executable-source or packaging impact

| Evidence | Result |
|---|---|
| Tested commit SHA | <!-- Full 40-character SHA, or N/A --> |
| Source / package tested | <!-- exact PR source, release workbook, add-in, other / N/A --> |
| VBA compile | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Regression entry point | `Run_cPerformanceManager_RegressionSuite` / N/A |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Cases | <!-- number / N/A --> |
| Assertions | <!-- number / N/A --> |
| Failures | <!-- number / N/A --> |
| Cleanup outcome | <!-- clean / cleanup failure / not observed / N/A --> |
| Excel version and build | <!-- full About Excel line / N/A --> |
| Office bitness | <!-- 32-bit / 64-bit / N/A --> |
| Windows version | <!-- version/build if runtime behavior is relevant / N/A --> |
| Host | <!-- source workbook, release workbook, add-in, other / N/A --> |

> [!IMPORTANT]
> A 64-bit Office run does not execute the 32-bit declarations or the Win32
> backend-2 binding. Shared unsigned arithmetic coverage is useful, but it is
> not Office 32-bit execution certification.

### Targeted and manual checks

<!--
List focused checks beyond the standard suite: fault injection, workbook names
with apostrophes, overlapping instances, clean/dirty workbook lifecycle,
release-asset smoke tests, or manual API examples.
-->

| Check | Result / evidence |
|---|---|
| <!-- scenario --> | <!-- PASS / FAIL / N/A + detail --> |
| <!-- scenario --> | <!-- PASS / FAIL / N/A + detail --> |

---

## 🧪 Regression coverage

- [ ] Existing regression cases cover the changed behavior
- [ ] New or amended cases cover the success path
- [ ] Failure, fallback, and cleanup paths are covered where applicable
- [ ] Strict and non-strict policies are both covered where applicable
- [ ] `TotalSteps` matches the executed case count
- [ ] Fault-injection seams remain deterministic and are reset after use
- [ ] No regression change is needed — rationale below

**Coverage rationale / new case names:**



---

## ⚠️ Risk, rollback, and recovery

### Risk level

- [ ] Low — documentation, metadata, or mechanically verified change
- [ ] Medium — localized runtime/tooling behavior with bounded impact
- [ ] High — timing integrity, process-wide Excel state, native API, release runner, or compatibility impact

### Principal risks

<!--
What can fail even if the normal-path tests pass? Consider mixed backends,
invalid zero samples, stale shared ownership, incomplete cleanup, conditional
compilation, Application.Run ambiguity, release-source drift, and binary identity.
-->

- <!-- principal risk -->

### Rollback / recovery

<!--
State how the change can be reverted and any Excel-process recovery required.
Do not rely on Class_Terminate after a hard VBA End.
-->

- <!-- rollback or recovery step -->

---

## 📚 Documentation and release hygiene

- [ ] `CHANGELOG.md` updated under `[Unreleased]` for material user-visible behavior
- [ ] Previously released CHANGELOG sections were not rewritten
- [ ] README / installation / security / releasing guidance updated where affected
- [ ] Wiki impact identified; updated now or explicitly deferred
- [ ] Procedure headers and examples agree with the implementation
- [ ] Version stamps remain unchanged unless this is the deliberate release-stamp commit
- [ ] Public API, named errors, tests, static checks, and required-source counts are not manually overstated
- [ ] No unintended workbook, binary, generated evidence, credential, or private file is included
- [ ] Line endings and binary classification remain consistent with `.gitattributes`
- [ ] Not applicable items are explained rather than silently ignored

---

## 👀 Reviewer focus

<!--
Point reviewers to the decisions most likely to be wrong despite green checks.
Include procedure names, file paths, invariants, or specific diff sections.
-->

1. <!-- first review focus -->
2. <!-- second review focus -->

### Unresolved questions or accepted trade-offs

<!-- Write "None" when there are none. -->



---

## ✅ Final author check

- [ ] The PR title describes the observable change
- [ ] The linked issue acceptance criteria are fully met or remaining work is explicit
- [ ] The evidence above belongs to the exact SHA claimed
- [ ] `Debug → Compile VBAProject` succeeded when executable VBA changed
- [ ] `git diff --check` passes
- [ ] No merge markers, placeholder evidence, unexplained `N/A`, or unrelated edits remain
- [ ] I reviewed the complete diff, including comments and documentation claims
