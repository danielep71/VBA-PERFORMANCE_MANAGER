<!--
  Keep this pull request focused on one coherent outcome.
  Complete every common section. Delete optional profile blocks that do not apply.
  Use NOT RUN or NOT APPLICABLE with a reason; never manufacture PASS evidence.
  Record only checks and environments exercised against the exact candidate.
  Report vulnerabilities privately through SECURITY.md; do not disclose secrets,
  exploitable details, confidential workbooks, or restricted data in a pull request.
-->

<div align="center">

# 🔀 VBA Performance Manager Pull Request

### Timing integrity · Excel-state ownership · Exact SHA · Reproducible evidence

[![Contributing](https://img.shields.io/badge/guide-CONTRIBUTING-217346?style=flat-square)](../CONTRIBUTING.md)
[![Security](https://img.shields.io/badge/security-private%20reporting-d73a49?style=flat-square)](../SECURITY.md)
[![Release](https://img.shields.io/badge/release-RELEASING-6f42c1?style=flat-square)](../RELEASING.md)
[![Changelog](https://img.shields.io/badge/changes-Unreleased-d97706?style=flat-square)](../CHANGELOG.md)

</div>

---

> [!IMPORTANT]
> A static pass is not Excel execution evidence. Any runtime, test, demo, or package change must identify the exact candidate tested and the completion and cleanup state.

## 📌 Summary

<!-- State the observable outcome and why it is needed. Prefer one precise purpose. -->

## 🔗 Related issues

```text
Closes #
Related to #
```

Use a closing keyword only when this pull request satisfies the issue's complete acceptance criteria.

## 🧭 Change classification

- [ ] Defect correction
- [ ] Backward-compatible capability
- [ ] Breaking API, behavior, deployment, or migration change
- [ ] Internal refactor with no intended supported-behavior change
- [ ] Test, fixture, reference-data, or validation change
- [ ] Performance change
- [ ] Security or trust-boundary hardening
- [ ] Documentation-only change
- [ ] Repository tooling, workflow, or governance change
- [ ] Packaging or release preparation
- [ ] Timing backend, statistics, or measurement-integrity change
- [ ] Excel-state, time-waster, lifecycle, or recovery change
- [ ] Fault-injection, regression, or release-provenance change

## 🎚️ Affected surface

- [ ] Timing backend, rollover, native read, or QPC behavior
- [ ] Measurement harness, baseline, overhead, or sample retention
- [ ] Statistics and sample-domain behavior
- [ ] Checkpoints and report output
- [ ] Excel Application state and shared time-waster ownership
- [ ] `timeBeginPeriod` / `timeEndPeriod` lifecycle
- [ ] `Application.Run` qualification and trusted-input boundary
- [ ] No runtime or supported surface — documentation/repository-only

---

## 📐 Scope and contract impact

### In scope

- <!-- Deliberate outcome -->

### Out of scope

- <!-- Reasonable adjacent work deliberately deferred -->

### Supported behavior and compatibility

```text
Supported behavior changed:       Yes / No
Backward compatible:              Yes / No / Uncertain
Suggested release impact:         none / patch / minor / major / uncertain
New supported members:
Removed or renamed members:
Changed signatures or defaults:
Changed results, errors, state, or side effects:
Migration required:
Known limitation introduced or retained:
```

Assess compatibility against documented behavior, not merely the VBA `Public` keyword. Infrastructure callbacks, Ribbon entry points, test seams, and `Application.Run` targets are not automatically supported API.

### Production source and package

The two-file production package: `src/modules/M_cPM_TIMEWASTERS.bas` and `src/classes/cPerformanceManager.cls`.

- [ ] Required source files and import order are unchanged.
- [ ] Required source files or order changed and `INSTALLATION.md` was updated.
- [ ] No production source/package impact.

## 🔧 Implementation notes

```text
Approach and key invariant:
Alternatives considered:
New dependency, reference, or generated input:
State ownership and cleanup:
Failure behavior:
```

Explain decisions a future reviewer cannot safely infer from the diff.

---

## ✅ Verification

### Candidate identity

| Evidence | Result |
| --- | --- |
| Exact PR HEAD SHA | <!-- Full 40-character SHA --> |
| Base branch and base SHA | <!-- Branch + full SHA --> |
| Working tree used locally | <!-- clean / dirty; explain --> |
| Source or package tested | <!-- Exact candidate source / artifact / N/A --> |

Evidence from another commit does not certify this candidate.

### Static and repository checks

- `python3 tools/vba_lint.py --json vba-lint-results.json`
- `git diff --check`

| Check | Result / evidence |
| --- | --- |
| Hosted required checks | <!-- PASS / FAIL / NOT RUN + workflow URL --> |
| Local static command | <!-- Command + PASS / FAIL / NOT RUN --> |
| Formatting / `git diff --check` | <!-- PASS / FAIL --> |
| Machine-readable artifact | <!-- Name / URL / not produced --> |

### Excel and VBA execution

- [ ] Required and completed against the exact PR HEAD.
- [ ] Required but incomplete — reason and merge/release consequence stated.
- [ ] Not required — documentation/repository-only change with no executable or packaging impact.

Relevant entry points:

- `Run_cPerformanceManager_RegressionSuite`
- Focused fault-injection, workbook-lifecycle, and package smoke scenarios

| Evidence | Result |
| --- | --- |
| Tested commit SHA | <!-- Full SHA or N/A --> |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Regression/certification entry point | <!-- Exact procedure --> |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Cases / assertions / failures | <!-- Counts or N/A --> |
| Skipped / cleanup outcome | <!-- Counts and state or N/A --> |
| Focused and manual checks | <!-- Scenarios + result --> |
| Evidence file or workflow | <!-- Name / URL / N/A --> |

### Validation environment

```text
Excel product, version, and build:
Office bitness:                    32-bit / 64-bit
Windows version/build:
Workbook or add-in host:
Deployment model:
Requested and effective timing backend
Measurement workload and Excel state
Release workbook filename when packaging
```

Record only tested environments. Source inspection does not constitute host execution, and one Office bitness does not execute the other conditional branch.

### Regression coverage

- [ ] Existing tests cover the changed success path.
- [ ] New or amended tests cover each corrected defect.
- [ ] Boundary, invalid-input, failure, fallback, and cleanup paths are covered as applicable.
- [ ] Test entry points and inventory/count metadata remain synchronized.
- [ ] Expected results come from the contract or an independent reference.
- [ ] No regression change is needed — rationale recorded below.

```text
Coverage rationale and new test names:
Unexecuted or deferred coverage:
```

---

## ⚠️ Risk, rollback, and recovery

- [ ] Low — documentation, metadata, or mechanically verified change.
- [ ] Medium — bounded runtime, tooling, or compatibility impact.
- [ ] High — numerical integrity, shared Excel state, native API, security, release, or breaking impact.

```text
Principal failure modes:
Residual risk after validation:
Rollback or revert procedure:
Excel-process, workbook, data, or artifact recovery:
Conditions that make rollback unsafe:
```

## 🔐 Security, data, and provenance

- [ ] No credential, secret, signing material, internal URL, or personal path is included.
- [ ] No client, employer, counterparty, student, personal, or restricted production data is included.
- [ ] Test data is synthetic, anonymized, or explicitly redistributable.
- [ ] External algorithms, code, datasets, and market/vendor data have attributable provenance and compatible licensing.
- [ ] Formula, command, path, callback, deserialization, and external-content injection surfaces were assessed.
- [ ] No security-sensitive detail belongs in private disclosure instead of this pull request.
- [ ] Generated evidence identifies its inputs, tool/runtime version, candidate SHA, and limitations.

```text
Security or privacy impact:
Source/data provenance:
New trust boundary:
```

## 📚 Documentation and release hygiene

- [ ] `README.md` reflects supported behavior and examples.
- [ ] `INSTALLATION.md` reflects paths, dependencies, import order, validation, upgrades, and removal.
- [ ] `CONTRIBUTING.md` reflects development and evidence requirements.
- [ ] `CHANGELOG.md` records material change under `[Unreleased]`.
- [ ] `SECURITY.md` reflects supported versions or trust boundaries.
- [ ] `RELEASING.md` reflects certification, package, provenance, or recovery changes.
- [ ] Source headers, API references, demos, Wiki pages, and counts remain synchronized.
- [ ] Version markers remain unchanged unless this is the deliberate release-stamp change.
- [ ] No documentation change is required — reason recorded below.

```text
Documentation impact:
Release, artifact, or migration impact:
```

---

## 🧩 Project-specific review

<details>
<summary><strong>⏱️ Timing and statistical integrity</strong></summary>

Keep for measurement, backend, statistics, formatting, or performance claims.

- [ ] Requested/effective backend and fallback behavior are explicit.
- [ ] Rollover, native-read, zero/negative/invalid samples, and extreme finite values are covered.
- [ ] Warm-up, sample count, baseline, overhead subtraction, dispersion, and clock environment are recorded.
- [ ] Observations produced after an invalid fallback are rejected according to contract.
- [ ] Formatting does not overflow or misrepresent large finite durations.
- [ ] Correctness is established independently of a favorable performance result.

</details>
<details>
<summary><strong>🔒 Excel state and time-waster ownership</strong></summary>

Keep when shared Excel state or auxiliary workloads can change.

- [ ] Each changed setting has readback before ownership is claimed.
- [ ] Caller-owned state is restored on success, failure, nesting, and partial initialization.
- [ ] `timeBeginPeriod` is balanced by `timeEndPeriod` on every acquired path.
- [ ] Instance-key exhaustion/overflow and overlapping managers are deterministic.
- [ ] Qualified `Application.Run` targets remain trusted and apostrophe-safe.
- [ ] Hard VBA termination limitations and process recovery are stated honestly.

</details>
<details>
<summary><strong>🧪 Regression and fault injection</strong></summary>

Keep for harness or coverage changes.

- [ ] Success, failure, fallback, strict/non-strict, and cleanup paths are covered.
- [ ] Fault-injection seams are deterministic and reset after use.
- [ ] `TotalSteps` agrees with the executed inventory.
- [ ] The result records cases, assertions, failures, cleanup, host, and exact commit.
- [ ] A headless or manual result is described only as the evidence it actually provides.

</details>
<details>
<summary><strong>📦 Release workbook and provenance</strong></summary>

Keep for packaging or release preparation.

- [ ] `PERFORMANCE.MANAGER.xlsm` is assembled from the exact candidate.
- [ ] The final-saved workbook is reopened and smoke-tested.
- [ ] `tools/release_provenance.py` output contains no TODO, missing, or inconsistent fields.
- [ ] Source/tag identity, Excel environment, suite result, asset filename, and SHA-256 agree.
- [ ] No post-validation binary edit or silent asset replacement occurred.

</details>

---

## 👀 Reviewer focus

```text
Highest-risk decision:
Files and procedures to inspect first:
Evidence to challenge:
Known boundary not proved by this pull request:
Unresolved question or accepted trade-off:
```

## ☑️ Final author check

- [ ] The title describes the observable outcome.
- [ ] The pull request has one coherent purpose and no unrelated churn.
- [ ] Linked issue acceptance criteria are met or remaining work is explicit.
- [ ] Compatibility and release impact are assessed.
- [ ] Evidence belongs to the exact candidate claimed.
- [ ] Required checks are terminal and passing; incomplete work is not presented as PASS.
- [ ] Executable VBA was compiled and tested when required.
- [ ] Failure, cleanup, and recovery behavior were reviewed.
- [ ] The complete diff, including comments, metadata, binary companions, and documentation, was reviewed.
- [ ] No merge marker, stale placeholder, unexplained N/A, accidental binary, or private material remains.

---

**Review principle:** approve the smallest coherent change whose contract, evidence, risk, and recovery can all be explained from this pull request.
