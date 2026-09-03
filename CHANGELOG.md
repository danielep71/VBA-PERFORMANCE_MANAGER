<div align="center">

# 📜 Changelog

### Release history for benchmark-grade VBA timing and execution control

[![Format](https://img.shields.io/badge/Format-Keep_a_Changelog-0969da?style=flat-square)](https://keepachangelog.com/en/1.1.0/)
[![Versioning](https://img.shields.io/badge/Versioning-SemVer-6f42c1?style=flat-square)](https://semver.org/spec/v2.0.0.html)
[![Dates](https://img.shields.io/badge/Dates-YYYY--MM--DD-217346?style=flat-square)](#date-and-version-rules)
[![Staging](https://img.shields.io/badge/Staging-Unreleased_first-d97706?style=flat-square)](#unreleased)
[![Contributing](https://img.shields.io/badge/Changes-Contribution_guide-2ea44f?style=flat-square)](CONTRIBUTING.md)

<br>

**User-visible history · Explicit compatibility · Reproducible evidence · Immutable releases**

</div>

---

All notable changes to **VBA Performance Manager** are documented here.

This changelog follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and [Semantic Versioning](https://semver.org/spec/v2.0.0.html). It records
released behavior and material unreleased changes; it is not a commit log, issue
tracker, or substitute for release evidence.

Versioning covers public timing and measurement APIs, backend/session semantics, sample classification, statistics, strict-mode failures, Excel-state ownership, result formats, and supported environments.

---

## 🧭 Maintenance policy

- Add material changes under **Unreleased** in the same pull request as the
  behavior or documentation they describe.
- Write from the user's perspective: describe the observable result, contract,
  compatibility impact, and migration need.
- Link the owning issue or pull request when it contains useful engineering
  detail.
- Keep entries concise; do not duplicate implementation notes already preserved
  in source, issues, or technical documentation.
- Record only validation actually performed. State skipped environments and
  known limitations plainly.
- Move Unreleased entries into a dated version section during release.
- Do not edit a published release entry except to correct a demonstrable factual
  or link error; annotate material corrections instead of rewriting history.
- Never claim that a tag, binary, workbook, hash, test run, or environment was
  certified unless the evidence binds it to the released source.

See [CONTRIBUTING.md](CONTRIBUTING.md) for change and evidence requirements and
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

<a id="date-and-version-rules"></a>

### Date and version rules

| Rule | Standard |
|---|---|
| Version | `MAJOR.MINOR.PATCH`, without the leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |
| Date | Gregorian calendar date in ISO `YYYY-MM-DD` format |
| Ordering | Unreleased first; released versions newest to oldest |
| Comparison | Unreleased → latest tag; each release → preceding tag |
| Patch | Backward-compatible correction or hardening |
| Minor | Backward-compatible capability |
| Major | Incompatible public-contract change |
| Pre-release | State maturity and compatibility boundaries explicitly |

A repository may remain below `1.0.0` while its supported surface is still
forming. Pre-release status does not excuse undocumented breaking changes.

<details>
<summary><strong>Entry categories</strong></summary>

<br>

| Category | Use for |
|---|---|
| **Added** | New supported capabilities, APIs, files, or tests |
| **Changed** | Changes to existing behavior, contracts, tooling, or documentation |
| **Deprecated** | Supported behavior scheduled for removal |
| **Removed** | Removed capabilities or compatibility |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Documentation** | Material documentation-only changes |
| **Validation** | Evidence actually produced |
| **Compatibility** | Upgrade or migration effects |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories needed by a release.

</details>

---

<a id="unreleased"></a>

## [Unreleased]

No changes recorded yet.

---

## [1.4.0] — 2026-08-31

Work completed so far on the `release/v1.4.0` development line.

**Current scope: repository governance, documentation, installation, release
assurance, and the first measurement-integrity corrections. The public API is
extended additively — new trailing optional parameters on existing members, and
new `Friend` test seams — so every existing call shape continues to compile and
behave. Behavior of existing members changes where noted; see the correction
notice below.**

> [!IMPORTANT]
> **BEHAVIORAL CORRECTION — requested-backend measurement integrity.**
>
> In v1.3.0, a non-strict repeated measurement could retain observations
> produced after the requested backend fell back to method 2. The harness read
> the worker's status only after `ElapsedSeconds`, which resets it on entry, so
> the fallback was erased before it could be seen and the method-2 observation
> was kept as though it came from the requested backend. Returned vectors could
> silently mix backends.
>
> v1.4.0 captures the start transaction immediately after `StartTimer` and
> rejects the cycle. Three consequences for existing callers:
>
> - a non-strict run on a host that cannot start the requested backend now
>   returns a shorter vector, or raises `ERR_CPM_MEASURE_NO_VALID_SAMPLES`,
>   where v1.3.0 returned mixed-backend results;
> - a rejected cycle does not dispatch the workload, so a procedure with side
>   effects executes fewer times than `Iterations` requested; and
> - `FailedReadsOut` and `RejectedSamplesOut` are set to `-1` on entry and hold
>   that value unless the run reaches a publication point. A caller who reads
>   them after a raise now sees a negative count rather than zeros. Code that
>   assumed a zero count meant a clean run must test for a negative count first.
>
> The legacy `OverheadMeasurement_Seconds` accumulated every endpoint read into
> a running total and divided by the requested iteration count. A failed
> non-strict read returns zero, so it was added as a zero and still counted in
> the denominator, pulling the mean down by the proportion of reads that failed
> and making a degraded host look faster. A start fallback was retained too, so
> the figure could describe method 2 while claiming to describe the requested
> backend.
>
> It now delegates to `MeasureOverhead_Samples` and averages only the retained
> samples. This property previously always returned a `Double`; it can now raise
> `ERR_CPM_MEASURE_NO_VALID_SAMPLES` when no cycle survives.
> `OverheadMeasurement_Text` renders that expected condition as
> `undefined (<original description>)`, carrying the description through so an
> all-fallback run reads differently from an all-failed one, and propagates
> every other error unchanged.

### Added

- **Harness failure and rejection evidence on `MeasureOverhead_Samples`.** It
  accepts trailing optional `FailedReadsOut`, `LastFailureStatusOut` and
  `RejectedSamplesOut`, adopting the contract already established by
  `MeasureProcedure`: native read failures counted including warm-up, backend
  rejections counted excluding warm-up, and both counts set to `-1` until the
  run reaches a publication point. Previously this routine tracked its last
  failure status internally and then discarded it, so a short vector was the
  only signal a caller had and nothing distinguished an unreliable clock from a
  backend the host could not start. All three are trailing and optional, so
  existing calls are unaffected. (#28)
- **`RejectedSamplesOut` on the repeated-measurement APIs.** `MeasureProcedure`
  and `MeasureBaseline` accept a trailing optional `ByRef RejectedSamplesOut As
  Long` reporting measured cycles rejected because the requested backend fell
  back. It is counted separately from `FailedReadsOut` because the two have
  different remedies: a failed read says the host's clock is unreliable, a
  rejection says this host cannot start the backend that was asked for. The
  parameter is trailing and optional, so existing calls are unaffected.

  Both counts are set to `-1` on entry, before any validation that can raise, and
  are overwritten only at a publication point. A negative count therefore means
  the run published nothing and `Err` carries the diagnosis;
  `LastFailureStatusOut` is meaningful only while the counts are non-negative.
  Writing zeros instead would make an aborted run indistinguishable from a
  flawless one. (#23)
- **`MeasureBaseline` now exposes harness failure evidence.** It accepts
  `FailedReadsOut`, `LastFailureStatusOut` and `RejectedSamplesOut`, forwarded
  unchanged from the delegated `MeasureProcedure` run. Previously a caller could
  see that a workload vector had been shortened but had no way to ask whether
  the baseline being subtracted had been shortened for the same reason. All
  three are trailing and optional.

  The same evidence boundary applies. `MeasureBaseline` applies the `-1` entry
  sentinel itself, before its own blank-name validation, so an error raised
  before delegation is as visible as one raised inside it. (#27)
- **Separate start and endpoint worker fault-injection seams.**
  `Test_ForceWorkerStartReadFailures` arms a forced native read failure on a
  measured cycle's start capture; `Test_ForceWorkerEndpointReadFailures` arms
  one on its endpoint read. Each selects the fault its backend can actually
  suffer, failing a QPC read for method 5 and a system-time read for method 4.
  The v1.3.0 seam armed after `Application.Run` and forced a QPC failure
  unconditionally, so it could never fail a start capture and had no effect on
  any other backend. `Test_ForceWorkerReadFailures` is retained as a documented
  compatibility alias for the endpoint seam. Both counters are cleared on every
  exit from `MeasureProcedure`, normal or raised. (#23)
- **Deterministic repository text and artifact policy.** A root
  `.gitattributes` now keeps exported VBA source on CRLF for predictable VBE
  round-trips, keeps documentation and cross-platform tooling on LF, prevents
  line-merging of Office packages and other binary artifacts, identifies VBA
  explicitly for GitHub Linguist, and excludes repository plumbing from source
  archives.
- **Editor-level formatting policy.** A root `.editorconfig` mirrors the Git
  rules before normalization occurs: four-space VBA and Python indentation,
  two-space YAML/JSON indentation, host-appropriate line endings, final
  newlines, and format-specific whitespace handling without forcing an unsafe
  VBA export encoding.
- **Dedicated installation and upgrade guide.** `INSTALLATION.md` documents the
  two-file runtime package, module-first import order, compile and smoke tests,
  optional regression host, embedded-source upgrades, strict/non-strict
  diagnostics, environment recovery, release-workbook verification, removal,
  and the distinction between 32/64-bit source support and execution evidence.
- **Project-specific Code of Conduct.** `CODE_OF_CONDUCT.md` adds respectful and
  evidence-led collaboration standards, benchmark-integrity expectations,
  review criteria for timing/statistics/Application-state changes, disclosure
  and licensing rules, and private enforcement channels.

### Fixed

- **Workbook-qualified measurement targets now resolve when the host workbook
  name contains an apostrophe.** `MeasureProcedure` and `MeasureBaseline` quote
  the host workbook name so that names containing spaces resolve, but embedded
  apostrophes were not escaped, so a workbook such as `O'Brien.xlsm` produced
  `'O'Brien.xlsm'!Proc`. The quoted section ended at the embedded apostrophe and
  the target did not resolve, failing the measurement for a reason unrelated to
  timing. Apostrophes are now doubled, so the same host produces
  `'O''Brien.xlsm'!Proc`.

  The procedure-name spacing policy is stated explicitly as part of the same
  fix. Leading and trailing spaces are removed once, before the target is
  classified, so explicitly qualified and automatically qualified names follow
  the same rule; previously a name such as `"  Proc  "` passed the blank check,
  which trims, and was then dispatched untrimmed. Blank and space-only names
  continue to raise `ERR_CPM_MEASURE_BLANK_PROC`, and interior spaces are
  preserved. `Trim$` removes space characters only, so a leading or trailing
  tab, carriage return or line feed is not normalized and still reaches
  `Application.Run`. That limit is deliberate: validation and qualification trim
  with the same function, which is what keeps their policies aligned. (#25)

### Changed

- **`OverheadMeasurement_Seconds` derives its mean from validated samples.** It
  delegates to `MeasureOverhead_Samples` and averages the retained vector, so a
  failed read no longer contributes a zero while still counting in the
  denominator, and a fallback cycle is no longer averaged in as though it
  measured the requested backend. The mean is now computed over the retained
  sample count rather than the requested iteration count, so a degraded run
  reports a higher and more honest figure than v1.3.0 did. The warm-up count is
  passed explicitly as 1, preserving this helper's own single warm-up cycle
  rather than inheriting the vector function's default of 10. Its signature is
  unchanged: the three evidence outputs are consumed internally, and a caller
  who needs them should use `MeasureOverhead_Samples`. (#24)
- **`OverheadMeasurement_Text` degrades instead of raising.** It traps only
  `ERR_CPM_MEASURE_NO_VALID_SAMPLES` and renders
  `undefined (<original description>)`, carrying the original text so an
  all-fallback run and an all-failed run read differently. Every other error is
  re-raised with its original number, source and description: an invalid
  argument is not an undefined measurement. This follows `Stats_Text`, which
  renders an undefined coefficient of variation while
  `Stats_CoefficientOfVariation` raises. (#24)
- **Overhead vectors are now backend-homogeneous.** `MeasureOverhead_Samples`
  captures the start transaction immediately after `StartTimer` and rejects a
  cycle whose requested backend fell back, rather than retaining a method-2
  observation as though it measured the requested backend. An overhead figure
  can no longer silently describe a different clock than the one asked for. On a
  host that cannot start the requested backend, a non-strict run now returns a
  shorter vector or raises `ERR_CPM_MEASURE_NO_VALID_SAMPLES` where v1.3.0
  returned method-2 numbers. A rejected cycle takes no endpoint read. (#28)
- **The deterministic worker seams now reach both harnesses.**
  `Test_ForceWorkerStartReadFailures` and `Test_ForceWorkerEndpointReadFailures`
  apply to `MeasureOverhead_Samples` as well as `MeasureProcedure`, arming
  before `StartTimer` and before `ElapsedSeconds` respectively, and neither is
  armed during warm-up. `Test_ForceWorkerReadFailures` documented itself as
  affecting both routines in v1.3.0, which was never true — the counter was
  consumed only by `MeasureProcedure`. That claim is now accurate for the
  endpoint seam it aliases. Both counters are cleared on every exit from either
  routine, normal or raised. (#28)
- **`Measure_NoValidSamplesText` takes the calling API's unit noun.** The
  aggregate no-valid-samples description says "cycle" for
  `MeasureOverhead_Samples` and "iteration" for `MeasureProcedure`, so a
  diagnostic reads in the vocabulary of the routine that produced it. (#28)
- **Repeated-measurement vectors are now backend-homogeneous.** Every element of
  a vector returned by `MeasureProcedure` or `MeasureBaseline` was measured on
  the backend the caller requested. A measured cycle whose requested backend
  fails to start now falls back, is detected, and is rejected rather than
  retained. See the behavioral-correction notice above for the impact on
  existing callers. (#23)
- **The no-valid-samples error is status-aware.**
  `ERR_CPM_MEASURE_NO_VALID_SAMPLES` is unchanged and remains accurate at the
  aggregate level, but its description now names the dominant cause. When every
  cycle was rejected it reports the fallback and the backend that was requested,
  rather than claiming no read produced a value: in that case the reads
  succeeded and only the clock was wrong. The optional outputs are published
  before this error is raised, so a caller who traps it can still read the
  counts and status that explain it. (#23)
- **The evidence boundary for the optional outputs is stated by control flow.**
  `FailedReadsOut`, `LastFailureStatusOut` and `RejectedSamplesOut` are
  authoritative on a normal return and on the aggregate no-valid-samples raise,
  where they are published first. On any other propagated error they have not
  reached their publication point; `Err.Number` and `Err.Description` are
  authoritative instead. This covers the blank-name raise, the worker raising in
  strict mode and the defensive internal-invariant raise, and is deliberately
  not phrased in terms of strict mode, since a non-strict run can also propagate
  an error before any cycle is classified.

  The boundary is enforced rather than advisory: the `-1` entry sentinel makes
  non-publication observable, so a caller can test for it instead of being asked
  to remember it. `Err_Handler` does not touch the outputs, so evidence already
  published by the aggregate raise survives being re-raised. (#23, #27)
- **The root README is now a verified project contract rather than a feature
  catalogue.** It documents the source-first deployment model, exact two-file
  runtime package, current public API, enums and named errors, timing and
  measurement semantics, shared Excel-state ownership, recovery paths,
  architecture, security boundaries, repository structure, and assurance
  limits. It also separates the certified v1.3.0 state from the unreleased
  v1.4.0 development line.
- **Enum type-safety wording is corrected in the current README.** The timer and
  pause enums improve readability and IntelliSense, but VBA enum parameters
  have `Long` semantics and overlapping numeric values are not made
  non-interchangeable at compile time. The remaining class-comment,
  regression-example, and wiki verification work is tracked by #26; the frozen
  v1.3.0 changelog statement remains historical and is not rewritten.
- **The public wiki has been comprehensively reconciled with the certified
  v1.3.0 implementation and contracts.** Twenty-two pages now accurately
  document the public API, execution and restoration semantics, timing
  backends, checkpoints, statistics, strict-mode behavior, measurement
  failures, shared TimeWasters ownership, diagnostics, benchmarking,
  installation, testing, limitations, and version history. Stale,
  contradictory, and unsupported claims were removed or qualified, while the
  certified v1.3.0 runtime state is clearly separated from the unreleased
  v1.4.0 repository work.
- **The maintainer release process is now bound to the exact tag target.**
  `RELEASING.md` distinguishes pre-merge validation from final certification;
  requires the hosted 12-check static result and real Excel compile/regression
  evidence for the final SHA; makes bitness evidence explicit and requires real
  32-bit Office execution evidence for v1.4.0 if that support remains in scope;
  validates the actual release workbook; integrates tag/source/asset
  provenance; treats published assets as immutable evidence; and defines
  pre- and post-publication recovery paths.
- **Pull requests now carry reviewable evidence.** The pull request template
  records exact-SHA static and Excel evidence, compatibility and source-package
  impact, regression coverage, Office bitness, cleanup, risk, rollback,
  documentation, and release hygiene. It no longer claims that the live branch
  rules require pull requests or the static check when they currently do not.
- **Bug reports now distinguish the evidence surfaces that matter.** The issue
  template requests exact source identity, Excel build and bitness, error and
  timer state, requested versus active backend, caller `LastReadStatus`,
  measurement-worker failure outputs, shared TimeWasters state, cleanup, and
  regression evidence, while routing suspected vulnerabilities to the private
  security process.
- **Feature requests are now problem-first and contract-aware.** The issue
  template adds user outcomes, testable acceptance criteria, affected runtime
  boundaries, deployment and SemVer impact, design invariants, verification,
  alternatives, and explicit non-goals.
- **Repository ignore rules now enforce the source-first model.** `.gitignore`
  excludes generated Office binaries, lock/recovery files, local regression and
  provenance output, build/cache/log debris, editor state, local environments,
  secrets, and signing keys while explicitly protecting authoritative VBA,
  demos, tests, tooling, workflows, documentation, and images from broad ignore
  patterns.
- **Tracked VBA test source was normalized under the new attributes policy.**
  `test/M_cPM_Test.bas` now has deterministic VBE-compatible line endings; the
  normalization contains no semantic test change.

### Removed

- **Obsolete `tag-1.2.0.txt`.** Historical release metadata remains available
  from the repository history, release page, tag, changelog, and release
  provenance rather than a duplicated root-level notes file.

### Security

- **A repository-specific security policy now defines the supported and trusted
  operating boundary.** `SECURITY.md` establishes private vulnerability
  reporting, the Excel/VBA and Windows trust model, trusted executable input for
  `Application.Run`, native API and Office-bitness boundaries, timer-resolution
  and shared Application-state ownership, timing-result integrity limits,
  source-first release trust, macro-enabled asset verification, secret handling,
  disclosure, and future self-hosted Excel-runner requirements.

---

## [1.3.0] — 2026-08-16

Process discipline turned into enforcement, every value the API takes given a
name, and a family of defects closed in which an invalid result was
indistinguishable from a valid one.

**Regression suite: 72 cases, 511 assertions, 0 failures.**

### Added

- **Static analysis gate.** `tools/vba_lint.py` runs eleven consistency checks
  over the exported VBA sources — merge conflict markers, procedure and block
  balance, reserved-word identifiers, error sources naming the right module,
  bare error numbers, undefined local callees, test wiring, `TotalSteps` drift,
  version-stamp agreement, and single-call-site native APIs. Run on every push
  and pull request via `.github/workflows/static-checks.yml`, and required by
  the `protect-main` ruleset. Six of the eleven correspond to defects that
  actually reached this repository during v1.2.0. (#10)
- **`ERR_CPM_STATS_UNDEFINED_CV`** (+1033) and
  **`ERR_CPM_STATS_BAD_CV_THRESHOLD`** (+1034).
- **Three regression cases** covering the coefficient-of-variation semantics
  below. Suite 63 → 66 cases, 447 assertions.
- **`cPM_TimerMethod`, `cPM_PauseMethod`, `cPM_Error` and `cPM_TWError` enums.**
  Timing backends, pause strategies and error numbers are now selectable by
  name. Values are unchanged, so existing numeric calls keep working. Separate
  timer and pause types make the two non-interchangeable at compile time, which
  v1.2.0 could only warn about in documentation. (#13, #14)
- **Issue forms and a pull request template.** The bug form requires the Excel
  build, bitness, bound backend, `StrictMode` and import order — the facts
  without which a timing report cannot be triaged, since several code paths are
  selected by `#If Win64`. The pull request template records the regression
  suite result and the environment it ran on, which is the manual stand-in for
  the headless Excel gate in #11. (#21)
- **`MeasureBaseline`.** Runs an empty procedure through the same
  `Application.Run` path as a real workload, so its median can legitimately be
  subtracted from a `MeasureProcedure` result. `MeasureOverhead_Samples` cannot
  — it does not dispatch through `Application.Run`. The caller supplies the
  empty procedure; `Application.Run` cannot reach anything in a module declared
  `Option Private Module`. (#7)
- **`FailedReadsOut` and `LastFailureStatusOut`** on `MeasureProcedure`. The
  harness measures on an isolated worker released before the vector is returned,
  so a run where every read failed was previously indistinguishable from a run
  of a procedure that does nothing. (#21)
- **`ERR_CPM_MEASURE_NO_VALID_SAMPLES`** (+1036), raised when no measured read
  in a harness run produced a value.
- **`cPM_ReadElapsedInvalid`** status, reported when an elapsed value had to be
  clamped because the timing source moved backwards.
- **`ERR_CPM_STATS_OUT_OF_DOMAIN`** (+1037), raised when a sample vector
  contains a negative or non-finite value.
- **Release provenance.** `tools/release_provenance.py` emits a SHA-256 for every
  shipped file and Release asset, the commit the release was cut from, and the
  Excel build and bitness the suite was certified on. Missing fields are marked
  `TODO` and the script exits non-zero, so an incomplete block cannot ship
  unnoticed. `RELEASING.md` records the process, including the three mechanical
  failures that preceded a correct v1.2.0 tag. (#12)
- **Four measurement usage examples** and the workload targets they need:
  repeated benchmarking with statistics, comparing two implementations on
  medians, subtracting a dispatch-matched baseline, and telling a failed read
  apart from a genuinely fast operation. (#17)

### Changed

- **BEHAVIORAL CORRECTION: `Stats_CoefficientOfVariation` now raises when the
  sample mean is not strictly positive.** Earlier releases returned zero.

  For non-negative timing data a zero mean means every observation was zero,
  which is what a run of failed non-strict reads produces. Returning zero then
  made `Stats_IsContaminated` report that run as perfectly stable — certifying a
  failed measurement as clean, the worst direction for the error to point. (#6)

- **`Stats_TryCoefficientOfVariation`** is now the single definition of
  validity, with three deliberately different policies on top of it:

  | Member | Non-positive mean |
  |---|---|
  | `Stats_CoefficientOfVariation` | Raises |
  | `Stats_IsContaminated` | Returns `True` — validity unestablished is suspect |
  | `Stats_Text` | Prints `undefined (mean is not positive)` |

- **`Stats_Text` distinguishes two warnings** — *high variance; re-run* from
  *sample validity cannot be established; inspect the observations*. Those mean
  different things and previously collapsed into one message.

- **`Elapsed_ComputeSeconds` and `Elapsed_Validate` no longer raise**, reporting
  outcomes through their status outputs instead. One raise remains deliberately:
  the defensive invalid-method branch, which has no corresponding status. Strict
  policy is applied by the public operation, which publishes the status first so
  the two public surfaces agree.

- **`MeasureProcedure` qualifies an unqualified procedure name with
  `ThisWorkbook`.** `Application.Run` previously resolved against the active
  workbook, so the harness could measure a same-named procedure in whichever
  workbook the user had in front. An explicit qualification is honored
  unchanged. (#9)

- **The checkpoint usage example no longer sets the run label before
  `StartTimer`.** Starting a session clears the label, so the shipped example
  produced an unlabelled report while appearing to set one. It now passes the
  label to `StartTimer`. The examples module also had no coverage of anything
  added in v1.2.0. (#17)

### Fixed

- **`Stats_IsContaminated` rejects a negative `CvThreshold`.** A negative
  threshold reported every sample set as contaminated, since the coefficient of
  variation is never negative — a silently useless answer rather than an
  error. (#8)

- **`LastReadStatus` now describes the failed read after a strict-mode error.**
  Strict mode raised from inside the private reader, before the status could be
  returned, and the public operation had already reset `LastReadStatus` to
  `cPM_ReadOK` — so a caller who trapped the error and inspected the status was
  told the last read succeeded.

  The error was never silent; `Err.Number` carried the right condition. But the
  two public surfaces disagreed, and the one documented as authoritative was the
  one that lied. Read policy now lives in `ElapsedSeconds` and `Checkpoint`,
  which publish the status and then decide whether to fail — the shape
  `StartTimer` already used. (CPM120-P2-01)

- **The measurement harness no longer stores a failed read as a zero sample.**
  In non-strict mode a failed endpoint read returns zero, and the harness
  recorded it. That zero then became the minimum of the sample set, so a rare
  native failure was indistinguishable from a genuinely fast iteration — and on
  small samples it moved the median too.

  Failed reads are now **excluded** from the vector rather than stored, so
  `Stats_Count` may be less than `Iterations` and the shortfall is itself the
  measured-failure count. A run in which every read failed raises
  `ERR_CPM_MEASURE_NO_VALID_SAMPLES`, because an empty vector is not a
  measurement. Warm-up failures are counted too, so `FailedReadsOut` can exceed
  the shortfall. (CPM120-P2-02)

- **A clamped negative elapsed no longer passes as a valid measurement.**
  `Elapsed_Validate` clamped a backwards-moving timing source to zero and
  returned only a number, so the result was "zero seconds, status OK" — which is
  indistinguishable from a genuine measurement of zero.

  A backward-clock event could therefore be recorded as a valid checkpoint, or
  kept as the fastest sample in a benchmark. The clamp now reports
  `cPM_ReadElapsedInvalid`, so every caller that already handles a failed read
  handles this too: `ElapsedSeconds` raises in strict mode or returns zero with
  the status, `Checkpoint` abandons the capture, and the harness excludes the
  sample. (CPM120-P2-03)

- **The statistics contract no longer promises more than it delivers.** The
  routines were documented as accepting any `Double()` vector while the
  coefficient of variation already assumed positive timing data — a contract
  broader than the implementation, which invites callers to rely on behavior
  that was never designed.

  The supported domain is now stated as **finite, non-negative timing
  observations**, and anything else is rejected at a single gate naming the
  offending index and reason. Not because `Stats_Min` breaks on a negative
  number — it does not — but because a negative value means the vector is not a
  timing vector, so every figure computed from it would describe something the
  caller did not intend. Zero remains in domain: an operation can genuinely be
  too fast to resolve. (CPM120-P2-04)

- **The trusted-input boundary of the measurement harness is now documented.**
  `MeasureProcedure` passes its argument to `Application.Run`, which executes
  it, so a name taken from a worksheet cell or a file is arbitrary code
  execution inside the project. Qualification, added earlier in this release,
  narrows a different risk and does not address this one.

  No code change: any syntactically valid procedure name is executable by
  design, so validating the string would look like protection without being any.
  The remaining criteria in this finding were already met by the qualification
  work. (CPM120-P2-05)

- **The 32-bit tick-count path no longer carries its own arithmetic.** Both
  32-bit branches of `Get_TickCountSeconds` had the signed-to-unsigned
  correction written out inline — the same logic as `UInt32ToDouble`, in code
  that cannot execute during a 64-bit test run. They now call it.

  The correction is verified at all four boundary values on whichever bitness
  the suite runs, so what remains bitness-specific is the declaration binding,
  the API return and the `RolloverSeconds` constant. A materially smaller
  untested surface, though not zero: 32-bit execution evidence is now a
  documented step for major releases. `UInt32ToDouble` also states its caller
  contract explicitly, since passing it a genuinely signed value returns a
  plausible-looking number near 4.29e9 and the misuse cannot be detected from
  the value. (CPM120-P2-06)

- **The static checks now publish a machine-readable result document.**
  `tools/vba_lint.py --json` emits the commit, branch, working-tree cleanliness,
  and a per-check pass/fail record; the workflow uploads it as an artifact on
  every run, including failing ones, where the detail is most useful.

  This closes the last genuinely missing control from the finding. The other
  three were already in place or are deliberate: the workflow runs on every push
  and pull request, and branch protection restricts deletions but does not
  require the check, because gating every commit meant a branch and a pull
  request for a one-line documentation edit.

  That trade is defensible for daily work and wrong at release time, since a tag
  could otherwise be created against a commit whose checks never ran.
  `RELEASING.md` therefore requires confirming the run for the exact tag target
  and recording its URL. (CPM120-P2-07)

- **The release manifest now verifies its own claim and states its limits.**
  `release_provenance.py --tag` compares every hashed source against that tag's
  blob and exits non-zero on any difference, so the manifest can no longer hash
  a working tree while claiming a commit the two disagree about. `--out` writes
  it as JSON for attachment to the release.

  What it deliberately does not do is assert that the workbook was built from
  that source. No automated step produces the workbook — modules are imported by
  hand and the file saved — and the VBA editor reformats on import, stripping
  alignment and appending lines, so extracted source could never match the
  repository byte for byte even in principle. Adding a "build script version"
  field to a manual process would have been fiction.

  The manifest therefore says plainly which claims it establishes and which it
  does not, and states that the tagged source is authoritative and the workbook
  a convenience copy. (CPM120-P2-08)

### Known limitations

- **`LastReadStatus` does not cover harness reads.** The harness measures on an
  isolated worker released before the vector is returned, so the caller's own
  status describes only its own direct reads. Harness outcomes arrive through
  the `FailedReadsOut` and `LastFailureStatusOut` arguments instead. This is by
  design rather than an outstanding gap, but it is easy to reach for the wrong
  surface.

---

## [1.2.0] — 2026-08-16

Correctness hardening across the native timing reads, a distribution-aware
measurement harness, and the removal of every magic error number from the
codebase.

**Regression suite: 63 cases, 431 assertions, 0 failures.**
Certified 2026-08-16 on Excel for Microsoft 365 MSO Version 2606
Build 16.0.20131.20152, 64-bit.

### Added

#### Measurement and statistics

- **`MeasureProcedure(Name, [Iterations], [Warmup], [Method])`** — runs a named
  `Public Sub` repeatedly through `Application.Run` and returns the full per-run
  elapsed vector as `Double()`.
- **`MeasureOverhead_Samples([Iterations], [Warmup], [Method])`** — per-cycle
  backend overhead as a sample vector. Its minimum is the **observed
  empty-cycle floor**: it contains the start read, the end read, class dispatch
  and loop overhead, and is therefore not the clock's own granularity.
- **Statistics surface** on any `Double()` vector: `Stats_Min`, `Stats_Max`,
  `Stats_Mean`, `Stats_Median`, `Stats_Percentile`, `Stats_StdDev`,
  `Stats_CoefficientOfVariation`, `Stats_IsContaminated`, `Stats_Count`,
  `Stats_Text`.
- **Contamination detection** — `Stats_IsContaminated` reports when the
  coefficient of variation exceeds 0.25, and `Stats_Text` appends a warning so a
  noisy run cannot be mistaken for a precise one.

#### Read-status reporting

- **`cPM_ReadStatus` enum and the `LastReadStatus` property.** In non-strict
  mode a failed read returns 0; without a status, a caller could not distinguish
  "zero seconds elapsed" from "no reading was obtained". Statuses: `cPM_ReadOK`,
  `cPM_ReadQpcFailed`, `cPM_ReadSystemTimeFailed`,
  `cPM_ReadSystemTimeFormatInvalid`, `cPM_ReadFallbackToMethod2`.
- **`TW_CalculationExempted`** and **`PM_TW_CalculationExempted`** — report when
  Calculation control could not be honoured on the current host.

#### API

- **Optional `RunLabel` argument on `StartTimer`**, so labelling a run is a
  single atomic call. Calling `SetRunLabel` *before* `StartTimer` is discarded
  by the session reset.

#### Testing

- **Three one-shot fault-injection seams** — `Test_ForceNextQPCReadFailure`,
  `Test_ForceNextSystemTimeReadFailure`,
  `Test_ForceNextSystemTimeFormatInvalid`. Each is consumed by the read it
  affects, so a forced failure cannot leak into a later case.
- **22 regression cases**, taking the suite from 41 to **63**. New coverage:
  injected native-read failures on both backends, an injected wrong-format
  `timeGetSystemTime` result, Calculation baseline lifecycle, checkpoint
  capacity growth across `ReDim Preserve` boundaries, `Class_Terminate` cleanup,
  and 75 instance create/destroy cycles.

### Changed

- **Backend dispatch consolidated** into a single private reader,
  `Elapsed_ComputeSeconds`.
- **`QPC_TryReadTick` and `SystemTime_TryReadMs` are now the only places that
  touch the native timing APIs.** Both report success as a return value and
  write their output parameter only on success.
- **`StartTimer` is transactional.** It captures through `Start_TryCapture` and
  commits nothing until a valid start timestamp exists. No session is committed
  whose start value was not successfully read.
- **`Checkpoint` no longer overwrites the cached `T2` and `ET`.** A failed read
  abandons the capture entirely, leaving the checkpoint count and the
  last-checkpoint marker unchanged.
- **All 19 inline `vbObjectError` literals replaced by named constants.** No
  bare offset appears anywhere in the code. Timing-method identifiers 1–6 remain
  numeric literals; tracked in
  [#13](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/13).
- **`UInt32ToDouble`, `Elapsed_Validate` and `RolloverSeconds` are `Friend`**
  rather than `Private`, so the suite can test their arithmetic directly.
- **`Class_Initialize` delegates to `Checkpoint_ClearState`.**
- **Report column positions are named constants** (`PM_RPT_COL_DELTA`,
  `PM_RPT_COL_CUMULATIVE`).
- The demo workbook is distributed as a **Release asset** rather than versioned
  in the repository.

### Fixed

- **A failed QPC start read no longer commits a zero-based session.** Non-strict
  mode previously committed method 5 with a QPC start of zero; a subsequent
  successful end read then measured from zero and reported a value resembling
  machine uptime. Non-strict now binds method 2 *before* any state is committed
  and records `cPM_ReadFallbackToMethod2`; strict raises with the previous
  session untouched.
- **A failed method-4 read no longer becomes a large positive duration.**
  `Get_SystemTimeMs` returned 0 on failure; that sentinel produced a negative
  delta which rollover correction then "fixed" by adding the 32-bit wrap period,
  yielding roughly 4,294,967 seconds. Every read is now validated before any
  subtraction or rollover correction.
- **A failed elapsed read no longer destroys the last known good measurement.**
- **Shared-TW instance keys no longer derive from `ObjPtr(Me)`.** VBA reuses
  heap addresses, so a session left by a destroyed instance could be silently
  inherited by an unrelated instance at the same address.
- **`TW_Turn_OFF` commits `m_Except` only after `PM_TW_BeginSession` succeeds.**
- **The Calculation baseline no longer conflates "unknown" with "Automatic".**
  A scope beginning with no workbook stored a synthetic
  `xlCalculationAutomatic`, which a later apply could write over a real
  workbook's setting. `g_TW_CALCULATION_VALID` now tracks whether a baseline was
  genuinely captured, separately from its value.
- **Negative elapsed times are no longer returned silently.** On Win64
  backend 2, `RolloverSeconds` returns 0, so a negative value passed through
  unchanged. Strict raises `ERR_CPM_NEGATIVE_ELAPSED`; non-strict clamps to
  zero. Backend 5 skips the check deliberately.
- **`Pause` methods 3 and 4 no longer overshoot by up to a second.** `Now` and
  `Application.Wait` are second-granular and `1/86400` is not exactly
  representable, so a coarse loop could wait a full extra tick. Both coarse
  loops are now bounded by the same elapsed measurement the trailing guard uses.
  Consequence: below roughly two seconds, method 3 issues no coarse wait and
  behaves like method 2.
- **Error sources naming `M_cPM_ReportHelpers` and `M_cPM_RegressionTests`** —
  modules that have never existed — corrected.

### Removed

- **`QPC_TryGetNextTimestampTick`** — around 100 lines of dead code, never
  called, and a near-duplicate of the live alignment helper.

### Documentation

- `TW_InstanceKey`'s header rewritten; it documented the removed `ObjPtr`
  approach in full.
- Four procedure header titles corrected where they did not match their names.
- New header sections: **Elapsed-time validation policy**, **Instance key
  policy**, **Host-state policy**, **Calculation baseline validity**,
  **Stable-host invariant**, **Error numbers**.
- README and wiki claims qualified where they overstated what the
  implementation guarantees: backwards-clock detection scoped to the backends
  that support it, "resolution floor" renamed to "observed empty-cycle floor",
  the unsupported `Application.Run` cost figure removed, and the incorrect
  baseline-subtraction guidance replaced. A dispatch-matched baseline helper is
  tracked in
  [#7](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/7).
- The test module's `M_DEMO_BUILDER` dependency documented.

### Known limitations

- **Calculation control requires a stable open-workbook set** for the life of a
  suppression scope. The genuinely workbook-less paths cannot be exercised by a
  suite running inside a workbook and remain manually verified only.
- **`MeasureOverhead_Samples` is not a matched baseline for
  `MeasureProcedure`**, because it does not dispatch through `Application.Run`.
- **Backend 1 cannot distinguish a backward clock adjustment from midnight
  rollover.**
- **Rollover correction handles a single wrap.**

---

## [1.1.0] — 2026-04-18

### Added

- Structured checkpoint capture with delta and cumulative timings.
- Per-session run labels via `SetRunLabel`.
- `CheckpointCount`, `ClearCheckpoints`, `ReportAsArray`, `ReportAsText`.
- `cPM_Report_WriteToRange` for direct worksheet output.
- Expanded regression coverage for checkpoint and reporting behavior.

### Changed

- Documentation refresh across README and wiki pages.

---

## [1.0.0] — 2026-03-28

Initial public release.

### Added

- Session-bound timing model with six selectable backends.
- Validation-aware elapsed-time reads bound to the session backend.
- Formatted elapsed-time output, including formatting an already measured value
  without taking a second sample.
- Structured diagnostics for QPC frequency and system tick interval.
- Benchmark-overhead estimation helpers.
- `Pause` with four selectable strategies.
- Explicit cleanup semantics through `ResetEnvironment`, with `Class_Terminate`
  as the safety net.
- Shared time-waster suppression through the required companion module.
- Demo, usage-example, and regression-test modules.

---

[Unreleased]: https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/compare/v1.4.0...HEAD
[1.4.0]: https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.3.0
[1.2.0]: https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.2.0
[1.1.0]: https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.1.0
[1.0.0]: https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.0.0
