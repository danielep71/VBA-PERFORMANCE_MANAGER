# Changelog

All notable changes to **Class Performance Manager** are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

Work merged to `main` but not yet released. Each entry is added by the pull
request that introduces it, so the record is written while the reasoning is
fresh rather than reconstructed at release time.

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

### Changed

- **BEHAVIOURAL CORRECTION: `Stats_CoefficientOfVariation` now raises when the
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

### Fixed

- **`Stats_IsContaminated` rejects a negative `CvThreshold`.** A negative
  threshold reported every sample set as contaminated, since the coefficient of
  variation is never negative — a silently useless answer rather than an
  error. (#8)

### Known limitations

- **Samples returned by `MeasureProcedure` cannot be diagnosed through
  `LastReadStatus`.** The harness runs its reads on an isolated worker instance
  that is released before the vector is returned, so the caller's own status is
  untouched by them. Tracked separately.

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
- **Release provenance.** `tools/release_provenance.py` emits a SHA-256 for every
  shipped file and Release asset, the commit the release was cut from, and the
  Excel build and bitness the suite was certified on. Missing fields are marked
  `TODO` and the script exits non-zero, so an incomplete block cannot ship
  unnoticed. `RELEASING.md` records the process, including the three mechanical
  failures that preceded a correct v1.2.0 tag. (#12)
  

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
  [#13](https://github.com/danielep71/vba-performance_manager/issues/13).
- **`UInt32ToDouble`, `Elapsed_Validate` and `RolloverSeconds` are `Friend`**
  rather than `Private`, so the suite can test their arithmetic directly.
- **`Class_Initialize` delegates to `Checkpoint_ClearState`.**
- **Report column positions are named constants** (`PM_RPT_COL_DELTA`,
  `PM_RPT_COL_CUMULATIVE`).
- The demo workbook is distributed as a **Release asset** rather than versioned
  in the repository.
- **`MeasureProcedure` qualifies an unqualified procedure name with
  `ThisWorkbook`.** `Application.Run` previously resolved against the active
  workbook, so the harness could measure a same-named procedure in whichever
  workbook the user had in front. An explicit qualification is honored
  unchanged. (#9)

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
  [#7](https://github.com/danielep71/vba-performance_manager/issues/7).
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
- Expanded regression coverage for checkpoint and reporting behaviour.

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

[Unreleased]: https://github.com/danielep71/vba-performance_manager/compare/v1.2.0...main
[1.2.0]: https://github.com/danielep71/vba-performance_manager/releases/tag/v1.2.0
[1.1.0]: https://github.com/danielep71/vba-performance_manager/releases/tag/v1.1.0
[1.0.0]: https://github.com/danielep71/vba-performance_manager/releases/tag/v1.0.0
