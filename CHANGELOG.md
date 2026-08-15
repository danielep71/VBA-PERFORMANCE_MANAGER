# Changelog

All notable changes to **Class Performance Manager** are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.2.0] — 2026-08-14

Correctness hardening, a distribution-aware measurement harness, and the removal
of every magic error number from the codebase.

### Added

- **`MeasureProcedure(Name, [Iterations], [Warmup], [Method])`** — runs a named
  `Public Sub` repeatedly through `Application.Run` and returns the full per-run
  elapsed vector as `Double()`. Warm-up iterations are discarded before
  measurement begins.
- **`MeasureOverhead_Samples([Iterations], [Warmup], [Method])`** — per-cycle
  backend overhead as a sample vector. Its minimum is the practical resolution
  floor of the selected backend.
- **Statistics surface** operating on any `Double()` vector: `Stats_Min`,
  `Stats_Max`, `Stats_Mean`, `Stats_Median`, `Stats_Percentile`, `Stats_StdDev`,
  `Stats_CoefficientOfVariation`, `Stats_IsContaminated`, `Stats_Count` and
  `Stats_Text`.
- **Contamination detection** — `Stats_IsContaminated` reports when the
  coefficient of variation exceeds 0.25, and `Stats_Text` appends an explicit
  warning line so a noisy run cannot be mistaken for a precise one.
- **Optional `RunLabel` argument on `StartTimer`**, so labelling a run is a
  single atomic call.
- **`PM_TW_NewInstanceKey`** in `M_cPM_TimeWasters` — issues collision-proof
  shared-TW registration keys.
- **`Elapsed_Validate`** — validates elapsed values after rollover correction.
- **11 regression cases**, taking the suite from 41 to 52 and 288 assertions.

### Changed

- **Backend dispatch consolidated** into a single private reader,
  `Elapsed_ComputeSeconds`, replacing duplicated `Select Case` blocks.
  Behaviour is unchanged.
- **`Checkpoint` no longer overwrites the cached `T2` and `ET` values.** It
  measures through the non-mutating reader, so an explicit `ElapsedSeconds`
  result survives any number of subsequent checkpoints. It also no longer pays
  the `Me.ElapsedSeconds` vtable cost.
- **All 19 inline `vbObjectError` literals replaced by named constants.** The
  codebase now contains no magic error numbers.
- **`UInt32ToDouble`, `Elapsed_Validate` and `RolloverSeconds` are `Friend`**
  rather than `Private`, so the regression suite can test their arithmetic
  directly. They remain invisible outside the project.
- **`Class_Initialize` delegates to `Checkpoint_ClearState`** instead of
  duplicating it, leaving one definition of a clean checkpoint baseline.
- **Report column positions are named constants** (`PM_RPT_COL_DELTA`,
  `PM_RPT_COL_CUMULATIVE`) rather than the literals 7 and 8.
- **`ReportAsArray` and `ReportAsText`** resolve each checkpoint row once
  through `With`.
- The demo workbook is now distributed as a **Release asset** rather than
  versioned in the repository.

### Fixed

- **Shared-TW instance keys no longer derive from `ObjPtr(Me)`.** VBA reuses
  heap addresses, so a session left behind by a destroyed instance could be
  silently inherited by an unrelated instance later allocated at the same
  address. Keys now come from a module-level counter that shares its lifetime
  with the session store.
- **`TW_Turn_OFF` commits `m_Except` only after `PM_TW_BeginSession` succeeds.**
  A failed registration previously left the instance reporting a suppression
  mask it had never applied.
- **Negative elapsed times are no longer returned silently.** On Win64
  backend 2, `RolloverSeconds` returns 0 because `GetTickCount64` is monotonic,
  so a negative value passed through the correction step unchanged. Strict mode
  now raises `ERR_CPM_NEGATIVE_ELAPSED`; non-strict mode clamps to zero.
  Backend 5 skips the check deliberately.
- **`Application.Calculation` access is guarded on `Workbooks.Count`.** Both the
  baseline capture and the effective-state apply paths raised in a workbook-less
  host.
- **Error sources naming `M_cPM_ReportHelpers`** — a module that has never
  existed — corrected to `M_cPM_TimeWasters`.
- **Error sources naming `M_cPM_RegressionTests`** in the test module corrected
  to `M_cPM_Test`.

### Removed

- **`QPC_TryGetNextTimestampTick`** — approximately 100 lines of dead code,
  never called from the class, tests, or demos, and a near-duplicate of
  `QPC_Get_NextTimestampTick`.

### Documentation

- `TW_InstanceKey`'s header rewritten; it documented the removed `ObjPtr`
  approach in full.
- Four procedure header titles corrected where they did not match their
  procedure names.
- `T2` and `ET` headers now record that `Checkpoint` updates them.
- `Pause`'s `iMethod` documented as a pause strategy, not a timing backend.
- New header sections: **Elapsed-time validation policy**, **Instance key
  policy**, **Host-state policy** and **Error numbers**.
- The test module's `M_DEMO_BUILDER` dependency is documented, having been
  unlisted despite 43 calls to `Demo_SB_SetProgress` alone.

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

[1.2.0]: https://github.com/danielep71/vba-performance_manager/releases/tag/v1.2.0
[1.1.0]: https://github.com/danielep71/vba-performance_manager/releases/tag/v1.1.0
[1.0.0]: https://github.com/danielep71/vba-performance_manager/releases/tag/v1.0.0
