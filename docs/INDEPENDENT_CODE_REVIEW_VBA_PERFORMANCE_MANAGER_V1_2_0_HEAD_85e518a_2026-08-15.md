# Independent Code Review — VBA Performance Manager v1.2.0

> \*\*Repository:\*\* \[`danielep71/VBA-PERFORMANCE\_MANAGER`](https://github.com/danielep71/VBA-PERFORMANCE\_MANAGER)  
> \*\*Branch reviewed:\*\* \[`release/v1.2.0`](https://github.com/danielep71/VBA-PERFORMANCE\_MANAGER/tree/release/v1.2.0)  
> \*\*Commit reviewed:\*\* \[`85e518aed2b853f6e9134cf4d0dbb608f91ec408`](https://github.com/danielep71/VBA-PERFORMANCE\_MANAGER/commit/85e518aed2b853f6e9134cf4d0dbb608f91ec408)  
> \*\*Current `main`:\*\* \[`b6e59486eed1bd812d2df8e7a36cfac174c51582`](https://github.com/danielep71/VBA-PERFORMANCE\_MANAGER/commit/b6e59486eed1bd812d2df8e7a36cfac174c51582)  
> \*\*Review date:\*\* 2026-08-15   
> \*\*Suggested repository path:\*\* `docs/INDEPENDENT\_CODE\_REVIEW\_V1.2.0\_2026-08-15.md`

\---

## 1\. Executive assessment

### Overall repository score: **8.2 / 10**

### Production implementation score: **8.6 / 10**

### Regression-suite design score: **8.7 / 10**

### Documentation score: **8.8 / 10**

### CI and release-engineering score: **5.7 / 10**

The current `release/v1.2.0` branch is materially stronger than the revision reviewed earlier on 2026-08-15.

Since commit `c59100f4c1b8cf089dc776a4bb10d58a4040be04`, the branch has added:

* a repeated-measurement harness;
* a public descriptive-statistics surface;
* a single-source elapsed-time backend reader;
* non-mutating checkpoint measurements;
* eleven additional regression cases;
* updated test metadata for version 1.2.0;
* a substantially rewritten README;
* a structured CHANGELOG;
* a committed independent review under `docs/`.

The class source has grown from approximately **153 KB to 203 KB**, and the test module from approximately **238 KB to 299 KB**. The regression inventory has grown from **41 to 52 cases**, with the README reporting **288 assertions**.

The core engineering direction is good:

* the public timing session remains coherent;
* QPC remains the right default;
* the new single-site elapsed reader reduces behavioral drift;
* checkpoint capture no longer mutates the caller-visible elapsed cache;
* the statistics and raw-sample APIs move the project beyond one-number benchmarking;
* the new tests cover substantially more of the pure arithmetic and lifecycle surface;
* source documentation is unusually detailed for VBA.

The repository is nevertheless **not yet ready for an unconditional v1.2.0 release**.

Two release-blocking correctness defects from the earlier review remain:

1. **A failed native timestamp read can still be converted into numeric zero and committed or consumed as if it were a valid timestamp in non-strict mode.**
2. **The workbook-less `Application.Calculation` guard still substitutes a synthetic `xlCalculationAutomatic` baseline and does not correctly model workbook-open/workbook-close transitions during an active suppression scope.**

The expanded measurement/statistics surface also creates new contract work:

* `MeasureOverhead\_Samples` does not reproduce the `Application.Run` dispatch cost included by `MeasureProcedure`, so it is not a valid full harness baseline for very short procedures;
* unqualified procedure names can resolve ambiguously when multiple workbooks or add-ins expose the same macro name;
* the statistics functions are documented as accepting any `Double()` vector, but their arithmetic is only robust for ordinary nonnegative timing samples of moderate magnitude;
* coefficient-of-variation semantics are incorrect for zero-mean or negative-mean general vectors;
* the README and CHANGELOG contain several exact contradictions and overclaims.

Finally, the assurance layer remains manual:

* no GitHub Actions workflows;
* no self-hosted Excel compile/test gate;
* no required status checks;
* no branch protection;
* no open release pull request;
* no `v1.2.0` tag or release;
* no source-to-`.xlsm` provenance manifest;
* the release branch is currently **11 commits ahead and 1 commit behind `main`**.

### Independent verdict

> \*\*v1.2.0 is a strong release candidate, not a release-ready build. The project can exceed 9/10 without redesigning the public class, but only after native-read failure semantics, Calculation-state transitions, executable test evidence, and release governance are closed.\*\*

\---

# 2\. What is required to reach 9+

The shortest credible route from **8.2** to **9+** is:

## Mandatory 9+ gates

1. **Eliminate zero-as-failure timestamp semantics**

   * use status-bearing native readers;
   * never commit a session after a failed start read;
   * never run rollover arithmetic on a failed end read;
   * add deterministic injected-failure tests.
2. **Define and implement a correct Calculation-state transition policy**

   * track whether the Calculation baseline was actually captured;
   * handle or explicitly reject workbook-set changes while a TW scope is active;
   * test begin-without-workbook, open-during-scope, close-before-restore, and final cleanup.
3. **Create a headless Excel regression gate**

   * import the required source;
   * compile;
   * run the core suite;
   * emit machine-readable counters;
   * fail CI on compile/runtime/assertion failure;
   * test Office 64-bit and obtain at least periodic Office 32-bit evidence.
4. **Make the release source-bound**

   * update the branch from `main`;
   * open a release PR;
   * require green checks;
   * tag the exact tested commit;
   * build the release workbook from that commit;
   * publish source and artifact SHA-256 values plus Excel version/build/bitness.
5. **Correct the measurement/statistics contract**

   * provide a dispatch-matched empty-procedure baseline;
   * qualify procedure resolution;
   * validate the statistical domain or use robust arithmetic;
   * fix CV semantics and threshold validation;
   * correct the related README guidance.

Completing Gates 1–4 should place the repository at approximately **9.0–9.1**. Gate 5 and the modularity work described later should move it toward **9.2–9.4**.

\---

# 3\. Review scope and methodology

## 3.1 Exact revision basis

The review is tied to:

```text
Branch: release/v1.2.0
Commit: 85e518aed2b853f6e9134cf4d0dbb608f91ec408
Date:   2026-08-15
```

The branch currently compares with `main` as:

```text
Ahead:  11 commits
Behind: 1 commit
Status: diverged
```

The one commit present on `main` but not the release branch removes framework-positioning sections from the README. That divergence should be resolved before opening or merging the release PR.

## 3.2 Production source reviewed

* [`src/classes/cPerformanceManager.cls`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/85e518aed2b853f6e9134cf4d0dbb608f91ec408/src/classes/cPerformanceManager.cls)
* [`src/modules/M\_cPM\_TIMEWASTERS.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/85e518aed2b853f6e9134cf4d0dbb608f91ec408/src/modules/M_cPM_TIMEWASTERS.bas)

## 3.3 Regression source reviewed

* [`test/M\_cPM\_Test.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/85e518aed2b853f6e9134cf4d0dbb608f91ec408/test/M_cPM_Test.bas)

## 3.4 Supporting material reviewed

* [`README.md`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/85e518aed2b853f6e9134cf4d0dbb608f91ec408/README.md)
* [`CHANGELOG.md`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/85e518aed2b853f6e9134cf4d0dbb608f91ec408/CHANGELOG.md)
* `.gitignore`
* demo/source directory structure;
* branch topology and protection state;
* GitHub Actions inventory;
* tags and releases;
* open issues and pull requests;
* current release-asset metadata.

## 3.5 Execution boundary

Desktop Excel was not available in the review environment.

The reviewer therefore did **not**:

* import the source into the VBE;
* execute `Debug -> Compile VBAProject`;
* run `Run\_cPerformanceManager\_RegressionSuite`;
* independently verify the reported 288 assertion count;
* dynamically force Windows API failures;
* test workbook-open/workbook-close transitions;
* test Office 32-bit declarations;
* measure benchmark distributions on physical Windows hardware;
* rebuild the `.xlsm` release artifact.

The review distinguishes between:

* source-confirmed control flow;
* repository metadata;
* source-declared test coverage;
* runtime behavior requiring Excel confirmation.

One initially suspected test-syntax issue was checked against the Microsoft VBA language specification and **not** counted as a defect: indexing `cPM.ReportAsArray(3, 8)` is valid because `ReportAsArray` is a zero-argument function returning `Variant`, which can be followed by an index argument list.

\---

# 4\. Hard repository metrics

## 4.1 Required production files

|File|SHA|Size|
|-|-|-:|
|`cPerformanceManager.cls`|`f3c669f71cb3a7e54dcef226faabd85458f19514`|202,953 bytes|
|`M\_cPM\_TIMEWASTERS.bas`|`a45cdf03ce06d94f7fce8f8f5392cd5bebb80d98`|46,666 bytes|
|**Required production total**||**249,619 bytes**|

## 4.2 Regression source

|Artifact|Current state|
|-|-:|
|Test module|`M\_cPM\_Test.bas`|
|SHA|`23f4ed49e57734ad58481f8cec6dc6ab6541448d`|
|Size|298,979 bytes|
|Version header|1.2.0|
|Registered cases|52|
|Assertions|288, repository-reported|
|Excel execution result|Not committed as an authoritative CI artifact|

## 4.3 Public surface

The README reports **41 public members**.

The logical surface now includes:

```text
Core timing and formatting
Session/state inspection
Diagnostics
Pause and cleanup
TW lifecycle
Checkpoint/reporting
Repeated measurement
Descriptive statistics
```

The README line:

```text
41 public members — 24 methods and 18 properties
```

contains an arithmetic inconsistency because 24 + 18 = 42. The inventory should be generated from source rather than typed manually.

## 4.4 Repository governance

```text
GitHub Actions workflows: 0
Open issues:              0
Open pull requests:       0
v1.2.0 tag:               absent
v1.2.0 release:           absent
main protected:           no
release branch protected: no
Required status checks:   none
```

\---

# 5\. Improvements since the earlier v1.2.0 review

## 5.1 Repeated-measurement harness

The branch now provides:

```text
MeasureProcedure
MeasureOverhead\_Samples
```

This is an important conceptual improvement.

A benchmark library should return raw observations, not only a pre-aggregated average. Returning a `Double()` sample vector enables:

* independent analysis;
* median and percentile reporting;
* contamination checks;
* cross-run comparison;
* external persistence;
* later methodology changes without rerunning the workload.

## 5.2 Public statistics surface

The new API includes:

```text
Stats\_Count
Stats\_Min
Stats\_Max
Stats\_Mean
Stats\_Median
Stats\_Percentile
Stats\_StdDev
Stats\_CoefficientOfVariation
Stats\_IsContaminated
Stats\_Text
```

The decision to emphasize median, minimum and P95 is directionally sound for right-skewed runtime measurements.

## 5.3 Single-source elapsed reader

Elapsed backend dispatch has been consolidated in:

```text
Elapsed\_ComputeSeconds
```

This reduces the risk that:

* `ElapsedSeconds`;
* `Checkpoint`;
* future measurement APIs

implement subtly different rollover or failure behavior.

## 5.4 Checkpoint cache isolation

`Checkpoint` now measures through the internal reader without overwriting:

```text
m\_T2
m\_ET
m\_qpcEnd
```

That is a good API-semantic correction. A caller that explicitly captured an elapsed result can now take later checkpoints without silently changing the cached diagnostic state.

## 5.5 Regression expansion

The suite now covers, among other things:

* `UInt32ToDouble` boundaries;
* rollover periods;
* strict/non-strict elapsed validation;
* 1,000-checkpoint capacity growth;
* checkpoint cache preservation;
* atomic `StartTimer` run labels;
* terminator cleanup;
* repeated object/session lifecycle;
* known statistics;
* statistics boundaries/errors;
* measurement-harness behavior.

## 5.6 Documentation

The README is now a serious product document rather than a basic repository landing page.

The CHANGELOG materially improves release traceability.

These are clear improvements.

\---

# 6\. Weighted scorecard — with an explicit path to 9+ in every area

|Area|Weight|Current score|Weighted contribution|What must change to reach 9+|
|-|-:|-:|-:|-|
|Functional correctness|18%|**8.2**|1.476|Remove failed-read sentinels from valid timestamp space; add transactional fallback/status semantics and injected failure tests|
|Timing-source robustness|14%|**8.0**|1.120|Define single-wrap limits, Timer backward-clock ambiguity, end-read failure behavior and non-strict status reporting|
|Architecture and modularity|10%|**8.7**|0.870|Extract statistics and worksheet reporting behind stable wrappers; reduce the 203 KB class responsibility set|
|Public API and statistical contract|10%|**8.7**|0.870|Add timer/pause/error enums, machine-readable status, qualified macro resolution and a precise statistics input domain|
|TW lifecycle and host-state correctness|9%|**8.2**|0.738|Track Calculation-baseline validity and define workbook-transition behavior explicitly|
|Regression testing|13%|**8.7**|1.131|Add native/API fault injection, headless execution, 32/64-bit evidence and machine-readable pass/fail artifacts|
|CI and release engineering|8%|**5.7**|0.456|Add hosted static checks, self-hosted Excel CI, branch protection, a release PR and source-bound artifacts|
|Documentation and governance|8%|**8.8**|0.704|Remove contradictions/overclaims, generate metrics, mark 1.2 as unreleased until tagged, document exact limitations|
|Maintainability and repository hygiene|5%|**8.3**|0.415|Split oversized modules, automate API/test inventories, add contribution/security/release policies|
|Performance methodology|5%|**8.4**|0.420|Add a dispatch-matched no-op baseline and publish environment-tagged median/P95/CV evidence|
|**Total**|**100%**||**8.200 / 10**||

\---

# 7\. Component scores — each with its route to 9+

|Component|Score|Assessment|To reach 9+|
|-|-:|-|-|
|`cPerformanceManager.cls`|**8.6**|Rich, coherent facade with strong documentation; native-read semantics and breadth remain issues|Fix P1 failures, expose status, extract pure statistics/native helpers|
|Native timing engine|**8.0**|Six backends, QPC path and rollover helpers are strong; failure sentinel and supported-duration contracts remain|Use `TryRead` APIs, test failures, document maximum interval and clock semantics|
|Measurement harness|**8.4**|Raw sample vectors and warm-up support are valuable|Qualify macro names, cap allocations, add a matching `Application.Run` baseline|
|Statistics layer|**8.4**|Useful timing-oriented API and sensible median/P95 emphasis|Robust summation/variance, safe median, CV domain fix, threshold validation|
|Checkpoints/reporting|**9.0**|Clear semantics, geometric growth, useful exports, cache isolation|Correct stale comments and improve large-report text construction|
|`M\_cPM\_TIMEWASTERS.bas`|**8.3**|Strong aggregate-mask and rollback model|Correct Calculation validity/transitions and separate worksheet reporting|
|Regression suite|**8.7**|Large and thoughtful 52-case source suite|Headless runner, fault injection, CI execution, 32-bit evidence|
|README/CHANGELOG|**8.8**|Highly polished and useful|Remove exact contradictions, generate counts, correct release state and measurement claims|
|Repository/release governance|**5.7**|Source hygiene improved, but release evidence remains manual|Workflows, protected PR, tag provenance, artifact manifest|
|Security/platform posture|**8.5**|Narrow Windows APIs, no production network/shell, explicit cleanup|Document `Application.Run` trust boundary and sign/source-bind release artifacts|

\---

# 8\. Architectural review

## 8.1 Current logical architecture

```text
cPerformanceManager
    |
    +-- session state
    +-- six timing backends
    +-- elapsed validation / rollover
    +-- timer-resolution lifecycle
    +-- pause strategies
    +-- diagnostics
    +-- repeated measurement
    +-- statistics
    +-- checkpoint capture
    +-- report serialization
    +-- TW facade
            |
            v
M\_cPM\_TIMEWASTERS
    |
    +-- process-wide TW session registry
    +-- baseline capture / restore
    +-- aggregate disable masks
    +-- rollback
    +-- worksheet report writer
```

## 8.2 Architectural strengths

### Session-bound clocks

The class prevents starting with one timing source and ending with another.

This remains the most important design distinction between this project and a casual timer helper.

### Shared Application-state ownership

The code correctly treats:

```text
ScreenUpdating
EnableEvents
DisplayAlerts
Calculation
Cursor
```

as Excel-process state rather than object-local state.

### Single elapsed dispatch

The new `Elapsed\_ComputeSeconds` is a good centralization boundary.

### Raw observations before statistics

The measurement API returns samples first and aggregates second.

That is the correct architecture for benchmark transparency.

### Checkpoint storage

The geometric capacity strategy avoids a `ReDim Preserve` on every capture.

## 8.3 Architectural weaknesses

### The facade now owns too many pure/stateless concerns

The statistics routines do not depend on timing-session state.

They are conceptually a separate numerical service.

A compatibility-preserving improvement would be:

```text
M\_cPM\_STATISTICS
    Private/Friend pure implementations

cPerformanceManager
    existing Public Stats\_\* wrappers
```

This preserves all callers while reducing class complexity.

### Reporting remains mixed into the TW manager

`cPM\_Report\_WriteToRange` belongs in:

```text
M\_cPM\_REPORTING.bas
```

not in a module whose primary responsibility is process-wide state restoration.

### Testability changes leak into production visibility

Three helpers are now `Friend` specifically so tests can invoke them.

That is acceptable, but a cleaner architecture would place pure arithmetic in an internal module and test it there, or use a conditional test facade.

### Path to 9+

To score above 9 in architecture:

1. keep the public class unchanged;
2. extract pure statistics;
3. extract worksheet report writing;
4. consider extracting low-level native readers behind a narrow internal API;
5. generate a dependency diagram and public API inventory in CI.

No public rewrite is needed.

\---

# 9\. Detailed production-code review

## 9.1 Initialization and lifecycle — **8.9 / 10**

### Strengths

* deterministic state initialization;
* one-time QPC frequency capture;
* explicit method-name table;
* centralized checkpoint reset;
* `Class\_Terminate` used only as best-effort fallback;
* explicit `ResetEnvironment` remains primary;
* timer-resolution and TW cleanup are coordinated.

### Remaining limitation

A destructor cannot surface cleanup failure.

If final TW restoration fails during `Class\_Terminate`, the error is suppressed and shared state may remain active.

The global recovery function is therefore necessary.

### To reach 9+

* expose a lightweight diagnostic such as `LastCleanupStatus`;
* test teardown-failure behavior through injection;
* document that deterministic cleanup requires `ResetEnvironment`, not garbage collection.

\---

## 9.2 `StartTimer` — **7.9 / 10**

### What is good

`StartTimer` captures into local variables before committing:

```text
method
start timestamp
QPC start tick
```

The optional `RunLabel` is applied after checkpoint-state reset.

This is correct sequencing.

### Release-blocking defect: QPC failure in non-strict mode

When QPC is selected and the initial read fails:

```vb
If m\_StrictMode Then
    Err.Raise ...
End If
NewQpcStart = 0@
```

The function then commits:

```text
active method = 5
active session = true
QPC start = 0
```

If the next QPC read succeeds, elapsed time is computed from zero rather than from the requested start instant.

This can report a duration resembling QPC uptime rather than operation runtime.

The same issue exists in the non-strict aligned-start path when its initial QPC read fails.

### Correct remediation

Use:

```vb
Private Function QPC\_TryReadTick(ByRef TickOut As Currency) As Boolean
```

as a real status contract.

For a failed **start** read:

```text
strict:
    raise; preserve previous session

non-strict:
    fall back to method 2;
    capture a valid method-2 start;
    commit method 2, not method 5
```

Do not commit any session whose start value was not successfully read.

### To reach 9+

* implement fallback-before-commit;
* add injected QPC failure tests;
* assert previous session preservation in strict mode;
* assert resolved backend = 2 in non-strict fallback;
* expose fallback status to the caller.

\---

## 9.3 `ElapsedSeconds` and `Elapsed\_ComputeSeconds` — **8.2 / 10**

### Strengths

* one backend dispatch site;
* session/method resolution before read;
* QPC return-code checking;
* unsigned 32-bit handling;
* backend-specific rollover;
* cached endpoint and elapsed values only in the public elapsed path;
* checkpoint calls the same arithmetic without mutating caches.

### Method-4 failure defect

`Get\_SystemTimeMs` returns `0` in non-strict mode when:

* the API call fails; or
* the returned format is not `TIME\_MS`.

The elapsed reader then performs:

```text
end = 0
delta = end - start
if delta < 0:
    delta += 2^32 milliseconds
```

A failed read can therefore become a large **positive** elapsed value and pass validation.

At session start, a method-4 failure can similarly commit zero, after which a successful end read produces an uptime-like value.

### Design rule

> A failed read must never enter timestamp arithmetic.

### Recommended API

```vb
Private Function SystemTime\_TryReadMs(ByRef ValueOut As Double) As Boolean
```

At elapsed end, the class cannot safely switch clocks because the session began on method 4.

Recommended behavior:

```text
strict:
    raise

non-strict:
    return 0 plus explicit LastStatus;
    do not apply rollover;
    do not overwrite the last successful endpoint with failure sentinel data
```

### To reach 9+

* status-bearing method-4 reads;
* no failure sentinel arithmetic;
* fault-injection tests for start/end/format failures;
* checkpoint behavior defined when an elapsed read fails;
* machine-readable failure status.

\---

## 9.4 Timer and rolling-counter domains — **8.0 / 10**

The implementation corrects one wrap.

Therefore the safe single-session intervals are approximately:

```text
Timer:                 less than 24 hours
GetTickCount, Win32:   less than 49.71 days
timeGetTime:           less than 49.71 days
timeGetSystemTime:     less than 49.71 days
GetTickCount64:        practically unbounded for normal use
QPC:                   preferred long-duration source
```

### Timer ambiguity

For backend 1, every negative raw delta is treated as midnight rollover.

A backward wall-clock adjustment during the day is therefore indistinguishable from midnight and can become a value close to 24 hours.

The general README claim of “backwards-clock detection” should be qualified.

### Multiple wraps

A session spanning more than one full 32-bit millisecond wrap cannot be recovered by a single addition.

The result may be positive and plausible, so validation cannot detect it.

### To reach 9+

* publish the supported interval per backend;
* recommend QPC or GetTickCount64 for long sessions;
* qualify Timer’s rollover/backward-clock ambiguity;
* include these limits in `MethodName` documentation and the README;
* add boundary arithmetic tests around one wrap;
* do not claim universal backwards-clock detection.

\---

## 9.5 `ElapsedTime` — **8.4 / 10**

### Strengths

* can format a premeasured value;
* avoids a second timing sample;
* does not wrap hours at 24;
* sub-second groups are deterministic;
* nanosecond text is correctly documented as presentation precision.

### Large-input boundary

The formatter narrows whole seconds to a VBA `Long`:

```vb
WholeS = CLng(Fix(Sec))
```

This creates an undocumented upper limit around 2.147 billion seconds, or about 68 years.

This is not an ordinary benchmark duration, but the function publicly accepts a `Variant` numeric value and does not document the boundary.

### To reach 9+

* keep whole seconds in `Double` or `Currency`;
* validate `ElapsedSecondsIn` explicitly;
* define strict/non-strict behavior for negative and nonnumeric supplied values;
* add a large-duration formatting test.

\---

## 9.6 QPC implementation — **9.0 / 10**

### Strengths

* frequency availability captured;
* native read status checked;
* tick alignment guarded;
* Currency scaling correctly documented;
* numeric frequency helper provided;
* QPC remains isolated from ordinary second-based timestamp paths;
* no unnecessary elapsed validator call on the normal monotonic hot path.

### Remaining issue

The QPC implementation itself is strong; the problem is how its failure status is consumed by `StartTimer`.

### To remain above 9

* correct the non-strict start path;
* add failure injection;
* publish measured QPC empty-cycle distributions by Excel/Office bitness.

\---

## 9.7 Pause strategies — **8.6 / 10**

### Strengths

* hard upper cap;
* explicit no-op boundaries;
* midnight-safe Timer delta loops;
* responsive `DoEvents` alternatives;
* blocking Sleep alternative;
* coarse Wait/Now paths guarded against undershoot;
* local variables do not overwrite timing cache.

### Weaknesses

* pause method IDs are unrelated to timer IDs but use the same integer-style API;
* `DoEvents` allows reentrancy;
* invalid pause method silently routes to the Else strategy;
* no symbolic enum.

### To reach 9+

* add `cPM\_PauseMethod`;
* document reentrancy explicitly;
* validate invalid pause methods consistently;
* add cancellation-aware guidance for long UI-yielding pauses.

\---

# 10\. Measurement-harness review

## 10.1 `MeasureProcedure` — **8.5 / 10**

### Strengths

* warm-up iterations;
* isolated worker;
* raw `Double()` vector;
* cleanup and original-error preservation;
* caller checkpoint/session state remains undisturbed;
* selected backend goes through ordinary class validation.

### Deterministic procedure resolution

The harness calls:

```vb
Application.Run ProcedureName
```

An unqualified name can be ambiguous if multiple open workbooks or add-ins expose the same public macro.

### Recommended correction

Accept a fully qualified name, or qualify unqualified names against the host workbook:

```text
'WorkbookName.xlsm'!ModuleName.ProcedureName
```

An optional target workbook parameter would make the contract explicit.

### Allocation contract

`Iterations` is coerced to at least one but has no practical upper bound.

An extreme value can cause:

* array allocation failure;
* very long UI blocking;
* avoidable memory pressure.

### To reach 9+

* deterministic workbook/module qualification;
* practical maximum or named allocation error;
* optional progress/cancellation outside the measured interval;
* document no-argument Public-Sub restriction as part of the API contract;
* add duplicate-procedure-name tests across two workbooks.

\---

## 10.2 Baseline mismatch — **material methodology issue**

`MeasureProcedure` includes:

```text
StartTimer
Application.Run dispatch
target procedure
ElapsedSeconds
```

`MeasureOverhead\_Samples` measures the backend timing cycle, but does not execute an empty procedure through `Application.Run`.

Therefore:

> `MeasureOverhead\_Samples` is not a complete baseline for the overhead included by `MeasureProcedure`.

The README recommendation to subtract it for short procedures is methodologically incomplete.

### Correct baseline

Provide or document:

```text
MeasureProcedure("QualifiedEmptyPublicSub", ...)
```

using the same dispatch path.

A dedicated helper could be:

```text
MeasureDispatchOverhead\_Samples
```

but it still needs a public standard-module no-op target because `Application.Run` cannot invoke a private class method.

### Terminology correction

The minimum empty-cycle observation is not strictly “timer resolution.”

It contains at least:

* start-read cost;
* end-read cost;
* class dispatch;
* loop overhead;
* possible clock quantization.

Call it:

```text
observed empty-cycle floor
```

and report QPC tick interval separately.

### To reach 9+

* add a dispatch-matched no-op baseline;
* distinguish clock resolution from measured cycle floor;
* remove unsupported “tens of microseconds” universal wording;
* publish example measurements with environment metadata.

\---

# 11\. Statistics review

## 11.1 Overall assessment — **8.4 / 10**

The statistics API is appropriate for timing analysis, but its documentation claims a broader numerical domain than the implementation safely supports.

## 11.2 Strengths

* returns sample count;
* min/max;
* arithmetic mean;
* median;
* observed-value nearest-rank percentile;
* sample standard deviation;
* coefficient of variation;
* contamination flag;
* readable summary;
* sorting operates on a copy;
* caller order is preserved;
* one-sample and empty-array behavior are tested.

## 11.3 Mean accumulation

The mean performs ordinary summation:

```text
Total = Total + Samples(i)
Mean  = Total / N
```

For ordinary millisecond/second timing values this is acceptable.

For the documented “any `Double()` vector” domain, it can:

* overflow when the true mean is representable;
* lose low-order information across large dynamic ranges.

### Improvement

Use a compensated summation or incremental mean.

## 11.4 Even-sample median

A conventional:

```text
(a + b) / 2
```

can overflow for two large same-sign finite values even when the median is representable.

Use a safe midpoint formulation.

## 11.5 Standard deviation

The two-pass squared-deviation approach is better than the one-pass `E\[x²]-E\[x]²` formula, but it can still overflow when deviations are large.

For a broad general-vector contract, use a scaled or stable online algorithm.

## 11.6 Coefficient of variation

Current semantics are:

```text
CV = StdDev / Mean
Mean = 0 -> CV = 0
```

This is suitable only for positive-mean timing data.

For a vector such as:

```text
\[-1, +1]
```

the mean is zero and spread is nonzero, yet the function returns zero and contamination is false.

For a negative-mean vector, CV becomes negative, so a positive threshold comparison cannot flag high variation.

### Correct options

Either narrow the API contract to:

```text
finite, nonnegative timing samples with positive mean
```

or implement general semantics:

```text
CV = StdDev / Abs(Mean)
Mean near zero with nonzero spread -> undefined/error or positive infinity policy
```

## 11.7 Threshold validation

`Stats\_IsContaminated` accepts any `CvThreshold`.

A negative threshold has unintuitive behavior.

Validate:

```text
CvThreshold >= 0
```

## 11.8 To reach 9+

Choose one coherent route.

### Route A — timing-specific contract

* reject negative samples;
* document finite/moderate-magnitude domain;
* require positive mean for CV;
* validate threshold;
* keep the implementation simple.

### Route B — general numerical contract

* compensated mean;
* safe median midpoint;
* scaled/Welford-style variance;
* absolute-mean CV;
* explicit near-zero-mean policy;
* magnitude and nonfinite validation;
* extreme-value tests.

For this repository, Route A is sufficient for 9+ and more consistent with the product’s purpose.

\---

# 12\. Checkpoint and reporting review

## 12.1 Checkpoint capture — **9.1 / 10**

The 1.2 change is good:

```text
Checkpoint -> Elapsed\_ComputeSeconds -> discard raw endpoint -> append row
```

It no longer mutates `T2` and `ET`.

The storage-doubling strategy is appropriate.

### Remaining issue

In non-strict mode, an underlying failed read that returns zero can still create:

* a zero cumulative time;
* a clamped zero delta;
* a non-monotonic report.

Fixing the native failure contract resolves this downstream problem.

## 12.2 Reporting — **8.9 / 10**

### Strengths

* 2-D Variant array;
* header-only empty result;
* deterministic columns;
* text format;
* worksheet writer;
* named numeric-column constants.

### Improvements

* build large text reports through a string array plus `Join` rather than repeated concatenation;
* move worksheet rendering to a separate optional module;
* generate report-column constants from one schema or expose the schema centrally.

## 12.3 Documentation contradiction

The source and CHANGELOG are inconsistent:

* implementation and README: `Checkpoint` no longer updates cached `T2`/`ET`;
* `T2` and `ET` source headers: still say `Checkpoint` updates them;
* CHANGELOG “Changed”: says no longer overwrites;
* CHANGELOG “Documentation”: says headers now record that it updates them.

### To reach 9+

* correct both property headers;
* correct the CHANGELOG sentence;
* add a static documentation check for this invariant.

\---

# 13\. Shared TW manager review

## 13.1 Aggregate-mask design — **9.2 / 10**

The central model is strong:

```text
first session captures baseline
each instance registers a disable mask
effective mask = OR of active masks
last session restores baseline
```

## 13.2 Transaction and rollback — **9.0 / 10**

Both begin and end preserve prior dictionary state and attempt rollback/reapplication when an Excel property setter fails.

Original errors are preserved and re-raised.

This is professional error handling for VBA.

## 13.3 Calculation-baseline validity defect — **8.0 / 10**

When no workbook is open, baseline capture sets:

```vb
g\_TW\_CALCULATION = xlCalculationAutomatic
```

That does not mean Automatic was actually the baseline.

It means the baseline was unknown.

Those states must not be represented by the same value.

### Scenario A — begin without workbook, then open one

1. TW session begins with `Workbooks.Count = 0`.
2. Automatic is stored synthetically.
3. No Calculation suppression can be applied at that moment.
4. A workbook opens.
5. A later update/end can write the synthetic Automatic value.

### Scenario B — begin with workbook, close all, then end

1. Real Calculation baseline is captured.
2. Manual mode is applied.
3. All workbooks close.
4. Final end skips Calculation restoration.
5. Shared state is reset and the real baseline is forgotten.
6. A later workbook may inherit a calculation mode that was not restored.

### Scenario C — workbook opens during an unchanged active scope

No event automatically reapplies the aggregate TW state.

The Calculation suppression request may remain unapplied until another begin/update/end call occurs.

### Required design decision

Either:

#### Full transition support

* track baseline validity;
* lazily capture first valid baseline;
* retain pending restoration;
* respond to workbook lifecycle events.

or:

#### Explicit stable-host invariant

* require the open-workbook set to remain stable while Calculation suppression is active;
* strict mode raises if Calculation cannot be captured/restored;
* non-strict mode automatically exempts Calculation;
* document the limitation.

The second option is simpler and sufficient for 9+ if clearly enforced.

## 13.4 Instance key allocator

The counter-issued key is much safer than `ObjPtr`.

The remaining `Long` overflow is theoretical at more than two billion allocations per project lifetime.

This is P3 only.

## 13.5 To reach 9+

* add Calculation-baseline validity;
* define workbook transition policy;
* add four lifecycle tests;
* move report rendering out;
* optionally make the key seed rollover-safe.

\---

# 14\. Public API and error-contract review

## 14.1 Strengths

The API is practical and cohesive from the caller’s perspective.

The new functions are additive.

No intentional breaking removal was identified.

## 14.2 Magic method IDs

Timing and pause strategies still use integers.

### Improvement

Add public enums:

```vb
Public Enum cPM\_TimerMethod
    cpmTimer = 1
    cpmGetTickCount = 2
    cpmTimeGetTime = 3
    cpmTimeGetSystemTime = 4
    cpmQPC = 5
    cpmNow = 6
End Enum
```

and:

```vb
Public Enum cPM\_PauseMethod
```

The existing integer-compatible signatures can remain.

## 14.3 Error numbers are private

Errors are named internally but consumers cannot use symbolic names.

Expose:

```vb
Public Enum cPM\_ErrorCode
```

or a standard-module public error registry.

## 14.4 Non-strict ambiguity

Returning zero can mean:

```text
true zero elapsed
no active session
native read failure
negative result clamped
fallback path
```

Add:

```text
LastStatus
LastErrorCode
ResolvedMethodID
```

or `Try...` methods.

## 14.5 To reach 9+

* symbolic enums;
* public error/status contract;
* deterministic procedure qualification;
* precise statistics domain;
* backward-compatible deprecation notes for raw integer use.

\---

# 15\. Regression-test review

## 15.1 Strengths

The suite is a substantial asset.

Current source declares:

```text
52 cases
288 assertions, repository-reported
```

The new 1.2 cases cover important pure behavior.

The test header now correctly identifies:

* version 1.2.0;
* demo-builder dependency;
* Friend test hooks;
* measurement/statistics assumptions;
* non-mutating checkpoint cache semantics.

## 15.2 Important coverage gap

The README claims:

```text
strict and non-strict behaviour on every failure path
```

The suite does not have deterministic injection for:

```text
QPC start read failure
QPC elapsed read failure
timeGetSystemTime API failure
timeGetSystemTime format failure
timeBeginPeriod failure
timeEndPeriod failure
TW Application-property setter failure
workbook transition during TW scope
```

The most important production defects are therefore outside executable coverage.

## 15.3 Test/UI coupling

The runner requires:

```text
M\_DEMO\_BUILDER
Btn\_Click
status-bar helpers
worksheet presentation helpers
```

This makes it harder to:

* run headlessly;
* build a minimal test workbook;
* distinguish test infrastructure failure from production failure;
* automate through a self-hosted runner.

## 15.4 Time-based flakiness

Runtime tests should avoid tight performance thresholds.

Good assertions include:

```text
nonnegative
monotonic where guaranteed
correct state
correct method
minimum requested wait not materially undershot
```

Precise upper bounds should be characterization, not deterministic regression.

## 15.5 Registration maintenance

`TotalSteps = 52` must be manually kept in sync with runner calls.

A static source checker should compare:

```text
registered test calls
TotalSteps
documented case count
README badge
```

## 15.6 To reach 9+

1. split a headless core runner from worksheet/UI reporting;
2. inject native-reader and setter failures;
3. export machine-readable counters and failure details;
4. run on a self-hosted Excel runner;
5. periodically run Office 32-bit;
6. statically check registration and assertion metrics;
7. archive the exact result for the release commit.

\---

# 16\. CI and release-engineering review

## 16.1 Current state — **5.7 / 10**

```text
Actions workflows: none
Branch protection: none
Required checks: none
Release PR: none
v1.2.0 tag: none
v1.2.0 release: none
```

The CHANGELOG nevertheless presents:

```text
## \[1.2.0] — 2026-08-14
```

and links to a release URL that does not yet exist.

Until the tag exists, use:

```text
## \[Unreleased]
```

or:

```text
## \[1.2.0-rc]
```

## 16.2 Branch divergence

The release branch is:

```text
11 ahead
1 behind
```

It must be updated from `main` before final review and merge.

## 16.3 Recommended hosted static gate

Run without Excel:

* required-file presence;
* `Option Explicit`;
* version synchronization;
* public API snapshot;
* duplicate procedure names;
* no reintroduction of `ObjPtr` keying;
* error-constant inventory;
* README count generation;
* test registration count;
* CHANGELOG release-state validation;
* line-ending/export sanity;
* forbidden binary files.

## 16.4 Recommended Excel gate

On a dedicated Windows/Excel runner:

1. create or open a clean temporary `.xlsm`;
2. import production source;
3. import headless tests;
4. compile;
5. run;
6. collect case/assertion/failure counters;
7. extract failure details;
8. fail the workflow on nonzero failure or runtime error;
9. upload log evidence.

## 16.5 Release artifact provenance

For the release workbook publish:

```text
tag commit SHA
source-file SHA-256 hashes
workbook SHA-256
Excel version/build
Office bitness
test counters
build timestamp
```

Prefer building the workbook from source during the release workflow.

## 16.6 To reach 9+

This area cannot reach 9 through documentation alone.

It requires:

* executable workflows;
* required checks;
* protected merge;
* exact tested commit;
* source-bound artifact.

\---

# 17\. Documentation and governance review

## 17.1 Strengths

The README is visually polished and technically substantial.

It now explains:

* why sessions matter;
* backend choices;
* raw samples;
* statistics;
* TW aggregation;
* checkpoint reports;
* strict mode;
* known limitations;
* installation;
* public API.

The CHANGELOG is a meaningful improvement.

## 17.2 Exact issues to correct

### Arithmetic inconsistency

```text
41 public members
24 methods
18 properties
```

The latter two total 42.

### “No external DLL”

The code declares `kernel32` and `winmm.dll`.

The intended claim is:

```text
no third-party DLL or bundled binary dependency
```

Use that wording.

### “No dependencies”

The project has no **external** dependency, but the class requires `M\_cPM\_TIMEWASTERS.bas`.

Distinguish internal companion source from external dependency.

### “Every failure path”

Not supported without fault injection.

### “Nothing fails silently”

Not true under the current non-strict native-read behavior.

### “Exact state restoration”

Needs qualification for workbook transitions.

### “Transactional session start”

True in strict successful-read semantics, but false when non-strict mode commits a zero QPC/method-4 timestamp.

### “No magic numbers in the codebase”

The required production error numbers are centralized, but optional demo/test code still contains literals.

Use:

```text
no inline error-number literals in the required production class
```

or centralize the entire repository.

### Measurement guidance

* “minimum is your resolution floor” is imprecise;
* “Application.Run costs tens of microseconds” is environment-dependent;
* subtracting `MeasureOverhead\_Samples` does not remove dispatch overhead.

### Changelog contradiction

It says both:

```text
Checkpoint no longer overwrites T2/ET
```

and:

```text
T2 and ET headers now record that Checkpoint updates them
```

The source headers still contain the stale “updates” wording.

## 17.3 To reach 9+

* generate counts from source;
* generate test metrics from execution;
* qualify all strong claims;
* align README, source headers and CHANGELOG;
* mark unreleased versions correctly;
* add a release checklist;
* add `CONTRIBUTING.md` and `SECURITY.md`;
* label prior reviews by exact commit and mark superseded reviews clearly.

\---

# 18\. Maintainability review

## 18.1 Strengths

* consistent naming;
* comprehensive headers;
* named production errors;
* centralized helpers;
* no opaque third-party framework;
* clean source export layout;
* binary workbook removed from ordinary Git history.

## 18.2 Current pressure points

### Class size

At 202,953 bytes, the class is becoming expensive to review.

### Test size

At 298,979 bytes, the single test module is also approaching a practical navigation boundary.

### Documentation density

The README claims 76% of the class is documentation.

Detailed contracts are valuable, but repeated boilerplate can obscure executable changes.

### Manual inventories

Several counts and claims are maintained by hand.

## 18.3 To reach 9+

* split tests by domain;
* extract stateless implementations;
* generate API and test inventories;
* retain concise procedure contracts;
* move long conceptual essays to docs/wiki;
* add static checks for version and source-header freshness.

\---

# 19\. Security and platform assessment

No high-severity security vulnerability was identified in the production runtime.

## Positive characteristics

* no production network access;
* no shell invocation;
* no filesystem mutation in the runtime core;
* narrow Windows API usage;
* no third-party native binary;
* late-bound dictionary;
* `Option Private Module` on the TW support module;
* explicit timer-resolution cleanup;
* explicit global-state recovery.

## Trust boundary: `Application.Run`

`MeasureProcedure` dynamically executes a caller-supplied macro name.

That is intentional, but should be documented as:

```text
trusted procedure names only
```

Qualifying the workbook/module reduces accidental execution of the wrong macro.

## Release workbook trust

The downloadable macro-enabled workbook should be:

* built from tagged source;
* hashed;
* optionally signed;
* accompanied by compile/test evidence.

## To reach 9+

* document dynamic invocation trust;
* qualify macro resolution;
* add `SECURITY.md`;
* source-bind or sign release artifacts.

\---

# 20\. Findings summary

|ID|Severity|Area|Finding|
|-|-|-|-|
|CPM12-P1-01|**P1**|QPC correctness|Non-strict QPC start failure commits zero as a valid method-5 start tick|
|CPM12-P1-02|**P1**|Method-4 correctness|`timeGetSystemTime` failure sentinel can enter ordinary elapsed/rollover arithmetic|
|CPM12-P2-01|**P2**|TW correctness|Calculation baseline does not distinguish captured value from unknown value|
|CPM12-P2-02|**P2**|TW lifecycle|Workbook-open/close transitions can leave suppression unapplied or restoration incomplete|
|CPM12-P2-03|**P2**|Regression assurance|Critical native/setter failure paths have no deterministic injection coverage|
|CPM12-P2-04|**P2**|CI|No automated Excel compile/regression gate or branch protection|
|CPM12-P2-05|**P2**|Statistics|“Any Double vector” contract exceeds numerical and CV semantics actually supported|
|CPM12-P2-06|**P2**|Measurement methodology|Backend empty-cycle samples are not a dispatch-matched baseline for `MeasureProcedure`|
|CPM12-P2-07|**P2**|Procedure resolution|Unqualified `Application.Run` names can resolve ambiguously|
|CPM12-P2-08|**P2**|Timing contract|Single-wrap limits and Timer backward-clock ambiguity are not explicit|
|CPM12-P2-09|**P2**|Release governance|Branch diverged, no release PR, no v1.2 tag/release, no source-bound artifact|
|CPM12-P2-10|**P2**|Documentation|README/CHANGELOG/source contain contradictions and unsupported absolute claims|
|CPM12-P2-11|**P2**|Formatter contract|`ElapsedTime` narrows whole seconds to `Long`|
|CPM12-P3-01|**P3**|Architecture|Statistics, native timing, reporting and session state are concentrated in one large class|
|CPM12-P3-02|**P3**|Test architecture|Core regression execution depends on demo/UI infrastructure|
|CPM12-P3-03|**P3**|API ergonomics|Timer/pause IDs are integers and error constants are private|
|CPM12-P3-04|**P3**|Key allocator|`Long` key seed has a theoretical overflow|
|CPM12-P3-05|**P3**|Reporting|Worksheet writer remains inside the TW manager|
|CPM12-P3-06|**P3**|Performance evidence|No committed environment-tagged benchmark characterization|

\---

# 21\. Detailed P1 findings

## CPM12-P1-01 — Failed QPC start can become an uptime-like elapsed result

### Affected surface

```text
StartTimer method 5
StartTimer method 5 with AlignToNextTick
MeasureProcedure using method 5
MeasureOverhead\_Samples using method 5
Checkpoint after a failed non-strict start
```

### Root cause

Failure is represented as `0@`, then committed.

### Impact

A later successful read computes from zero.

This is a silent wrong-result path in an accepted public mode.

### Required fix

* no session commit on failed read;
* non-strict fallback to method 2;
* expose resolved method/status;
* injected regression.

### 9+ acceptance criteria

```text
\[ ] strict failure preserves previous session
\[ ] non-strict failure commits a valid fallback start
\[ ] ActiveMethodID reports fallback
\[ ] no zero sentinel stored as valid QPC start
\[ ] MeasureProcedure propagates fallback consistently
\[ ] checkpoint behavior remains monotonic
```

\---

## CPM12-P1-02 — Method-4 failure can be converted to a large positive duration

### Root cause

```text
failure -> 0 ms
0 - positive start -> negative
negative + rollover -> large positive
```

### Impact

The public result looks valid and validation cannot distinguish it.

### Required fix

* `TryRead` method;
* branch on success before arithmetic;
* no rollover on failure;
* explicit non-strict status.

### 9+ acceptance criteria

```text
\[ ] failed start never commits
\[ ] failed end never enters rollover arithmetic
\[ ] format mismatch follows the same rule
\[ ] strict/non-strict tests cover both start and end
\[ ] cached endpoint remains last successful endpoint or documented neutral value
```

\---

# 22\. Prioritized remediation plan

## Gate 1 — Native timing failure semantics

**Priority:** immediate  
**Expected score effect:** 8.2 -> approximately 8.5

Implement status-bearing readers and fallback-before-commit.

Add tests before further feature work.

## Gate 2 — Calculation-state lifecycle

**Priority:** immediate  
**Expected score effect:** approximately 8.5 -> 8.6

Choose full transition support or an enforced stable-workbook invariant.

Do not retain synthetic baseline semantics.

## Gate 3 — Headless fault-injected regression

**Priority:** before release  
**Expected score effect:** approximately 8.6 -> 8.8

Split the runner, add seams and execute in Excel.

## Gate 4 — CI, branch protection and release provenance

**Priority:** before tag  
**Expected score effect:** approximately 8.8 -> 9.05

This is the gate that converts high-quality source into a high-quality repository.

## Gate 5 — Statistics/harness contract

**Priority:** before claiming benchmark-grade generality  
**Expected score effect:** approximately 9.05 -> 9.2

Narrow or harden statistics; add a matched dispatch baseline; qualify procedure names.

## Gate 6 — Modularity and generated governance

**Priority:** post-release hardening  
**Expected score effect:** approximately 9.2 -> 9.4

Extract pure layers and generate inventories.

\---

# 23\. Concrete route from 8.2 to 9+

|Stage|Approx. score|Deliverable|
|-|-:|-|
|Current head|**8.2**|Strong RC; two P1 defects; manual assurance|
|Fix native reads + Calculation lifecycle|**8.6**|No known silent wrong-result path in accepted timing flow|
|Add fault injection + headless Excel suite|**8.8**|Critical failures are executable and repeatable|
|Add CI + protected PR + source-bound release|**9.0–9.1**|Repository becomes independently release-verifiable|
|Harden stats/harness + correct docs|**9.2**|Benchmark methodology and public claims align|
|Split pure layers + generate inventories|**9.3–9.4**|Maintainability and governance become premium|
|Dual-bitness recurring certification + signed reproducible artifact|**9.5+**|Exceptional VBA release discipline|

\---

# 24\. Release-readiness assessment

## Suitable now

The current branch is suitable for:

* continued development;
* manual review;
* controlled internal use;
* strict-mode QPC timing on ordinary supported Windows systems;
* checkpoint/reporting work;
* statistics on normal nonnegative timing vectors;
* teaching/demo use with explicit limitations.

## Not yet suitable for

* an unconditional v1.2.0 tag;
* a claim that every failure path is covered;
* a claim of exact state restoration across workbook lifecycle transitions;
* very short-procedure overhead subtraction using the current baseline helper;
* a general-purpose arbitrary-Double statistics claim;
* high-governance distribution without archived Excel execution evidence.

## Minimum release checklist

```text
CORRECTNESS
\[ ] QPC start failure fixed
\[ ] method-4 start/end failure fixed
\[ ] Calculation baseline validity fixed
\[ ] workbook-transition contract fixed or narrowed

TESTS
\[ ] native failures injected
\[ ] setter failures injected
\[ ] headless suite passes
\[ ] reported case/assertion counts generated from run
\[ ] Office 64-bit evidence archived
\[ ] Office 32-bit evidence obtained or limitation declared

BRANCH
\[ ] release branch updated from main
\[ ] release PR opened
\[ ] review findings linked to issues
\[ ] required checks green

DOCUMENTATION
\[ ] source T2/ET comments corrected
\[ ] CHANGELOG contradiction corrected
\[ ] 41/24/18 count corrected
\[ ] external-vs-third-party DLL wording corrected
\[ ] failure-coverage claim corrected
\[ ] baseline guidance corrected
\[ ] version marked Unreleased until tag

RELEASE
\[ ] exact commit tagged
\[ ] workbook built from tagged source
\[ ] SHA-256 manifest published
\[ ] Excel version/build/bitness recorded
\[ ] release asset immutable
```

\---

# 25\. What to do next, in exact order

1. **Fix `StartTimer` QPC non-strict failure.**
2. **Refactor `Get\_SystemTimeMs` into a Boolean `TryRead` contract.**
3. **Add fault injection for both paths.**
4. **Fix Calculation baseline validity and define workbook-transition behavior.**
5. **Add the four workbook lifecycle tests.**
6. **Split a headless test runner from UI logging.**
7. **Run and archive the suite manually once in Excel before adding more features.**
8. **Add self-hosted Excel CI and hosted static checks.**
9. **Update the branch from `main`; open a PR.**
10. **Fix README, CHANGELOG and source-header contradictions.**
11. **Add a dispatch-matched no-op measurement baseline.**
12. **Qualify `Application.Run` procedure names.**
13. **Narrow or harden the statistics contract.**
14. **Tag only the exact commit that passed the Excel gate.**

\---

# 26\. Final verdict

The current v1.2.0 branch represents real progress.

It has moved from a timing/session utility with checkpoints toward a small benchmarking framework:

* raw observation capture;
* robust-first reporting;
* structured sub-measurement;
* Excel environment control;
* multiple timing backends;
* explicit session contracts;
* a substantial regression suite.

The code’s strongest characteristics are:

* conceptual clarity;
* careful state ownership;
* transparent VBA implementation;
* unusually good procedure-level documentation;
* explicit cleanup;
* conservative QPC positioning;
* improved test breadth.

The remaining weaknesses are not reasons to redesign the project.

They are reasons to finish the release discipline:

* native failure status must not become data;
* workbook lifecycle must be modeled honestly;
* tests must execute automatically;
* release artifacts must be tied to source;
* numerical and documentation claims must match their real domain.

> \*\*Final score: 8.2 / 10\*\*  
> \*\*Classification: advanced v1.2.0 release candidate with strong source quality, two release-blocking correctness paths, and a material automated-assurance gap.\*\*

> \*\*9+ recommendation: fix the two P1 paths, automate fault-injected Excel tests, source-bind the release, then harden the measurement/statistics contract. No public API rewrite is required.\*\*

\---

# Appendix A — Area-by-area 9+ checklist

## Functional correctness

* \[ ] no native failure sentinel enters arithmetic
* \[ ] no failed start commits
* \[ ] strict and non-strict outcomes are explicit
* \[ ] previous session preservation tested
* \[ ] fallback method visible

## Timing robustness

* \[ ] safe duration by backend documented
* \[ ] Timer backward-clock ambiguity documented
* \[ ] single-wrap limits tested
* \[ ] failed read status separated from zero elapsed
* \[ ] checkpoint failure semantics explicit

## Architecture

* \[ ] pure statistics extracted
* \[ ] worksheet reporting extracted
* \[ ] native reader seam established
* \[ ] class facade preserved
* \[ ] dependency diagram current

## Public API

* \[ ] timer enum
* \[ ] pause enum
* \[ ] public error/status enum
* \[ ] resolved method/status
* \[ ] qualified procedure contract
* \[ ] statistics domain explicit

## TW lifecycle

* \[ ] Calculation validity flag
* \[ ] begin-without-workbook tested
* \[ ] open-during-scope tested
* \[ ] close-before-end tested
* \[ ] deferred restore or explicit invariant
* \[ ] exact restore claim matches implementation

## Regression

* \[ ] headless core suite
* \[ ] API failure injection
* \[ ] property-setter failure injection
* \[ ] machine-readable counters
* \[ ] 64-bit run artifact
* \[ ] 32-bit evidence
* \[ ] registration count generated

## CI/release

* \[ ] hosted static workflow
* \[ ] self-hosted Excel workflow
* \[ ] protected main
* \[ ] required checks
* \[ ] release PR
* \[ ] exact tested tag
* \[ ] artifact manifest

## Documentation

* \[ ] generated public count
* \[ ] generated test metrics
* \[ ] no contradictory checkpoint claims
* \[ ] no unsupported “every failure path”
* \[ ] no ambiguous “no DLL”
* \[ ] correct baseline guidance
* \[ ] Unreleased section until tag

## Maintainability

* \[ ] split test modules by domain
* \[ ] split pure stats/reporting implementation
* \[ ] static API drift check
* \[ ] version synchronization
* \[ ] prior reviews marked by commit
* \[ ] contribution/security policies

## Performance methodology

* \[ ] dispatch-matched no-op baseline
* \[ ] clock resolution separated from cycle floor
* \[ ] environment metadata
* \[ ] median/P95/CV artifacts
* \[ ] no tight CI timing thresholds
* \[ ] comparison guidance for short procedures

\---

# Appendix B — Recommended new regression cases

```text
01 QPC\_StartReadFailure\_Strict\_PreservesPreviousSession
02 QPC\_StartReadFailure\_NonStrict\_FallsBackToMethod2
03 QPC\_AlignedStartFailure\_NonStrict\_FallsBack
04 QPC\_EndReadFailure\_Strict\_Raises
05 QPC\_EndReadFailure\_NonStrict\_ReturnsStatusNotFabricatedTime
06 SystemTime\_StartFailure\_Strict\_DoesNotCommit
07 SystemTime\_StartFailure\_NonStrict\_FallsBack
08 SystemTime\_EndFailure\_Strict\_Raises
09 SystemTime\_EndFailure\_NonStrict\_SkipsRollover
10 SystemTime\_FormatFailure\_FollowsReadFailureContract
11 TW\_BeginWithoutWorkbook\_CalculationPolicy
12 TW\_OpenWorkbookDuringScope\_CalculationPolicy
13 TW\_CloseAllBeforeFinalEnd\_RestorePolicy
14 TW\_ApplySetterFailure\_BeginRollback
15 TW\_ApplySetterFailure\_EndRollback
16 MeasureProcedure\_QualifiedName\_SelectsCorrectWorkbook
17 MeasureProcedure\_DuplicateUnqualifiedName\_IsRejectedOrDefined
18 MeasureProcedure\_IterationAllocationLimit
19 Stats\_CV\_ZeroMean\_NonzeroSpread
20 Stats\_CV\_NegativeMean
21 Stats\_NegativeThreshold\_Rejected
22 Stats\_Mean\_LargeFiniteVector
23 Stats\_Median\_LargeSameSignPair
24 Stats\_StdDev\_LargeOffsetVector
25 ElapsedTime\_LargeDuration
26 Timer\_SingleMidnightWrapArithmetic
27 RollingCounter\_NearWrapArithmetic
```

\---

# Appendix C — Suggested GitHub issues

1. `Do not commit zero QPC start after non-strict read failure`
2. `Prevent timeGetSystemTime failure sentinel from entering rollover arithmetic`
3. `Model Application.Calculation baseline validity explicitly`
4. `Define workbook lifecycle invariants during TW suppression`
5. `Add injectable timing backends for failure-path tests`
6. `Create a headless cPerformanceManager regression runner`
7. `Add self-hosted Excel compile and regression workflow`
8. `Add hosted static source and documentation checks`
9. `Qualify MeasureProcedure macro resolution`
10. `Add an Application.Run-matched empty-procedure baseline`
11. `Define and enforce the Stats\_\* numeric domain`
12. `Correct README and CHANGELOG v1.2.0 contradictions`
13. `Generate public API and test-count metrics`
14. `Update release/v1.2.0 from main and open release PR`
15. `Build the release workbook from tagged source`
16. `Publish release source/artifact SHA-256 manifest`
17. `Extract M\_cPM\_STATISTICS behind compatibility wrappers`
18. `Move cPM\_Report\_WriteToRange to M\_cPM\_REPORTING`

\---

# Appendix D — Evidence confidence

|Conclusion|Confidence|
|-|-|
|Exact branch and commit|High|
|Current branch divergence|High|
|File SHAs and sizes|High|
|No workflows/protection/tag/PR|High|
|QPC zero-start defect|High from deterministic source flow|
|Method-4 failure/rollover defect|High from deterministic source flow|
|Calculation baseline validity defect|High from deterministic state model|
|Statistics domain limitations|High from arithmetic source|
|Baseline mismatch|High from measurement-path structure|
|52 registered cases|High from source metadata|
|288 assertions|Repository-reported; not independently executed|
|Current VBE compile result|Not independently executed|
|Current regression pass result|Not independently executed|
|32-bit Office behavior|Not independently executed|
|Runtime overhead values|Not independently measured|



