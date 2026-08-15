# Independent Code Review — VBA Performance Manager

> **Repository:** [`danielep71/VBA-PERFORMANCE_MANAGER`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER)  
> **Primary branch reviewed:** [`release/v1.2.0`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/tree/release/v1.2.0)  
> **1.2.0 commit reviewed:** [`c59100f4c1b8cf089dc776a4bb10d58a4040be04`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/commit/c59100f4c1b8cf089dc776a4bb10d58a4040be04)  
> **Comparison baseline:** [`main`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/tree/main) at [`c27fa650465af1ef5a4b06d9a3e684f93ba01c94`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/commit/c27fa650465af1ef5a4b06d9a3e684f93ba01c94)  
> **Published release reference:** `v1.1.0` at `2d5f31c8ea490aef6c642d91116543f24986e27a`  
> **Review date:** 2026-08-15  
> **Reviewer:** OpenAI GPT-5.6 Sol  
> **Suggested repository path:** `docs/INDEPENDENT_CODE_REVIEW_2026-08-15.md`

---

## 1. Executive assessment

### Overall repository score — `release/v1.2.0`: **8.1 / 10**

### Production runtime implementation score: **8.7 / 10**

### Regression / release-engineering score: **7.0 / 10**

### Current `main` baseline score: **7.9 / 10**

The `release/v1.2.0` branch is a **material improvement over `main`**. The core design is already well beyond a simple VBA timer wrapper: it combines six timing backends, session-bound validation, high-resolution QPC timing, Windows timer-resolution management, shared Excel Application-state suppression, structured checkpoints, machine-readable reporting, diagnostics, and a substantial regression harness.

The 1.2.0 work also addresses several genuine weaknesses in the current 1.1.x code line:

- replaces `ObjPtr(Me)`-derived shared-state keys with a dedicated instance-key allocator;
- adds validation for negative elapsed results after rollover correction;
- makes `TW_Turn_OFF` commit local state only after shared registration succeeds;
- guards `Application.Calculation` access when no workbook is open;
- adds an optional `RunLabel` argument to `StartTimer`;
- centralizes error numbers in named constants;
- removes dead/duplicate logic and improves error-source naming;
- replaces hard-coded report column numbers with constants;
- removes the versioned binary demo workbook and adds Excel binary/lock patterns to `.gitignore`.

Those are good 1.2.0 changes and, in source-quality terms, the branch is clearly stronger than `main`.

However, I would **not tag `v1.2.0` yet**. The branch currently has one high-severity correctness issue and several release-assurance gaps:

1. **Non-strict native timing-source failures can be converted into a zero timestamp and then treated as a valid timing sample.** For QPC this can produce a grossly overstated elapsed time after a failed start read; for `timeGetSystemTime`, a failed end read can be transformed by rollover correction into a very large positive elapsed time.
2. The new workbook-less `Calculation` guard uses a synthetic `xlCalculationAutomatic` baseline. If a workbook appears while the TW session is still active, the manager can later write that synthetic value back even though it never captured the real calculation mode.
3. The complete `test/` tree is unchanged between `main` and `release/v1.2.0`. The suite remains versioned as **1.1.0**, runs **41 cases**, and contains no regression for the new 1.2.0 behaviors.
4. There are **zero GitHub Actions workflows**, no automated Excel execution gate, no static source gate, and neither `main` nor `release/v1.2.0` is branch-protected.
5. The release workbook is distributed as a GitHub Release asset, but there is no source-to-artifact build/provenance pipeline binding the `.xlsm` to the exact tag/source revision.

### Independent verdict

> **The 1.2.0 branch is directionally correct and substantially better engineered than `main`, but it is not yet release-ready. Fix the non-strict timing failure semantics and the workbook-less Calculation baseline, add dedicated 1.2 regression cases, and make the test suite an automated release gate before tagging.**

With those items closed, this repository can reasonably move into the **9.0–9.3** range. Reaching a genuine **10/10** requires additional API-contract, reproducibility, performance-evidence, and maintainability work described below.

---

# 2. Review scope and methodology

## 2.1 Exact source basis

The review was performed against the exact branch heads available on 2026-08-15.

### `main`

```text
Branch: main
Commit: c27fa650465af1ef5a4b06d9a3e684f93ba01c94
Date:   2026-06-17
```

### `release/v1.2.0`

```text
Branch: release/v1.2.0
Commit: c59100f4c1b8cf089dc776a4bb10d58a4040be04
Date:   2026-08-15
Ahead of main: 5 commits
Behind main:   0 commits
```

The branch comparison reviewed was:

[`main...release/v1.2.0`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/compare/main...release/v1.2.0)

## 2.2 Production code reviewed

### Required production source

- [`src/classes/cPerformanceManager.cls`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/release/v1.2.0/src/classes/cPerformanceManager.cls)
- [`src/modules/M_cPM_TIMEWASTERS.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/release/v1.2.0/src/modules/M_cPM_TIMEWASTERS.bas)

### Regression source

- [`test/M_cPM_Test.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/release/v1.2.0/test/M_cPM_Test.bas)

### Demo / integration source

- [`demo/M_DEMO_BUILDER.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/release/v1.2.0/demo/M_DEMO_BUILDER.bas)
- [`demo/M_cPM_DEMO.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/release/v1.2.0/demo/M_cPM_DEMO.bas)
- [`demo/M_cPM_USAGE_EXAMPLES.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/release/v1.2.0/demo/M_cPM_USAGE_EXAMPLES.bas)

### Repository / release material

- `README.md`
- `.gitignore` on `release/v1.2.0`
- branch metadata and protection state
- GitHub Releases
- open issues and pull requests
- GitHub Actions workflow inventory

## 2.3 Execution boundary

Desktop Excel was not available in the review environment.

The reviewer therefore did **not**:

- import the modules into the VBE;
- execute `Debug -> Compile VBAProject`;
- run `Run_cPerformanceManager_RegressionSuite`;
- validate 32-bit Office behavior dynamically;
- inject failures into `QueryPerformanceCounter`, `timeGetSystemTime`, `timeBeginPeriod`, `timeEndPeriod`, or Excel Application-state setters;
- measure timing overhead on physical Windows hardware.

The review distinguishes between:

- **source-confirmed control-flow behavior**;
- **repository metadata / branch state**;
- **test coverage visible in source**;
- **runtime behavior that still requires desktop-Excel confirmation**.

The two principal correctness findings are source-deterministic conditional paths: if the documented failure/host transition condition occurs, the current branch control flow produces the behavior described.

---

# 3. Repository state and release topology

## 3.1 Published release versus `main`

The published `v1.1.0` tag points to:

```text
2d5f31c8ea490aef6c642d91116543f24986e27a
```

Current `main` is **28 commits ahead** of that tag.

That means:

> `main` is not the exact `v1.1.0` release baseline.

This is not inherently wrong—trunk development is normal—but the distinction should be explicit in documentation and release engineering.

## 3.2 `release/v1.2.0` branch

The branch is a clean forward branch from current `main`:

```text
ahead_by:  5
behind_by: 0
```

The five branch commits are logically grouped:

1. stop versioning the binary demo workbook;
2. ignore Excel binary/lock artifacts;
3. correct `.gitignore` patterns;
4. add `.xlam` / `.bak` patterns;
5. harden production source and bump source version to 1.2.0.

This is a sensible branch history.

## 3.3 Current governance state

At review time:

```text
GitHub Actions workflows: 0
Open issues:              0
Open pull requests:       0
main protected:           No
release/v1.2.0 protected: No
Required status checks:   None
```

For a personal VBA repository this is understandable, but it is the main reason the release-engineering score is materially below the source-code score.

---

# 4. Hard repository metrics

## 4.1 Required production source

Approximate physical source scale on `release/v1.2.0`:

| Component | Approx. physical lines | File size |
|---|---:|---:|
| `cPerformanceManager.cls` | ~4,090 | 153,181 bytes |
| `M_cPM_TIMEWASTERS.bas` | ~1,270 | 46,666 bytes |
| **Required production total** | **~5,360** | **~200 KB** |

## 4.2 Regression source

| Artifact | Current state |
|---|---:|
| Test module | `M_cPM_Test.bas` |
| Test-module version header | **1.1.0** |
| Registered runner cases | **41** |
| Approx. physical lines | **~5,890** |
| File SHA on `main` and `release/v1.2.0` | **identical** |

The unchanged test SHA is particularly important because the production source changed materially for 1.2.0.

## 4.3 Public surface

The class exposes approximately **29 user-facing members**, including:

- timing start / elapsed measurement;
- formatted elapsed output;
- strict-mode state;
- raw timing diagnostics;
- QPC diagnostics;
- pause / cleanup;
- TW suppression lifecycle;
- checkpoint lifecycle;
- structured array/text reporting.

The support module additionally exposes project-public helpers for:

- TW instance-key allocation;
- session begin/end;
- diagnostics/recovery;
- worksheet report output.

Because `M_cPM_TIMEWASTERS` uses `Option Private Module`, these procedures do not pollute the ordinary macro dialog, which is good.

## 4.4 Timing backends

The library supports six timing sources:

| ID | Backend | Intended role |
|---:|---|---|
| 1 | `Timer` | simple VBA timing |
| 2 | `GetTickCount / GetTickCount64` | monotonic uptime timing |
| 3 | `timeGetTime` | millisecond multimedia timer |
| 4 | `timeGetSystemTime` | millisecond multimedia system-time source |
| 5 | `QueryPerformanceCounter` | default / recommended benchmark source |
| 6 | `Now() * 86400` | wall-clock diagnostic timing |

---

# 5. What improved in `release/v1.2.0`

This branch contains several changes I consider genuinely valuable rather than cosmetic.

## 5.1 Shared TW instance keying

### `main`

The earlier design used an object-address-derived key.

### `release/v1.2.0`

The branch introduces:

```text
PM_TW_NewInstanceKey
```

and lazily caches the issued key on each class instance.

This is a meaningful correctness improvement because an object address is an allocation detail, not an object identity contract. The new design prevents a stale registration from being accidentally associated with a later object that happens to reuse the same address.

### Assessment

**Strong improvement.**

One theoretical overflow remains in the `Long` seed, discussed later, but the design is substantially safer than the `ObjPtr` approach.

---

## 5.2 Transactional TW local-state commit

`TW_Turn_OFF` now:

1. asks the shared manager to begin/update the session;
2. only after success commits:
   - `m_Except`;
   - `m_TW_Active`.

This is the correct transaction order.

A failed shared-state update no longer leaves the class claiming that a suppression configuration was applied when it was not.

### Assessment

**Strong improvement.**

---

## 5.3 Negative elapsed validation

The new `Elapsed_Validate` helper detects negative elapsed values after rollover correction:

```text
strict mode     -> raise
non-strict mode -> clamp to zero
```

This is particularly useful for the Win64 `GetTickCount64` path, where the rollover period is correctly zero and a negative value should not be normalized by adding a rollover constant.

### Assessment

**Good correctness hardening.**

---

## 5.4 Workbook-less host guard

The branch guards `Application.Calculation` read/write access with `Workbooks.Count`.

This correctly addresses the immediate runtime-error problem when no workbook is open.

### Assessment

**Right direction, incomplete state model.**

The implementation fixes the access exception but introduces a baseline-validity problem described in `CPM-P2-01`.

---

## 5.5 Atomic `RunLabel` at session start

`StartTimer` now accepts an optional trailing parameter:

```vb
StartTimer iMethod, AlignToNextTick, RunLabel
```

This is a good API enhancement because the class already clears checkpoint/report state at session start. A run label supplied through `StartTimer` can now survive that reset atomically.

The change is additive and preserves existing source calls.

### Assessment

**Good backward-compatible API improvement.**

---

## 5.6 Named error constants

The branch replaces many literal `vbObjectError + N` usages with named constants.

### Assessment

**Good maintainability improvement.**

A later enhancement should expose a public error-code enum so consuming projects do not need to hard-code the same numeric values.

---

## 5.7 Repository hygiene

The binary demo workbook is no longer versioned in the branch, and `.gitignore` now covers:

```text
*.xlsm
*.xlsb
~$*
*.xlam
*.bak
```

### Assessment

**Good repository-hygiene improvement.**

For a VBA source-first library, tagged source plus generated/release artifacts is preferable to repeatedly committing opaque binary workbook revisions.

---

# 6. Scoring methodology

A score of 10 requires:

- correct behavior across the documented public contract;
- no known silent wrong-result paths;
- explicit failure semantics;
- strong lifecycle cleanup;
- deterministic regression coverage for success and failure paths;
- automated Excel execution or equivalent release evidence;
- source-to-release-artifact provenance;
- synchronized version/documentation state;
- maintainable module boundaries;
- reproducible performance characterization.

## Weighted scorecard

| Area | Weight | `main` | `release/v1.2.0` | Assessment |
|---|---:|---:|---:|---|
| Functional correctness | 18% | 7.9 | **8.4** | 1.2 fixes real defects; non-strict failure path remains |
| Timing-source robustness | 15% | 7.6 | **8.2** | Better validation; backend failure semantics still unsafe |
| Architecture / modularity | 10% | 8.8 | **8.9** | Good conceptual layering; physical modules too large |
| Public API design | 8% | 8.5 | **8.6** | Stable and useful; magic method IDs and private error codes remain |
| TW lifecycle / host-state control | 10% | 7.8 | **8.3** | Better keying/transactions; Calculation baseline validity gap |
| Regression testing | 10% | 8.4 | **8.1** | Suite is strong, but 1.2 changes have no new tests |
| CI / release engineering | 10% | 5.5 | **5.7** | Binary hygiene improved; still no CI/protection/provenance |
| Documentation | 7% | 8.4 | **8.2** | Excellent source docs; branch README/test version is stale |
| Maintainability / repository hygiene | 7% | 8.3 | **8.8** | Better source hygiene; large modules remain |
| Performance engineering | 5% | 8.0 | **8.0** | Good primitives; no reproducible committed baseline |
| **Weighted overall** | **100%** | **7.9** | **8.1** |  |

---

# 7. Component scores — `release/v1.2.0`

| Component | Score | Assessment |
|---|---:|---|
| `cPerformanceManager.cls` | **8.7** | Strong timing/session architecture; native-read failure semantics need correction |
| `M_cPM_TIMEWASTERS.bas` | **8.4** | Good shared-state transaction model; workbook-less Calculation baseline needs a validity flag |
| Timing API / backends | **8.6** | Broad and carefully documented; safe-duration domain and failure states need tightening |
| TW shared-state manager | **8.5** | Good overlapping-instance design and rollback |
| Checkpoint/reporting | **9.0** | Clean session semantics, geometric storage growth, useful array export |
| Regression harness | **8.1** | Substantial 41-case suite, but stale against 1.2 and coupled to demo infrastructure |
| Documentation in source | **9.2** | Exceptionally detailed procedure-level contracts |
| Root README | **7.9** | Polished, but unchanged for 1.2 and contains version/status/typo drift |
| CI / release governance | **5.7** | Largest maturity gap |
| Repository hygiene | **8.9** | 1.2 binary-removal strategy is correct |

---

# 8. Architectural review

## 8.1 Conceptual architecture

The runtime has three logical layers:

```text
cPerformanceManager
    |
    +-- timing/session state
    +-- native timing backends
    +-- diagnostics
    +-- checkpoint/report capture
    |
    v
M_cPM_TIMEWASTERS
    |
    +-- process-wide Excel Application-state coordination
    +-- overlapping-session aggregation
    +-- baseline restore
    +-- worksheet report writer
```

The test and demo layers sit above this.

The architecture correctly recognizes a crucial Excel fact:

> Application-level performance settings are process-global state and therefore cannot safely be restored independently by each `cPerformanceManager` instance.

The shared dictionary + aggregate disable-mask design is a good solution.

## 8.2 Strong architectural choices

### Session-bound timing

`StartTimer` establishes an active timing backend, and elapsed calls validate against that session.

That prevents a common benchmarking mistake: starting with one clock and ending with another.

### Dedicated QPC path

QPC uses its own tick/frequency representation instead of forcing every backend into one generic timestamp abstraction.

That is appropriate because QPC semantics differ from second-based clocks.

### Explicit cleanup

`ResetEnvironment` is first-class and `Class_Terminate` acts only as a best-effort safety net.

That is the correct priority.

### Shared TW aggregation

Each active instance contributes a disable mask, and the effective process state is the OR of all active requests.

This gives predictable “any active suppressor wins” semantics.

### Geometric checkpoint storage growth

Checkpoint capacity doubles instead of `ReDim Preserve` on every checkpoint.

That is the correct amortized-allocation strategy in VBA.

## 8.3 Architectural concerns

### Concern A — the class is too physically large

At roughly 4,090 physical lines, the class is reviewable but expensive to reason about.

The problem is not the public facade; the public class is a good abstraction. The issue is that the class currently owns:

- API declarations;
- six timer implementations;
- pause logic;
- session validation;
- QPC helpers;
- rollover helpers;
- timer-resolution lifecycle;
- diagnostics;
- TW facade;
- checkpoint capture;
- report serialization.

A 10/10 architecture would keep the `cPerformanceManager` facade but extract implementation detail into project-private modules, for example:

```text
M_cPM_NATIVE_TIMING
M_cPM_TIMING_HELPERS
M_cPM_TW
M_cPM_REPORTING
```

This can be done without breaking the public class API.

### Concern B — TW module contains unrelated report output

`M_cPM_TIMEWASTERS` now also owns:

```text
cPM_Report_WriteToRange
```

That is a single-responsibility leak.

The worksheet writer should live in an optional reporting module, not in the required process-wide state manager.

### Architectural verdict

**8.9 / 10**

The conceptual architecture is strong. The next step is physical decomposition, not redesign.

---

# 9. Production-code review — `cPerformanceManager`

## 9.1 Class initialization and lifecycle

### Strengths

- QPC frequency availability is established once during initialization.
- method labels are initialized centrally.
- session and checkpoint state are deterministic.
- `Checkpoint_ClearState` is now reused rather than duplicating initialization.
- `Class_Terminate` does not attempt complicated recovery logic; it delegates to explicit cleanup and suppresses termination-time errors.

### Residual limitation

If `ResetEnvironment` fails during `Class_Terminate`, the error is deliberately suppressed. That is unavoidable for a destructor-style fallback, but it means a stale TW registration may remain in shared state.

The new collision-proof key prevents that stale entry from being inherited by a future object, which is good, but the stale registration can still keep an aggregate suppression request alive until explicit global recovery.

The documented `PM_TW_EndAllSessions` recovery path is therefore important.

---

## 9.2 `StartTimer`

### Strengths

`StartTimer` is now much closer to transactional behavior:

1. validate/normalize method;
2. capture a local start timestamp;
3. only then commit the active session;
4. reset checkpoint state;
5. apply optional `RunLabel`.

That order is good.

### Critical residual issue

In non-strict mode, a failed QPC read is converted to:

```vb
NewQpcStart = 0@
```

and the session is still committed as method 5.

If a later QPC read succeeds, the class can calculate:

```text
(current QPC tick - 0) / frequency
```

which is not the elapsed time of the requested operation.

This is a **failure sentinel being promoted into valid state**.

See `CPM-P1-01`.

---

## 9.3 `ElapsedSeconds`

### Strengths

- session existence is validated;
- explicit method mismatches are rejected/coerced according to strict mode;
- 32-bit rollover is handled centrally;
- the Win64 method-2 no-rollover case is explicit;
- negative elapsed values after rollover correction are now validated;
- QPC return status is checked;
- cached `T2` / `ET` diagnostics are maintained.

### Important non-strict issue — method 4

`Get_SystemTimeMs` returns zero in non-strict mode when:

- `timeGetSystemTime` fails; or
- it returns an unexpected format.

At an elapsed read, that zero is then used as an actual timestamp.

If the start timestamp was positive:

```text
ET = 0 - start
```

The method-4 rollover branch sees a negative value and adds the 32-bit rollover period:

```text
ET = rollover - start
```

The result is now **large and positive**, so `Elapsed_Validate` accepts it.

That is worse than returning an explicit zero/failure sentinel because it looks like a valid measurement.

### Correct design

Native timestamp acquisition should return:

```text
success flag + timestamp
```

not overload numeric zero to mean both:

- a valid timestamp;
- a failed read.

---

## 9.4 Safe-duration domain

The code corrects at most one rollover for periodic counters.

Therefore the mathematically safe single-session domains are approximately:

```text
Timer:                    < 24 hours
GetTickCount (Win32):     < 2^32 ms  ~= 49.71 days
timeGetTime:              < 2^32 ms  ~= 49.71 days
timeGetSystemTime:        < 2^32 ms  ~= 49.71 days
GetTickCount64 (Win64):   effectively unbounded for practical VBA sessions
QPC:                      practical long-duration source
```

A session that spans more than one complete wrap can produce a small positive remainder and therefore evade negative-elapsed validation entirely.

This is not a problem for ordinary benchmarks, but it is part of the public numerical/timing contract and should be documented explicitly.

---

## 9.5 QPC implementation

### Strengths

- QPC read success is checked.
- QPC frequency availability is checked.
- Currency storage is explained in detail.
- frequency conversion is centralized.
- next-tick alignment has a finite spin guard.
- method 5 avoids unnecessary generic validation overhead on its hot path.

### Improvement

The current non-strict QPC start failure should not commit method 5 with a zero start.

Preferred behavior:

```text
strict mode:
    raise QPC-read error

non-strict mode:
    fall back to method 2 BEFORE committing the timing session
```

At an end-read failure, switching clocks is impossible because the session already started with QPC. In non-strict mode the safest fallback is therefore an explicit failure result such as zero plus a status surface, rather than fabricating elapsed time.

---

## 9.6 `ElapsedTime`

The formatter supports an already measured elapsed-seconds value, which is useful and avoids a second measurement.

However, the conversion path uses a VBA `Long` for whole seconds:

```vb
WholeS = CLng(Fix(Sec))
```

That creates an undocumented upper bound near:

```text
2,147,483,647 seconds
~68.1 years
```

The method accepts a `Variant`, but it does not document:

- accepted numeric types;
- maximum representable input;
- behavior for nonnumeric values;
- overflow behavior.

This is not a practical benchmark-duration problem, but it is a public input-contract gap.

A 10/10 implementation would keep whole seconds in `Double` or `Currency`, derive day/hour/minute components without narrowing to `Long`, and validate the supplied Variant explicitly.

---

## 9.7 Pause methods

### Strengths

- explicit upper duration cap;
- no-op behavior for nonpositive requests;
- midnight-safe Timer delta logic;
- `DoEvents` only on yielding methods;
- method 3/4 guard loops prevent coarse time sources from undershooting the requested duration.

### Improvement

The pause API uses integer method IDs unrelated to the timing IDs. That is easy to misuse.

Add a public enum for pause mode while preserving the numeric API.

---

## 9.8 Timer-resolution management

The `timeBeginPeriod(1)` / `timeEndPeriod(1)` lifecycle is handled carefully:

- request only when needed;
- track successful ownership;
- balance in cleanup;
- do TW cleanup before re-raising a timer-resolution release failure.

This is good.

The source documentation also correctly warns that the request affects broader system timer behavior and power use.

---

## 9.9 Diagnostics and overhead

The class provides useful diagnostic properties and isolates overhead measurement in a temporary worker instance.

The documentation explicitly acknowledges that arithmetic mean is sensitive to scheduler outliers.

That honesty is good.

For benchmark-grade characterization, add:

- median;
- minimum;
- p90/p95/p99;
- warm-up count;
- observation count;
- environment metadata.

Keep those as diagnostics, not hard pass/fail thresholds.

---

# 10. Production-code review — `M_cPM_TIMEWASTERS`

## 10.1 Shared session model

### Strengths

- lazy dictionary creation;
- no early-bound Scripting Runtime reference;
- per-instance registration;
- aggregate mask recomputation;
- first-session baseline capture;
- final-session restore;
- rollback when effective-state application fails;
- preservation/re-raise of the original error;
- `Option Private Module`.

This is strong VBA infrastructure.

---

## 10.2 New instance-key allocator

The 1.2 design is materially better than `ObjPtr`.

### Residual theoretical issue

The key seed is:

```vb
Private g_TW_KeySeed As Long
```

and `PM_TW_NewInstanceKey` increments it without overflow handling.

After `2,147,483,647` allocations in one VBA project lifetime, the increment can overflow despite the documented “does not raise errors” contract.

This is not practically reachable in ordinary usage and is therefore P3, but it is a contract inconsistency.

Use a larger monotonic representation or explicit wrap/collision checking if you want the allocator to be formally total.

---

## 10.3 Begin-session transaction

`PM_TW_BeginSession` is thoughtfully implemented:

- captures previous registration state;
- captures baseline only for the first session;
- writes the new registration;
- applies the aggregate state;
- on failure rolls the registration back;
- best-effort reapplies the previous effective state;
- preserves the original error.

This is one of the strongest parts of the repository.

---

## 10.4 End-session transaction

The end path similarly:

- remembers prior registration;
- removes the target key;
- restores baseline if final;
- otherwise recomputes aggregate state;
- rolls removal back on failure.

Good.

One documentation nuance: an unknown key is described as idempotent/no-op, but when a store exists the routine can still recompute/reapply global state. It therefore is not a completely side-effect-free no-op.

That is minor.

---

## 10.5 Workbook-less Calculation baseline — important defect

The 1.2 branch correctly guards `Application.Calculation` when no workbook is open, but the fallback baseline is:

```vb
g_TW_CALCULATION = xlCalculationAutomatic
```

even though the actual Calculation state was not read.

Later, `PM_TW_ApplyEffectiveState` does this whenever a workbook exists:

```text
disabled -> set Manual
not disabled -> set cached g_TW_CALCULATION
```

Consider this valid lifecycle:

```text
1. First TW session begins while Workbooks.Count = 0.
2. Calculation cannot be read, so the manager caches Automatic as a synthetic value.
3. A workbook is opened while the session remains active.
4. A later update/end applies Calculation state.
5. The manager may write Automatic even though Automatic was never the captured baseline.
```

The branch therefore fixes the original “cannot access Calculation” problem but does not distinguish:

```text
baseline value = Automatic
```

from:

```text
baseline value = unknown / not captured
```

### Correct design

Add:

```vb
g_TW_CalculationBaselineValid As Boolean
```

Rules:

```text
first session with workbook:
    capture Calculation
    valid = True

first session without workbook:
    valid = False
    do not invent a baseline

later apply with workbook and Calculation suppression requested:
    if not valid:
        lazily capture current Calculation
        valid = True
    set Manual

later restore:
    if valid:
        restore captured value
    else:
        do not touch Calculation
```

This preserves the exact baseline contract.

---

## 10.6 Report writer location

`cPM_Report_WriteToRange` is useful, but it does not belong in the TW manager.

Move it to:

```text
M_cPM_REPORTING.bas
```

This will:

- make the required TW module smaller;
- keep global-state code focused;
- make reporting optional;
- reduce coupling between report column structure and the state manager.

---

# 11. Public API and error-contract review

## 11.1 Strengths

The user-facing API is pragmatic and consistent enough for VBA:

```text
StartTimer
ElapsedSeconds
ElapsedTime
Pause
ResetEnvironment

TW_Turn_OFF
TW_Turn_ON

SetRunLabel
Checkpoint
ClearCheckpoints
ReportAsArray
ReportAsText
```

The structured report API is especially useful because it separates:

- capture;
- machine-readable export;
- text presentation;
- worksheet rendering.

## 11.2 Additive 1.2 compatibility

The new trailing optional `RunLabel` argument on `StartTimer` is source-compatible with existing calls.

No intentional public member removal was identified in the 1.2 branch.

## 11.3 Magic method IDs

The timer and pause surfaces still rely heavily on numeric IDs.

Add:

```vb
Public Enum cPM_TimerMethod
    cpmTimer = 1
    cpmGetTickCount = 2
    cpmTimeGetTime = 3
    cpmTimeGetSystemTime = 4
    cpmQPC = 5
    cpmNow = 6
End Enum
```

and an equivalent pause enum.

Keep the current numeric signature for compatibility if desired.

## 11.4 Error codes are named but private

1.2 centralizes error numbers, which is good internally.

But consuming modules cannot write:

```vb
If Err.Number = ERR_CPM_NO_ACTIVE_SESSION Then ...
```

because the constants are private.

Expose a public enum or a documented error-code module.

That would make strict-mode errors a real API contract rather than an implementation detail.

## 11.5 Strict versus non-strict semantics need a table

Current non-strict behavior varies by condition:

- invalid requested method -> fallback;
- QPC unavailable -> fallback;
- no session -> zero;
- method mismatch -> active method;
- negative elapsed -> zero;
- QPC read failure at elapsed -> zero;
- QPC read failure at start -> zero timestamp committed;
- `timeGetSystemTime` failure -> zero timestamp consumed downstream.

This inconsistency is the root cause of the P1 finding.

Document a single rule:

> **Non-strict mode may coerce inputs or select an alternate backend only before a session is committed. It must never convert a backend read failure into a timestamp that can be mistaken for valid data.**

---

# 12. Regression-test review

## 12.1 Existing strengths

The regression harness is substantial for a VBA project.

It has **41 explicitly registered cases** spanning:

### Core timing

- constructor/default state;
- method-name mapping;
- start-session state;
- elapsed seconds;
- formatted elapsed time;
- aligned starts;
- QPC accessors;
- method-4 behavior.

### Validation

- elapsed before start;
- method mismatch;
- invalid start method;
- non-strict fallback.

### Diagnostics

- numeric overhead;
- formatted overhead;
- system/QPC diagnostics.

### Pause

- methods 1–4;
- boundary/no-op behavior.

### TW lifecycle

- blank keys;
- single instance;
- overlapping instances;
- reset/idempotence;
- inactive `TW_Turn_ON`.

### Checkpoint/reporting

- no-session errors;
- run labels;
- default names;
- count;
- array/text reports;
- clearing/reuse;
- delta/cumulative semantics.

That is a strong base.

## 12.2 Release-blocking assurance gap

The test file on `release/v1.2.0` is byte-for-byte the same as the current `main` test file.

Its own header still says:

```text
VERSION 1.1.0
UPDATED 2026-04-18
TotalSteps = 41
```

Yet 1.2 changes:

- shared instance identity;
- TW transaction order;
- elapsed validation;
- workbook-less host behavior;
- `StartTimer` signature;
- error constants/sources;
- internal cleanup helpers.

No new tests are registered.

### Required 1.2 regression cases

At minimum add:

```text
StartTimer_RunLabel_AssignedAtomically
StartTimer_RunLabel_ClearedOnNextSession
TW_InstanceKeys_AreUniqueAcrossObjectLifetimes
TW_FailedBegin_DoesNotCommitLocalState
Elapsed_Negative_StrictRaises
Elapsed_Negative_NonStrictReturnsZero
QPC_StartReadFailure_NonStrict_DoesNotCommitZeroTick
TimeGetSystemTime_StartFailure_NonStrict_NoBogusSession
TimeGetSystemTime_EndFailure_NonStrict_NoRolloverFabrication
TW_NoWorkbook_StartThenWorkbookOpen_DoesNotInventCalculationBaseline
TW_NoWorkbook_LazyCalculationCapture_WhenFirstTouched
```

Failure injection is required for several of these.

## 12.3 Test harness is not self-contained

The test module says its dependencies are essentially the production class/module and Excel, but the runner directly invokes demo infrastructure such as:

```text
DEMO_Begin_FastMode
DEMO_End_FastMode
DEMO_Build_DemoTemplate
Demo_SB_SetProgress
Btn_Click
```

Therefore the regression suite is coupled to optional demo code.

That is undesirable for CI and release assurance.

### Recommendation

Create a **headless/core regression runner** that imports only:

```text
cPerformanceManager.cls
M_cPM_TIMEWASTERS.bas
M_cPM_Test.bas
```

Keep pretty worksheet logging as an optional wrapper.

## 12.4 Failure injection

Several critical branches cannot be tested reliably unless native/API failure can be forced.

Add test seams around:

```text
QueryPerformanceCounter
timeGetSystemTime
timeBeginPeriod
timeEndPeriod
PM_TW_ApplyEffectiveState
```

This can be done without production overhead through:

- thin private wrappers;
- conditional compilation;
- test-only module flags.

---

# 13. CI and release-engineering review

## 13.1 Current state

The GitHub repository currently exposes:

```text
Actions workflows: 0
```

Both relevant branches are unprotected and have no required status checks.

This is the largest maturity gap relative to the quality of the source itself.

## 13.2 Recommended two-layer CI

### Layer 1 — hosted static/source gate

Can run on ordinary GitHub-hosted runners:

- required files exist;
- `Option Explicit` present;
- version headers synchronized;
- public API inventory snapshot;
- no duplicate public procedures;
- no reintroduction of `ObjPtr` instance keying;
- no literal `vbObjectError + N` in class body;
- test `TotalSteps` matches registered cases;
- test version matches release version;
- README release examples are current;
- `.gitignore` policy validated;
- line-ending/export sanity checks.

### Layer 2 — self-hosted Windows + Excel gate

Use a dedicated Windows/Excel runner to:

1. create/open a temporary macro-enabled workbook;
2. import production source;
3. import the headless regression module;
4. compile;
5. execute the suite;
6. return machine-readable counters;
7. fail the workflow on any assertion/runtime error;
8. upload the log workbook/text artifact.

This is exactly the kind of project where a self-hosted Excel runner adds significant credibility.

## 13.3 Branch protection

Protect `main` and require:

```text
static-source-check
excel-vba-regression
```

A release branch can remain writable, but the merge/tag should require green checks.

## 13.4 No current PR for 1.2

There is no open PR for:

```text
release/v1.2.0 -> main
```

Open one before release and use the review checklist as its acceptance criteria.

---

# 14. Release artifact and provenance review

## 14.1 Binary removal from Git is good

Moving the demo workbook out of ordinary source history is correct.

Binary `.xlsm` files:

- do not diff meaningfully;
- inflate Git history;
- are difficult to review.

The 1.2 approach is better.

## 14.2 But the release asset is not source-bound

The current `v1.1.0` tag is dated April 2026, while its downloadable `.xlsm` asset was created/uploaded in August 2026.

The asset has a SHA-256 digest in GitHub metadata, which is useful, but there is no repository workflow showing:

```text
tag source
  -> deterministic workbook assembly/import
  -> compile/test
  -> SHA-256 manifest
  -> release upload
```

Therefore the workbook is not independently reproducible from the tag.

### Recommendation

For `v1.2.0`, publish a manifest:

```text
tag/commit SHA
cPerformanceManager.cls SHA-256
M_cPM_TIMEWASTERS.bas SHA-256
demo source SHA-256
test source SHA-256
PERFORMANCE.MANAGER.xlsm SHA-256
Excel version/build used to assemble/test
Office bitness
test result summary
```

Best case: automate workbook assembly from a template on the self-hosted Excel runner.

Optional higher-assurance step: digitally sign the VBA project/release workbook.

---

# 15. Documentation review

## 15.1 Source documentation

The production source is unusually well documented for VBA.

Procedure headers consistently cover:

- purpose;
- rationale;
- inputs;
- behavior;
- error policy;
- dependencies;
- implementation notes.

That is a major strength.

## 15.2 README is stale relative to 1.2

The `release/v1.2.0` README is unchanged from `main`.

It therefore does not explain:

- the 1.2 release;
- new `StartTimer(..., RunLabel)` usage;
- new elapsed validation semantics;
- new shared-key model;
- workbook-less host behavior;
- binary-distribution policy.

The checkpoint example still uses:

```text
StartTimer
SetRunLabel
```

rather than showing the new atomic `StartTimer` label option.

## 15.3 Version drift

Current state:

```text
production source: 1.2.0
test module:       1.1.0
README:            no 1.2 release documentation
demo source:       unchanged from main
```

That should be resolved before tagging.

## 15.4 Minor README polish

Visible issues include:

```text
Role-Eexecution Engine
Host-Wndows Excel VBA
Status-FINAL
```

`Status-FINAL` is also awkward on an actively developed branch.

Use a version/release-status badge instead.

---

# 16. Performance-engineering review

## 16.1 Strengths

- QPC is correctly treated as the preferred benchmark path.
- aligned-start logic is explicitly optional.
- repeated checkpoint allocation is amortized.
- overhead measurement runs on an isolated instance.
- source comments acknowledge scheduler contamination of arithmetic means.

## 16.2 Missing evidence

There is no committed reproducible performance baseline for the library itself.

A performance manager should ideally publish non-gating characterization for:

```text
StartTimer + ElapsedSeconds QPC overhead
Checkpoint overhead
ReportAsArray scaling
TW_Turn_OFF / TW_Turn_ON overhead
aligned-start overhead
method 1..6 observed granularity
```

Record:

```text
Excel version
Office bitness
Windows version
CPU
power plan
iterations
warmups
median
p95
minimum
```

Do not make tight timing thresholds release-blocking because shared-runner timing is noisy.

---

# 17. Security and platform assessment

No high-severity security defect was identified in the production source.

Positive characteristics:

- no production network access;
- no shell execution;
- no external command invocation;
- no filesystem mutation in the runtime core;
- Windows API use is narrowly scoped to timing/sleep/clock functions;
- late-bound `Scripting.Dictionary` avoids a manual reference;
- internal support module is `Option Private Module`;
- the class is not exposed as a COM-creatable public class.

The principal operational risk is not malicious behavior but **global Excel Application state**. The repository treats that risk seriously through baseline capture and coordinated restore.

For public `.xlsm` distribution, stronger release provenance and optional VBA code signing would further improve trust.

---

# 18. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| CPM-P1-01 | **P1** | Timing correctness | Non-strict native read failures can be converted into zero timestamps and later produce fabricated elapsed values |
| CPM-P2-01 | **P2** | TW correctness | Workbook-less Calculation baseline uses synthetic `Automatic` and can later overwrite an uncaptured real baseline |
| CPM-P2-02 | **P2** | Regression assurance | 1.2 production source changed materially but the complete test tree is unchanged and still versioned 1.1.0 |
| CPM-P2-03 | **P2** | CI / governance | Zero Actions workflows; no automated compile/test gate; branches unprotected |
| CPM-P2-04 | **P2** | Timing contract | Maximum safe single-session duration for rolling clocks is not explicit/enforced |
| CPM-P2-05 | **P2** | Release provenance | `.xlsm` release artifact is not reproducibly/source-bound to the tag |
| CPM-P2-06 | **P2** | API/error contract | `ElapsedTime` narrows whole seconds to `Long` and has an undocumented large-input overflow boundary |
| CPM-P2-07 | **P2** | Test architecture | Regression harness depends on optional demo infrastructure and is not headless/self-contained |
| CPM-P2-08 | **P2** | Documentation | Source is 1.2.0 while tests/README/examples remain on the 1.1-era surface |
| CPM-P3-01 | **P3** | Key allocator | `Long` instance-key seed has a theoretical overflow despite “does not raise” contract |
| CPM-P3-02 | **P3** | Modularity | 4k-line class and TW/reporting mixed responsibilities increase review cost |
| CPM-P3-03 | **P3** | API ergonomics | Timing/pause IDs are magic integers; error constants are private |
| CPM-P3-04 | **P3** | Performance evidence | Arithmetic-mean overhead helper is useful but no reproducible benchmark distribution is committed |
| CPM-P3-05 | **P3** | Documentation polish | README badges/status contain typos and active-development drift |

---

# 19. Detailed findings

## CPM-P1-01 — Non-strict native read failures can fabricate elapsed time

### Severity

**P1 — release blocker**

### Affected paths

```text
StartTimer, method 5 / QPC
QPC aligned-start initial read
StartTimer, method 4 / timeGetSystemTime
ElapsedSeconds, method 4 / timeGetSystemTime
```

### Root cause

A failed native timestamp read is represented as numeric zero.

The caller then cannot distinguish:

```text
timestamp = 0 because counter value is actually zero
```

from:

```text
timestamp = 0 because the read failed
```

### QPC start example

In non-strict mode:

```text
QPC read fails
-> NewQpcStart = 0
-> method 5 session is committed
-> later QPC read succeeds
-> elapsed = current_tick / frequency
```

The result can be orders of magnitude larger than the true operation duration.

### Method-4 end example

```text
start timestamp = positive
end read fails -> 0
raw delta = negative
rollover correction adds 2^32 ms
result becomes large positive
Elapsed_Validate accepts it
```

### Remediation

Introduce status-bearing readers:

```vb
Private Function QPC_TryReadTick(ByRef TickOut As Currency) As Boolean
Private Function SystemTime_TryReadMs(ByRef MsOut As Double) As Boolean
```

Do not commit a session until the start read succeeds.

Recommended non-strict start policy:

```text
requested QPC fails -> fallback to method 2, capture method-2 start, commit method 2
requested method 4 fails -> fallback to method 2, capture method-2 start, commit method 2
```

Recommended end-read policy:

```text
strict -> raise
non-strict -> return explicit failure sentinel/status; do not run rollover arithmetic on failed data
```

Add deterministic failure-injection tests.

---

## CPM-P2-01 — Workbook-less Calculation baseline validity is lost

### Severity

**P2 — fix before release**

### Root cause

No-workbook baseline capture sets:

```text
g_TW_CALCULATION = xlCalculationAutomatic
```

instead of representing “not captured.”

### Consequence

A later workbook-open transition can cause the manager to restore `Automatic` even though that value was never captured.

### Remediation

Add a validity flag and lazy first-touch capture.

---

## CPM-P2-02 — 1.2 source has no 1.2 regression delta

### Severity

**P2 — release assurance**

### Evidence

The 1.2 branch changes both required production source files.

The test file SHA is unchanged from `main` and the test header still reports:

```text
VERSION 1.1.0
TotalSteps = 41
```

Searches of the test source show no dedicated coverage for:

```text
PM_TW_NewInstanceKey
ERR_CPM_NEGATIVE_ELAPSED
StartTimer third RunLabel argument
workbook-less host transition
```

### Remediation

Add the 1.2 regression inventory before tagging.

---

## CPM-P2-03 — No automated release gate

### Severity

**P2**

### Evidence

```text
Actions workflows = 0
main protected = false
release/v1.2.0 protected = false
```

### Remediation

Add hosted static checks + self-hosted Excel execution and require them on merge/tag.

---

## CPM-P2-04 — Rolling-clock safe duration is implicit

### Severity

**P2 — contract/documentation**

### Risk

Single-roll correction cannot recover multiple wraps.

### Remediation

Document supported maximum elapsed interval per backend and recommend QPC / GetTickCount64 for long sessions.

Optional: reject clearly unsupported long-duration use where detectable, though exact multi-wrap detection is impossible without intermediate samples.

---

## CPM-P2-05 — Release workbook is not reproducibly source-bound

### Severity

**P2 — release provenance**

### Evidence

The published `v1.1.0` tag and the current release asset were created months apart, and the repository has no release build workflow.

### Remediation

Build/upload the `.xlsm` from tagged source in automation and publish a manifest of source/artifact hashes and Excel environment.

---

## CPM-P2-06 — `ElapsedTime` large-input overflow boundary

### Severity

**P2 — public input contract**

### Root cause

```vb
WholeS = CLng(Fix(Sec))
```

### Remediation

Avoid narrowing elapsed seconds to `Long`; validate `ElapsedSecondsIn` explicitly.

---

## CPM-P2-07 — Regression suite is coupled to demo code

### Severity

**P2 — test architecture**

### Evidence

The runner calls demo helpers for sheet creation, fast mode, progress/status, and button behavior.

### Remediation

Split:

```text
M_cPM_CoreTests.bas      -> no demo dependency
M_cPM_TestUI.bas         -> optional worksheet visualization
```

The CI runner should use only the core suite.

---

## CPM-P2-08 — 1.2 documentation/version drift

### Severity

**P2 — release hygiene**

### Remediation

Before tag:

- bump test module to 1.2.0;
- add 1.2 release notes / CHANGELOG;
- document `StartTimer(..., RunLabel)`;
- document new error/fallback behavior;
- explain binary distribution policy;
- update examples and badges.

---

# 20. Prioritized remediation plan

## Release Gate 1 — Fix native timestamp failure semantics

**Must complete before tag.**

1. Never use zero as both timestamp and failure status.
2. Add status-bearing read helpers.
3. Make non-strict start failure fall back before session commit.
4. Make non-strict end failure return an explicit safe failure result.
5. Add injected QPC and method-4 failure tests.

---

## Release Gate 2 — Correct workbook-less Calculation state

**Must complete before tag if workbook-less support is claimed.**

1. add `g_TW_CalculationBaselineValid`;
2. remove synthetic baseline semantics;
3. lazily capture Calculation when first touch becomes possible;
4. restore only when a real baseline exists;
5. test the host transition.

---

## Release Gate 3 — Bring tests to 1.2

1. update test header/version;
2. add new test cases;
3. update `TotalSteps`;
4. verify all new errors and rollback paths;
5. add failure injection.

Target:

```text
41 existing cases
+ 10–15 focused 1.2 cases
= ~51–56 cases
```

---

## Release Gate 4 — Add CI

### Hosted

- static source checks;
- version synchronization;
- API snapshot;
- test registration consistency.

### Self-hosted Windows/Excel

- import;
- compile;
- run;
- return machine-readable pass/fail;
- upload evidence.

---

## Release Gate 5 — Source-bound release artifact

1. tag exact commit;
2. build/import workbook on release runner;
3. compile/run regression;
4. compute SHA-256;
5. publish manifest;
6. upload asset.

---

## Release Gate 6 — Documentation

- README 1.2 changes;
- CHANGELOG;
- supported timing-duration table;
- strict/non-strict behavior matrix;
- `StartTimer` run-label example;
- release artifact provenance instructions.

---

# 21. What is required to reach 10/10

Closing the release blockers gets the repository into the low 9s.

To reach a defensible 10:

## 21.1 Public symbolic enums

Add:

```text
cPM_TimerMethod
cPM_PauseMethod
cPM_ErrorCode
```

without breaking the existing surface.

## 21.2 Split internal implementation

Keep the class facade stable while extracting:

```text
native timing
TW global state
worksheet reporting
```

into focused modules.

## 21.3 Deterministic fault injection

Every native/error branch should be testable without waiting for an actual Windows API failure.

## 21.4 32-bit and 64-bit evidence

Publish a small certification matrix:

```text
Office 32-bit
Office 64-bit
methods 1..6
TW lifecycle
checkpoint/reporting
```

## 21.5 Reproducible benchmark characterization

Publish environment-tagged median/p95 overhead rather than a single mean.

## 21.6 Release immutability

Treat tagged release assets as immutable outputs of the tagged source.

## 21.7 Static API drift gate

Have CI extract the public member list and fail on accidental breaking changes unless explicitly approved.

---

# 22. Release-readiness assessment

## Suitable now for development / internal use

Yes.

The branch is already suitable for:

- continued development;
- manual Excel testing;
- ordinary short-duration QPC timing;
- checkpoint/reporting use;
- TW suppression in conventional workbook-hosted scenarios.

## Ready to merge to `main`

**After the P1 native-read failure fix and new 1.2 regression tests.**

## Ready to tag `v1.2.0`

**Not yet.**

Minimum tag criteria:

```text
[ ] CPM-P1-01 fixed
[ ] CPM-P2-01 fixed or workbook-less support explicitly narrowed
[ ] 1.2 regression cases added
[ ] full suite executed successfully in Excel
[ ] README / version metadata synchronized
[ ] release asset built from the tag and hash recorded
```

Strongly recommended:

```text
[ ] CI workflow active
[ ] release PR reviewed
[ ] main protected with required checks
```

---

# 23. Main versus 1.2 summary

| Area | `main` | `release/v1.2.0` |
|---|---|---|
| TW instance identity | object-address based | **counter-issued key — better** |
| TW local commit | less transactional | **commit-after-success — better** |
| negative elapsed | weaker validation | **explicit validation — better** |
| no-workbook Calculation | can fail on access | **access guarded, but baseline-validity gap remains** |
| StartTimer run label | separate call only | **optional atomic argument** |
| error numbers | literals dispersed | **named constants** |
| report columns | hard-coded | **named constants** |
| binary workbook in Git | present | **removed** |
| `.gitignore` | absent | **Excel artifacts ignored** |
| regression tests | 41 cases | **same 41 cases — no 1.2 additions** |
| CI | none | **none** |
| branch protection | none | **none** |
| README | 1.1-era | **unchanged** |

### Net assessment

> **1.2 is clearly the better code line. The problem is not the direction of the branch; it is that the assurance layer has not moved with the production layer.**

---

# 24. Final verdict

`VBA-PERFORMANCE_MANAGER` is an unusually mature VBA utility in terms of source-level reasoning and documentation.

Its best qualities are:

- session-bound timing semantics;
- multiple carefully separated timing backends;
- QPC as the preferred benchmark source;
- explicit rollover handling;
- timer-resolution ownership/cleanup;
- process-wide TW coordination rather than unsafe per-object restoration;
- overlapping-session aggregation;
- transactional shared-state rollback;
- structured checkpoint/report output;
- extensive in-source contracts;
- a serious regression harness;
- improved 1.2 repository hygiene.

The 1.2.0 branch fixes several real design weaknesses in `main`, and I recommend continuing from it rather than trying to patch `main` independently.

The principal technical issue is now narrow and identifiable:

> **A failed native timing read must never become an ordinary numeric timestamp.**

The principal process issue is equally clear:

> **The production code has moved to 1.2.0, while the regression/release machinery is still effectively 1.1-era and manual.**

### Final scores

```text
Current main:                   7.9 / 10
release/v1.2.0 repository:      8.1 / 10
release/v1.2.0 source quality:  8.7 / 10
release/test engineering:       7.0 / 10
```

### Recommendation

> **Keep the `release/v1.2.0` branch, fix the two state/failure semantics described above, add a focused 1.2 test delta, automate the Excel suite, and then release. Do not redesign the public class. The next gains come from hardening and assurance, not from adding more features.**

---

# Appendix A — Recommended 1.2 regression inventory

```text
01 StartTimer_RunLabel_AssignedAtomically
02 StartTimer_RunLabel_BlankClearsPriorLabel
03 StartTimer_RunLabel_PreservedInCheckpointExport
04 TW_InstanceKey_UniqueAcrossRepeatedObjects
05 TW_FailedBegin_LocalStateRemainsInactive
06 TW_FailedUpdate_PriorMaskPreserved
07 Elapsed_Negative_StrictRaises
08 Elapsed_Negative_NonStrictReturnsZero
09 QPC_StartFailure_StrictRaises
10 QPC_StartFailure_NonStrictFallbackIsValid
11 QPC_EndFailure_NonStrictDoesNotFabricateElapsed
12 SystemTime_StartFailure_StrictRaises
13 SystemTime_StartFailure_NonStrictFallbackIsValid
14 SystemTime_EndFailure_NonStrictDoesNotApplyRolloverToFailure
15 TW_NoWorkbook_BeginDoesNotInventCalculationBaseline
16 TW_NoWorkbook_OpenWorkbook_EndDoesNotOverwriteCalculation
17 TW_NoWorkbook_FirstCalculationTouchCapturesRealBaseline
18 PM_TW_NewInstanceKey_BlankNeverIssued
19 PM_TW_EndSession_UnknownKey_PreservesEffectiveState
20 SourceVersion_TestVersion_AreSynchronized
```

---

# Appendix B — Suggested GitHub issues

1. `Do not commit zero QPC timestamp after non-strict start-read failure`
2. `Prevent timeGetSystemTime failure sentinel from entering rollover arithmetic`
3. `Track whether Application.Calculation baseline was actually captured`
4. `Add release/v1.2.0 regression cases and bump test suite version`
5. `Decouple core regression suite from demo builder/UI helpers`
6. `Add self-hosted Excel VBA regression workflow`
7. `Add hosted static VBA source/API/version checks`
8. `Protect main with required release checks`
9. `Build release workbook from tagged source and publish SHA-256 manifest`
10. `Document maximum safe elapsed interval by timing backend`
11. `Remove Long narrowing from ElapsedTime formatter`
12. `Move cPM_Report_WriteToRange into a dedicated reporting module`
13. `Expose timer/pause/error enums while preserving numeric compatibility`
14. `Add reproducible timing-overhead characterization`
15. `Update README, examples and CHANGELOG for v1.2.0`

---

# Appendix C — Evidence confidence

| Conclusion | Confidence |
|---|---|
| Exact branch/commit topology | High |
| 1.2 source changes | High |
| Test tree unchanged | High |
| 41-case test inventory | High |
| No Actions workflows | High |
| No branch protection | High |
| Non-strict QPC zero-start defect | High from source control flow |
| Method-4 failure/rollover defect | High from source control flow |
| Workbook-less Calculation baseline issue | High from source state model |
| Current regression runtime result | Not independently executed |
| 32-bit Office runtime behavior | Not independently executed |
| Actual QPC/timeGetSystemTime failure frequency | Not assessed |
| Runtime performance | Medium; source/algorithm review only |

---

# Appendix D — Release checklist for `v1.2.0`

```text
SOURCE
[ ] cPerformanceManager version = 1.2.0
[ ] M_cPM_TIMEWASTERS version = 1.2.0
[ ] test suite version = 1.2.0
[ ] README documents 1.2.0
[ ] CHANGELOG documents 1.2.0

CORRECTNESS
[ ] non-strict QPC failure cannot commit zero start
[ ] method-4 failure cannot enter rollover arithmetic
[ ] Calculation baseline has explicit validity state
[ ] rolling-counter supported-duration contract documented

TESTS
[ ] 1.2 cases added
[ ] fault injection available
[ ] suite passes on Office 64-bit
[ ] suite passes on Office 32-bit or limitation documented

CI
[ ] static source gate green
[ ] Excel regression gate green
[ ] required checks enabled

RELEASE
[ ] exact tag commit recorded
[ ] workbook assembled from tag source
[ ] workbook compiled and tested
[ ] SHA-256 manifest published
[ ] release notes include compatibility statement
[ ] release asset treated as immutable
```
