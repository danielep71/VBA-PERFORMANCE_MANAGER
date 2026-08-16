# Independent Code Review — VBA Performance Manager v1.2.0

> **Repository:** [`danielep71/VBA-PERFORMANCE_MANAGER`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER)  
> **Published release:** [`v1.2.0`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.2.0)  
> **Tag / reviewed commit:** [`7dd472f683f2a65c1cc9da2b76874f053a37bc67`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/commit/7dd472f683f2a65c1cc9da2b76874f053a37bc67)  
> **Release PR:** [`#16 — Release v1.2.0`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/16)  
> **Release date:** 2026-08-16  
> **Review date:** 2026-08-16  
> **Reviewer:** OpenAI GPT-5.6 Pro  
> **Suggested repository path:** `docs/INDEPENDENT_CODE_REVIEW_V1.2.0_PUBLISHED_2026-08-16.md`

---

# 1. Executive assessment

## Overall repository score: **8.9 / 10**

## Production runtime implementation score: **9.0 / 10**

## Timing/session core score: **9.3 / 10**

## Regression-assurance score: **9.3 / 10**

## Documentation score: **9.0 / 10**

## CI and release-engineering score: **7.6 / 10**

`VBA-PERFORMANCE_MANAGER` v1.2.0 is a credible, professional-quality pure-VBA runtime component.

It is materially more than a convenience timer. The release combines:

- six timing backends behind one session-bound interface;
- QueryPerformanceCounter as the preferred high-resolution backend;
- transactional timing-session start;
- explicit native-read status;
- strict and non-strict failure policies;
- unsigned 32-bit counter conversion and rollover handling;
- shared, multi-instance Excel Application-state suppression;
- timer-resolution lifecycle management;
- structured named checkpoints;
- machine-readable and human-readable reports;
- repeated procedure measurement;
- descriptive timing statistics;
- deterministic failure injection for otherwise unreachable native-read paths;
- a substantial Excel regression suite;
- a tagged release workbook with a published SHA-256 digest.

The release tag is attached to the merge commit of PR #16:

```text
7dd472f683f2a65c1cc9da2b76874f053a37bc67
```

The same commit is the current `main` head.

The published release includes:

```text
PERFORMANCE.MANAGER.xlsm
Size:    3,008,479 bytes
SHA-256: 53313bdffc419ad3aad18611077aaa78441453ba97996c07a4273893363575bc
```

The release records:

```text
63 regression cases
431 assertions
0 failures

Excel for Microsoft 365 MSO Version 2606
Build 16.0.20131.20152
64-bit
```

That is meaningful release evidence.

The strongest engineering areas are:

1. **Transactional timing start.** A session is not committed until a valid start timestamp exists.
2. **Native-read isolation.** QPC and method-4 reads use status-bearing `TryRead` contracts.
3. **Failure-safe elapsed arithmetic.** Failed reads do not enter subtraction or rollover correction.
4. **Observable non-strict behavior.** `LastReadStatus` distinguishes a failed read from a genuine zero-duration result.
5. **Shared TW ownership.** Overlapping instances coordinate process-global Excel state through aggregate masks.
6. **Checkpoint integrity.** A failed checkpoint read appends no row and advances no marker.
7. **Fault-injected regression coverage.** QPC failure, method-4 failure and invalid method-4 format are tested deterministically.
8. **Release documentation.** README, CHANGELOG and release notes describe the public surface and important limitations in substantial detail.

No P1 release-blocking defect was identified in the tagged production source.

The main residual issues are P2/P3 hardening items rather than reasons to withdraw the release:

- a strict endpoint-read failure raises before `LastReadStatus` is updated by the public wrapper;
- the non-strict repeated-measurement harness can record a failed read as a zero sample without returning a parallel status;
- a non-strict negative elapsed value is clamped to zero without changing `LastReadStatus`;
- the statistics API is broader in wording than in numerical contract;
- `Application.Run` procedure resolution remains caller-controlled and potentially ambiguous;
- Calculation suppression requires the documented stable-workbook-set invariant;
- only 64-bit Excel execution evidence is published;
- the release asset has a digest but no complete source-to-workbook build manifest;
- no automated GitHub Actions or protected status checks exist;
- the class and test modules are now physically large and would benefit from internal decomposition.

## Independent verdict

> **v1.2.0 is a sound published release with a 9.0-level production implementation and unusually strong regression discipline for VBA. The repository as a whole remains just below 9 because release verification is manual, 32-bit evidence is absent, artifact provenance is partial, and several status/statistics contracts need tightening.**

## What is required to move the overall repository above 9

The shortest route is:

1. fix strict-read status propagation;
2. prevent non-strict harness failures from becoming ordinary zero samples;
3. make negative-elapsed clamping status-bearing;
4. add a headless Excel regression workflow;
5. protect `main` with required checks;
6. publish 32-bit execution evidence;
7. publish a source/artifact manifest and reproducible workbook-build procedure;
8. narrow or harden the statistics contract;
9. qualify `Application.Run` targets;
10. split internal timing/statistics/reporting responsibilities without changing the public facade.

Completing items 1–7 should move the repository to approximately **9.1–9.2**. Completing all ten should place it around **9.4–9.5**.

---

# 2. Review scope and methodology

## 2.1 Exact revision basis

This review is tied to the immutable Git tag reference:

```text
Tag:    v1.2.0
Commit: 7dd472f683f2a65c1cc9da2b76874f053a37bc67
Date:   2026-08-16
```

The reviewed commit is:

- the merge commit of release PR #16;
- a GitHub-verified commit;
- the current `main` head;
- the commit referenced by the published v1.2.0 tag.

## 2.2 Production source reviewed

- [`src/classes/cPerformanceManager.cls`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/v1.2.0/src/classes/cPerformanceManager.cls)
- [`src/modules/M_cPM_TIMEWASTERS.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/v1.2.0/src/modules/M_cPM_TIMEWASTERS.bas)

## 2.3 Regression source reviewed

- [`test/M_cPM_Test.bas`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/v1.2.0/test/M_cPM_Test.bas)

## 2.4 Supporting material reviewed

- [`README.md`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/v1.2.0/README.md)
- [`CHANGELOG.md`](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/v1.2.0/CHANGELOG.md)
- `.gitignore`
- repository structure;
- release PR metadata;
- tag and release metadata;
- release-asset digest;
- branch protection;
- GitHub Actions inventory;
- current issue inventory.

## 2.5 Review methods

The review included:

- control-flow inspection;
- native API contract inspection;
- state-transition analysis;
- strict/non-strict behavior analysis;
- public API inventory;
- error/status contract inspection;
- rollover-domain analysis;
- TW lifecycle analysis;
- statistics-method review;
- test inventory and failure-injection review;
- source/documentation reconciliation;
- release-governance inspection;
- source and artifact metadata reconciliation.

## 2.6 Execution and artifact boundary

The published release provides a recorded Excel execution result:

```text
63 cases
431 assertions
0 failures
Excel Microsoft 365 MSO 2606
Build 16.0.20131.20152
64-bit
```

That result is treated as published release evidence.

The reviewer did **not** independently:

- rerun the Excel suite;
- reproduce the 431 assertions;
- execute Office 32-bit;
- extract the VBA project from the distributed `.xlsm`;
- compare the workbook’s embedded VBA byte-for-byte against the tagged source;
- reproduce the workbook build;
- measure timing distributions on the release machine.

Therefore:

- source-level conclusions are independently reviewed;
- release execution results are publisher-certified;
- binary/source equivalence is not independently established;
- 32-bit runtime compatibility remains source-supported but not release-certified.

---

# 3. Published release topology and evidence

## 3.1 Tag and merge

```text
v1.2.0
└── 7dd472f683f2a65c1cc9da2b76874f053a37bc67
    └── Merge pull request #16 — Release v1.2.0
```

The release therefore has:

- an explicit pull request;
- an exact merge commit;
- a signed/verified GitHub merge commit;
- a matching `main` head;
- a tag pointing to that commit.

This is a solid release topology.

## 3.2 Release artifact

| Field | Value |
|---|---|
| Asset | `PERFORMANCE.MANAGER.xlsm` |
| Content type | Excel macro-enabled workbook |
| Size | 3,008,479 bytes |
| SHA-256 | `53313bdffc419ad3aad18611077aaa78441453ba97996c07a4273893363575bc` |
| Release date | 2026-08-16 |
| Release type | Final, not prerelease |
| GitHub immutability flag | `false` |

Publishing the digest is a strong positive control.

The remaining provenance gap is that the release does not include a manifest showing:

```text
tag SHA
source file SHA-256 values
workbook SHA-256
build/import procedure
Excel build
Office bitness
test output
embedded-source equivalence
```

## 3.3 Regression certification

Published certification:

```text
Cases:      63
Assertions: 431
Failures:   0
Platform:   Excel for Microsoft 365
Version:    2606
Build:      16.0.20131.20152
Bitness:    64-bit
Date:       2026-08-16
```

This is materially stronger than an unqualified “tests pass” claim.

The principal missing certification dimension is Office 32-bit.

---

# 4. Hard repository metrics

## 4.1 Required production source

| File | Git blob SHA | Size |
|---|---|---:|
| `cPerformanceManager.cls` | `d6cff2a18c6880444281ce5cf36998ec5fecfe10` | 225,723 bytes |
| `M_cPM_TIMEWASTERS.bas` | `3bc8cc6ccd2b64a1e3e92da094cc8a6e33191546` | 55,993 bytes |
| **Required production total** |  | **281,716 bytes** |

Approximate physical scale:

```text
cPerformanceManager.cls:   ~6,040 lines
M_cPM_TIMEWASTERS.bas:     ~1,450 lines
Required runtime source:   ~7,490 lines
```

## 4.2 Regression source

| Metric | Value |
|---|---:|
| Test file | `M_cPM_Test.bas` |
| Git blob SHA | `ff84f7291ce63215dc8ed9404540b4de9a10c590` |
| Size | 362,006 bytes |
| Registered cases | 63 |
| Certified assertions | 431 |
| Certified failures | 0 |
| Approximate physical lines | ~9,000 |

## 4.3 Public surface

The release README reports:

```text
43 public members
24 methods
19 properties
```

The class surface spans:

- timing start and elapsed measurement;
- formatted elapsed output;
- session inspection;
- diagnostics;
- pause and cleanup;
- TW lifecycle;
- checkpoints and reports;
- repeated measurement;
- descriptive statistics;
- native-read status;
- Calculation-exemption status.

## 4.4 Timing backends

| ID | Backend | Primary role |
|---:|---|---|
| 1 | `Timer` | simple VBA clock, midnight wrap |
| 2 | `GetTickCount / GetTickCount64` | monotonic uptime timing |
| 3 | `timeGetTime` | millisecond multimedia timing |
| 4 | `timeGetSystemTime` | independent millisecond backend |
| 5 | `QueryPerformanceCounter` | default high-resolution timing |
| 6 | `Now() * 86400` | wall-clock diagnostics |

## 4.5 Governance state

```text
GitHub Actions workflows: 0
main protected:           no
required status checks:   none
release PR:               merged
tag:                      present
release:                  present
asset digest:             present
32-bit certification:     absent
```

---

# 5. Scoring methodology

A score of 9 or more requires:

- no known common-path silent wrong-result behavior;
- explicit failure semantics;
- deterministic lifecycle cleanup;
- executable failure-path tests;
- accurate documentation;
- a stable public contract;
- reproducible release evidence;
- appropriate platform certification;
- maintainable module boundaries.

A score of 10 additionally requires:

- automated Excel CI;
- recurring 32-bit and 64-bit certification;
- protected required checks;
- reproducible source-to-workbook build;
- immutable/signature-backed artifacts;
- numerically robust generic statistics or a tightly enforced timing-only domain;
- generated API/test/documentation inventories;
- focused internal modules.

---

# 6. Weighted scorecard

| Area | Weight | Score | Weighted contribution | What is needed to reach or remain above 9 |
|---|---:|---:|---:|---|
| Functional correctness | 17% | **9.1** | 1.547 | Propagate strict endpoint status; make negative clamping explicit; harden harness failure handling |
| Timing-source robustness | 13% | **9.3** | 1.209 | Roll back method-3 resolution acquisition on failed aligned start; publish 32-bit evidence |
| Architecture and modularity | 9% | **8.6** | 0.774 | Extract native timing, statistics, measurement and reporting internals behind the same facade |
| Public API and error contracts | 8% | **8.8** | 0.704 | Add timer/pause/error enums; expose consistent operation status; clarify QPC inspection units |
| TW lifecycle and host state | 9% | **8.9** | 0.801 | Automate or certify workbook-transition scenarios; optionally add lifecycle event tracking |
| Measurement and statistics | 10% | **8.3** | 0.830 | Qualify macro targets; define sample domain; prevent failed zero samples; harden CV/numerical methods |
| Checkpoints and reporting | 6% | **9.3** | 0.558 | Centralize report schema and preserve failure propagation; separate worksheet rendering |
| Regression testing | 12% | **9.3** | 1.116 | Add headless execution, setter failure injection, 32-bit runs and CI artifacts |
| Documentation and governance | 7% | **9.0** | 0.630 | Correct residual source-header/DLL wording; generate counts; add contribution/security policy |
| CI and release engineering | 6% | **7.6** | 0.456 | Add automated Excel CI, branch protection and a complete provenance manifest |
| Maintainability and security | 3% | **8.8** | 0.264 | Conditionalize test seams, split oversized modules, document dynamic invocation trust |
| **Total** | **100%** |  | **8.889 / 10** |  |

Rounded score:

```text
8.9 / 10
```

---

# 7. Component scores

| Component | Score | Assessment | Route to 9+ |
|---|---:|---|---|
| `cPerformanceManager.cls` | **9.0** | Strong public facade and timing semantics; physically oversized | Split private implementation while preserving public API |
| Core timing/session engine | **9.3** | Transactional start, read-first arithmetic, explicit status | Fix strict endpoint status and method-3 environmental rollback |
| QPC backend | **9.5** | Correct TryRead ownership, fallback-before-commit, injected failures | Publish platform characterization and 32-bit evidence |
| Method-4 backend | **9.5** | Read/format failures cannot enter arithmetic; deterministic tests | Retain single-reader invariant |
| Methods 1–3 and 6 | **8.8** | Explicit limits and rollover logic; Timer ambiguity remains | Add status for backward/clamped clocks and 32-bit runtime tests |
| `LastReadStatus` contract | **8.7** | Major usability improvement; strict endpoint and clamp gaps remain | Set status before strict raise; add validation-status values |
| `M_cPM_TIMEWASTERS.bas` | **8.9** | Strong aggregate state and baseline validity | Automate stable-host lifecycle evidence; move report writer |
| Checkpoint subsystem | **9.4** | Failed reads abandon cleanly; cache isolation; geometric growth | Add report schema/version constant |
| Reporting | **9.1** | Useful array/text/range surfaces | Separate worksheet writer and improve large-text construction |
| Measurement harness | **8.4** | Raw vectors and warm-up are valuable | Reject failed samples; qualify procedures; cap allocations |
| Statistics | **8.2** | Good for normal positive timing samples | Narrow domain or use stable generic algorithms |
| Regression suite | **9.3** | Broad, fault-injected, certified 63/431/0 | Headless CI, setter injection, 32-bit run |
| README / CHANGELOG | **9.0** | Comprehensive, synchronized with release | Correct residual overclaims and generate inventories |
| Release artifact | **8.5** | Tag, PR, digest and Excel certification | Reproducible build, source manifest, immutable/signature policy |
| Repository governance | **7.4** | Release discipline exists, but automation/protection do not | Workflows, required checks, protected branch |
| Security/platform posture | **8.8** | Narrow native surface, no third-party runtime | Source-bind artifact, SECURITY.md, trusted-macro documentation |

---

# 8. Architectural review

## 8.1 Logical architecture

```text
cPerformanceManager
    |
    +-- timing session facade
    +-- native API declarations
    +-- timing-source readers
    +-- read status / error mapping
    +-- rollover and elapsed arithmetic
    +-- timer-resolution lifecycle
    +-- pause strategies
    +-- diagnostics
    +-- TW facade
    +-- checkpoints
    +-- report serialization
    +-- repeated measurement
    +-- statistics
    +-- regression fault seams
            |
            v
M_cPM_TIMEWASTERS
    |
    +-- process-wide session registry
    +-- per-instance disable masks
    +-- per-instance host strictness
    +-- aggregate state
    +-- baseline capture / validity
    +-- restore / rollback
    +-- Calculation exemption state
    +-- worksheet report rendering
```

## 8.2 Architectural strengths

### Session-bound timing

`StartTimer` binds a backend, and elapsed reads are validated against that backend.

This prevents one of the most common timing errors: starting with one clock and ending with another.

### Transactional session start

`StartTimer` captures into local variables and commits only after successful capture.

Strict failure preserves the current session.

Non-strict native-read failure can switch to method 2 before state is committed.

### Single-site native reads

`QPC_TryReadTick` is the only QPC read site.

`SystemTime_TryReadMs` is the only `timeGetSystemTime` read site.

That design makes failure handling reviewable and testable.

### Single-site elapsed dispatch

`Elapsed_ComputeSeconds` centralizes elapsed arithmetic for all backends.

`ElapsedSeconds` caches successful results.

`Checkpoint` uses the same computation without mutating the cache.

### Shared Application-state ownership

Excel Application flags are correctly treated as process-global state.

The module-level manager aggregates every active participant’s request rather than allowing each object to restore state independently.

### Full samples before aggregation

The measurement API returns `Double()` vectors.

Statistics operate afterward.

That preserves transparency and allows alternate analysis.

## 8.3 Architectural weaknesses

### Class size and responsibility breadth

The class is approximately 225 KB and about 6,000 physical lines.

It owns:

- platform declarations;
- six clocks;
- session state;
- error/status semantics;
- fault injection;
- pause;
- diagnostics;
- TW facade;
- checkpoints;
- reports;
- measurement harness;
- statistics;
- sorting.

The conceptual design remains coherent, but physical review cost is high.

### Reporting inside the TW manager

`cPM_Report_WriteToRange` is presentation logic inside the process-state manager.

This mixes independent responsibilities.

### Test seams in distributed production code

The three fault-injection procedures are `Friend`.

They are invisible outside the VBA project but become callable by any module inside a consuming project after import.

They are well designed as one-shot flags, but they remain test support in the production class.

## To reach 9+ in architecture

Preserve the public `cPerformanceManager` facade and extract:

```text
M_cPM_NATIVE_TIMING
M_cPM_MEASUREMENT
M_cPM_STATISTICS
M_cPM_REPORTING
```

Recommended boundaries:

- native API declarations and raw reads;
- pure rollover/unsigned arithmetic;
- sample generation and procedure invocation;
- pure statistical algorithms;
- worksheet rendering.

Use conditional compilation for test hooks:

```vb
#If CPM_TESTING Then
    Friend Sub Test_...
#End If
```

This can be done without breaking callers.

---

# 9. Core timing and session review

## 9.1 `StartTimer` — **9.5 / 10**

The implementation establishes an excellent invariant:

> No session is committed unless the selected or fallback backend produced a valid start timestamp.

The sequence is:

```text
normalize requested method
attempt capture
if strict failure:
    set failure status
    raise
    preserve current session
if non-strict failure:
    switch to method 2
    capture valid method-2 timestamp
commit method actually used
reset elapsed/checkpoint state
apply run label
```

This is strong transactional design.

### Residual issue: method-3 environmental side effect

For an aligned method-3 start:

```text
EnsureTimerResolution1ms
then poll for next tick
```

If strict alignment reaches the spin timeout, `StartTimer` does not commit a timing session, but the instance may retain the newly acquired timer-resolution request until `ResetEnvironment` or destruction.

That is not a wrong measurement, but it weakens full transactionality of environmental side effects.

### To reach 10

Track whether the current start attempt newly acquired `timeBeginPeriod(1)` and release it if capture fails before commit.

---

## 9.2 `QPC_TryReadTick` — **9.6 / 10**

Strengths:

- one native call site;
- reads into a local;
- output parameter written only on success;
- Boolean success contract;
- no sentinel timestamp;
- one-shot test injection;
- no hidden sticky failure state.

This is an excellent low-level contract.

---

## 9.3 `SystemTime_TryReadMs` — **9.6 / 10**

Strengths:

- one native call site;
- API failure separated from invalid returned format;
- output written only on full success;
- no failed value enters arithmetic;
- one-shot API-failure injection;
- one-shot invalid-format injection;
- strict/non-strict behavior tested.

The tagged suite includes a dedicated invalid-format regression that verifies:

- distinct status;
- zero non-strict return;
- no rollover inflation;
- cache preservation;
- strict raise;
- self-clearing test flag.

That is high-quality failure-path coverage.

---

## 9.4 `Elapsed_ComputeSeconds` — **9.2 / 10**

The critical invariant is correct:

```text
read
validate read
only then subtract
only then correct rollover
only then publish endpoint
```

Failed QPC or method-4 reads:

- raise in strict mode;
- return zero with non-OK status in non-strict mode;
- do not overwrite endpoints;
- do not enter rollover arithmetic.

### Residual issue: strict failure status propagation

`ElapsedSeconds` resets:

```vb
m_LastReadStatus = cPM_ReadOK
```

then calls `Elapsed_ComputeSeconds`.

On strict QPC or method-4 failure, the private function raises immediately.

The public wrapper never reaches:

```vb
m_LastReadStatus = Status
```

Therefore, after a caller catches the strict-mode error, `LastReadStatus` may still report `cPM_ReadOK`.

This conflicts with the property’s description as the outcome of the most recent native read.

The error number and message are correct, so this is not a silent failure. It is an API-consistency defect.

### Recommended fix

Make the private function return status without raising.

Then let the public operation:

1. assign `m_LastReadStatus`;
2. raise if strict;
3. return zero if non-strict.

Alternatively, add a local error handler that sets status before rethrowing.

---

## 9.5 Non-strict negative elapsed — **8.3 / 10**

`Elapsed_Validate`:

```text
strict negative -> raise
non-strict negative -> clamp to zero
```

The clamp does not update `LastReadStatus`.

The public operation can therefore return:

```text
0 seconds
LastReadStatus = cPM_ReadOK
```

even though the result was produced by a backward-moving source rather than by a valid zero interval.

Consequences:

- callers cannot distinguish this from a true zero;
- `Checkpoint` can treat it as a valid capture;
- a checkpoint can append cumulative zero and affect the next delta;
- the README statement that failed/non-valid zero results are distinguishable is broader than the current status model.

### Recommended fix

Extend the enum:

```vb
cPM_ReadClockMovedBackward
cPM_ReadElapsedClamped
```

Better still, make `Elapsed_Validate` return both value and status.

For checkpoints, any non-OK status should abandon capture.

---

## 9.6 Supported duration domains

The release correctly documents:

- backend 1 cannot distinguish clock rollback from midnight;
- rollover correction handles one wrap.

Practical safe intervals:

```text
Timer:                    < 24 hours
GetTickCount, Win32:      < 49.71 days
timeGetTime:              < 49.71 days
timeGetSystemTime:        < 49.71 days
GetTickCount64, Win64:    effectively unbounded for normal use
QPC:                      preferred for long/high-resolution sessions
```

The documentation is appropriately conservative.

---

# 10. Read-status and error-contract review

## 10.1 `cPM_ReadStatus` — **8.8 / 10**

Current statuses:

```text
cPM_ReadOK
cPM_ReadQpcFailed
cPM_ReadSystemTimeFailed
cPM_ReadSystemTimeFormatInvalid
cPM_ReadFallbackToMethod2
```

This gives non-strict timing a machine-readable outcome.

The distinction between:

```text
valid zero
failed read
valid fallback session
```

is valuable.

## 10.2 Status scope

`LastReadStatus` is intentionally operation-scoped.

A successful later read resets it to OK.

That is a reasonable design, and `ActiveMethodID` continues to reveal a fallback-bound method.

## 10.3 Contract gaps

The enum covers native-read outcomes but not:

- no active session in non-strict mode;
- invalid explicit method coerced to active method;
- negative elapsed clamped to zero;
- alignment timeout fallback;
- timer-resolution request failure tolerated in non-strict mode.

These are meaningful non-strict outcomes.

## 10.4 Error constants

Runtime error numbers are centralized and named, which is strong internally.

They remain private, so consumers cannot write symbolic handlers.

## To reach 9+ in API/error contracts

Add:

```vb
Public Enum cPM_TimerMethod
Public Enum cPM_PauseMethod
Public Enum cPM_ErrorCode
Public Enum cPM_OperationStatus
```

Keep numeric compatibility.

Either expand `cPM_ReadStatus` or add a broader `LastOperationStatus`.

Expose the backend actually requested and the backend actually bound if both are useful.

---

# 11. Raw timestamp and formatter review

## 11.1 T1/T2 unit semantics — **8.2 / 10**

For QPC sessions, documentation describes:

```text
CDbl(m_qpcStart) = raw QPC ticks
```

However, `Currency` is a scaled 64-bit representation.

`CDbl(Currency)` produces the scaled numeric interpretation, not the original unscaled integer tick count.

Elapsed arithmetic is correct because identical scaling cancels between tick delta and frequency.

The issue is diagnostic terminology, not timing correctness.

### Recommended fix

Document:

```text
Currency-scaled QPC value
```

or expose:

```text
QPC_StartTicksCurrency
QPC_EndTicksCurrency
QPC_StartTicksAsDouble = CDbl(m_qpcStart) * 10000#
```

## 11.2 Stale T2/ET comments

The release correctly implements:

> Checkpoint does not overwrite cached T2 or ET.

The T2 and ET property headers still say that Checkpoint updates them.

That contradicts both the implementation and CHANGELOG.

This should be corrected in the next documentation patch.

## 11.3 `ElapsedTime` narrowing — **8.4 / 10**

The formatter uses:

```vb
WholeS As Long
WholeS = CLng(Fix(Sec))
```

This creates an undocumented ceiling near:

```text
2,147,483,647 seconds
approximately 68.1 years
```

It also relies on implicit `CDbl` conversion for a supplied Variant.

This is not relevant to ordinary benchmarks, but it is a public input-contract boundary.

### Recommended fix

- validate numeric input explicitly;
- reject errors/empty/nonnumeric values with a named error;
- keep whole duration in `Double` or `Currency`;
- add a large-duration regression.

---

# 12. Shared TW manager review

## 12.1 Aggregate ownership model — **9.3 / 10**

The core model is strong:

```text
first session captures baseline
each instance registers disable mask
each instance registers host strictness
aggregate mask = OR of active masks
strictest active requirement wins
last session restores baseline
```

This correctly models process-global Excel state.

## 12.2 Instance identity — **9.1 / 10**

A module-level counter replaces object-address identity.

That avoids heap-address reuse collisions.

The remaining theoretical issue is `Long` overflow after more than two billion allocations in one project lifetime.

That is not practically material.

## 12.3 Begin/end rollback — **9.1 / 10**

Begin and end capture prior dictionary state and attempt rollback if Excel property application fails.

The original error is preserved and re-raised.

This is professional error handling for VBA.

## 12.4 Calculation baseline validity — **9.0 / 10**

The manager distinguishes:

```text
baseline value
baseline validity
```

It does not write a synthetic Automatic value to a live workbook.

When Calculation control cannot be honored:

- non-strict mode exempts the flag;
- exemption is observable;
- strict participation raises on begin;
- no lazy baseline is invented later.

This is defensible and conservative.

## 12.5 Stable-host invariant

Calculation control requires the open-workbook set to remain stable for the suppression scope.

The release documents that genuinely workbook-less transitions are manually verified only.

This is an explicit supported-domain constraint rather than a hidden assumption.

### Remaining operational edge

If a real Calculation baseline is captured, all workbooks are closed before final restore, and the scope ends, the manager cannot write the baseline because Excel exposes no writable Calculation state without a workbook.

It records exemption and clears shared state.

The invariant makes this caller misuse, but strict mode cannot intercept every after-start lifecycle transition without workbook event tracking.

### To reach 9.5+

Choose one:

1. publish a repeatable second-Excel-instance manual certification protocol; or
2. add workbook lifecycle interception; or
3. track a pending restore and reapply it when a workbook becomes available.

For the current release, the explicit invariant is acceptable.

---

# 13. Checkpoint and reporting review

## 13.1 Checkpoint capture — **9.4 / 10**

Strengths:

- requires an active session;
- uses the central elapsed reader;
- does not mutate cached `T2`/`ET`;
- abandons failed native reads;
- appends no failed row;
- leaves count unchanged;
- leaves last-success marker unchanged;
- preserves correct next delta;
- supports run label and note;
- stores backend metadata;
- grows storage geometrically.

The injected regression verifies that a failed checkpoint does not distort the following checkpoint.

This is one of the best subsystems in the release.

### Residual gap

A non-strict negative elapsed clamp currently returns OK status and can still be appended.

Fixing status-bearing validation resolves this.

## 13.2 `ReportAsArray` — **9.3 / 10**

Strengths:

- deterministic eight-column schema;
- header-only empty output;
- 1-based 2-D array;
- direct worksheet compatibility;
- no caller-specific formatting assumptions.

## 13.3 `ReportAsText` — **8.9 / 10**

The output is readable and appropriate for diagnostics.

Repeated string concatenation may scale poorly for very large checkpoint sets.

Use a string array plus `Join` for high-volume reports.

## 13.4 Worksheet writer — **8.8 / 10**

The helper is useful, validates inputs, formats numeric columns and remains hidden from the Macro dialog through `Option Private Module`.

Its location is the concern: it belongs in a reporting module, not the TW state manager.

---

# 14. Measurement-harness review

## 14.1 `MeasureProcedure` — **8.5 / 10**

Strengths:

- raw sample vector;
- warm-up iterations;
- isolated worker;
- selected backend;
- strict-mode policy mirrored;
- caller session/checkpoints unaffected;
- measured-procedure errors preserved after cleanup;
- no claim that backend-only overhead is a matched dispatch baseline.

## 14.2 Macro resolution

The harness calls:

```vb
Application.Run ProcedureName
```

A bare name can resolve differently when:

- multiple open workbooks contain the same procedure;
- an add-in exposes the same name;
- workbook focus changes;
- the caller passes an unintended but valid public macro.

### Recommended fix

Accept one of:

```text
fully qualified macro string
target Workbook + module + procedure
target workbook name + procedure
```

Example:

```text
'Model.xlsm'!M_Benchmark.RebuildCurve
```

Document the procedure name as a trusted dynamic-invocation boundary.

## 14.3 Non-strict failed samples

The worker mirrors the caller’s `StrictMode`.

In non-strict mode:

```text
native endpoint read fails
ElapsedSeconds returns 0
worker LastReadStatus becomes failure
sample array receives 0
worker is discarded
```

The caller cannot inspect the worker’s status.

The zero sample is indistinguishable from an actual measured zero and can distort:

- minimum;
- median for small samples;
- mean;
- standard deviation;
- CV;
- contamination judgment.

This is the most important remaining measurement-harness defect.

### Recommended behavior

Choose one:

1. force the worker to strict mode regardless of caller policy;
2. raise a named harness error if `worker.LastReadStatus <> OK`;
3. return a parallel status vector;
4. return a structured result containing values and statuses;
5. omit failed samples and report failure count.

The safest backward-compatible behavior is:

```text
harness treats every failed sample as fatal
```

## 14.4 Iteration bounds

Iterations are only coerced at the lower bound.

A very large request can allocate a large array and block Excel for an impractical period.

Add a practical upper bound or a named allocation error.

## 14.5 Matched baseline

The release correctly states that `MeasureOverhead_Samples` is not a matched `Application.Run` baseline.

For very short procedures, callers should measure an empty public Sub through `MeasureProcedure`.

A future `MeasureDispatchOverhead_Samples` helper would make this easier.

---

# 15. Statistics review

## 15.1 Overall assessment — **8.2 / 10**

The statistics surface is useful for ordinary nonnegative timing observations.

It is not yet a robust general-purpose numerical statistics package.

## 15.2 Strengths

- raw samples retained;
- min/max;
- arithmetic mean;
- median;
- nearest-rank percentile;
- sample standard deviation;
- coefficient of variation;
- contamination heuristic;
- readable text summary;
- sorted copy preserves caller order;
- empty-array behavior;
- one-sample spread behavior;
- known-vector regression.

## 15.3 Public-domain mismatch

README and CHANGELOG describe the routines as operating on “any `Double()` vector.”

The implementation is conceptually timing-oriented.

Problem cases include:

### Zero-mean, nonzero-spread data

```text
[-1, +1]
mean = 0
stddev > 0
CV returns 0
contamination returns false
```

### Negative-mean data

```text
CV = stddev / negative mean
```

The result is negative and cannot exceed a positive contamination threshold.

### Extreme finite magnitudes

Naive sum can overflow or lose precision.

Even-sample median:

```text
(a + b) / 2
```

can overflow when the correct midpoint is representable.

Squared deviations can overflow.

### Negative threshold

`Stats_IsContaminated` does not reject a negative threshold.

## 15.4 Recommended contract

The simplest high-quality path is to narrow the statistics API to:

```text
finite
nonnegative
timing observations
moderate magnitude
```

Then:

- reject negative samples;
- validate threshold >= 0;
- define all-zero CV as 0;
- document generic use as unsupported.

If a general numerical contract is desired, use:

- compensated summation or incremental mean;
- safe midpoint;
- Welford/scaled variance;
- `Abs(mean)` for CV;
- explicit near-zero-mean behavior;
- nonfinite/magnitude validation.

## 15.5 Interpretation guidance

Median-first reporting is sensible.

Minimum should be described carefully:

- it is the least disturbed observed run;
- it can also reflect clock quantization or an unrepresentative cache state;
- it is not automatically the “true cost.”

P95-versus-median is a useful tail diagnostic.

CV is a heuristic, not a formal contamination test.

The release appropriately states that.

---

# 16. Regression-suite review

## 16.1 Score: **9.3 / 10**

The suite is unusually substantial for a VBA component.

Published result:

```text
63 cases
431 assertions
0 failures
Excel Microsoft 365 64-bit
```

## 16.2 Coverage strengths

The suite covers:

### Core timing

- constructor state;
- method names;
- all six start paths;
- all six elapsed paths;
- formatting;
- aligned starts;
- QPC accessors;
- method-4 behavior.

### Validation

- elapsed before start;
- method mismatch;
- invalid start;
- non-strict coercion;
- negative elapsed validation.

### Native arithmetic

- unsigned 32-bit boundaries;
- rollover periods;
- Win64 conditional behavior.

### Native failure injection

- QPC strict start failure;
- QPC non-strict fallback;
- QPC elapsed failure;
- method-4 start failure;
- method-4 elapsed failure;
- method-4 invalid format;
- cache preservation;
- self-clearing injection flags.

### TW lifecycle

- blank keys;
- single instance;
- overlapping instances;
- reset idempotence;
- destructor cleanup;
- instance-key reuse protection;
- Calculation baseline validity;
- deliberate exemption;
- overlapping Calculation scopes;
- no synthetic baseline.

### Checkpoints

- no-session error;
- labels;
- defaults;
- array/text output;
- clear/reuse;
- cumulative/delta semantics;
- cache preservation;
- 1,000-checkpoint capacity growth;
- failed-read abandonment.

### Measurement/statistics

- overhead samples;
- repeated procedure measurement;
- known-vector statistics;
- percentile and boundary errors;
- caller-session preservation.

## 16.3 Remaining gaps

### No setter-failure injection

Rollback is not deterministically tested for failures while setting:

```text
ScreenUpdating
EnableEvents
DisplayAlerts
Calculation
Cursor
```

### No automated headless runner

The suite depends on demo/UI helpers and a worksheet log.

### No Office 32-bit execution evidence

Conditional branches compile differently on 32-bit Office.

### Genuinely workbook-less paths

A test workbook cannot close the final workbook and continue the same suite normally.

These paths remain manual.

### Time-dependent flakiness

Pause and elapsed tests rely on real timing.

The suite appears to use broad sanity ranges, which is appropriate, but execution remains environment-dependent.

## To reach 9.5+

- split core assertions from worksheet presentation;
- add machine-readable final result;
- add Application-setter fault injection;
- run in CI on a dedicated Excel runner;
- publish periodic Office 32-bit evidence;
- automate a second Excel instance for workbook-less lifecycle scenarios.

---

# 17. Release-engineering review

## 17.1 Strengths

The published release has:

- a release branch;
- a merged release PR;
- an exact merge commit;
- a verified GitHub commit;
- a matching tag;
- a final release;
- a macro-enabled workbook asset;
- an asset SHA-256 digest;
- explicit Excel version/build/bitness;
- explicit case/assertion/failure counts;
- source-first Git history;
- no versioned workbook binary in ordinary source history.

This is a meaningful release process.

## 17.2 Gaps

### No automated workflows

The repository has zero GitHub Actions workflows.

### No branch protection

`main` is unprotected and has no required checks.

### Manual certification

The Excel test result is recorded, but no workflow artifact binds machine-readable output to the commit.

### Partial artifact provenance

The asset digest is published, but there is no build manifest binding:

```text
tagged source
import process
compiled workbook
test result
release asset
```

### Release mutability

GitHub reports the release as not immutable.

### No 32-bit release certificate

The README supports 32/64-bit, but the published certification is 64-bit only.

## 17.3 Recommended two-layer CI

### Hosted static workflow

- required files;
- module names;
- `Option Explicit`;
- version consistency;
- API inventory;
- test registration count;
- README metrics;
- CHANGELOG/tag consistency;
- no conflict markers;
- no binary files in source;
- no reintroduction of object-address keys;
- no bare runtime error offsets.

### Windows/Excel workflow

On a dedicated runner:

1. create/open clean `.xlsm`;
2. import tagged source;
3. compile;
4. run headless regression;
5. emit JSON/text counters;
6. fail on compile/runtime/assertion failure;
7. upload evidence;
8. build release workbook;
9. compute SHA-256;
10. publish manifest.

## 17.4 Score path

Current release engineering:

```text
7.6 / 10
```

Add static CI + headless Excel execution:

```text
8.8–9.0
```

Add protected required checks + reproducible artifact:

```text
9.2
```

Add dual-bitness recurring certification + signing/immutability:

```text
9.5+
```

---

# 18. Documentation and governance review

## 18.1 Strengths

The README is comprehensive and usable.

It covers:

- product positioning;
- installation;
- backends;
- quick start;
- measurement;
- statistics;
- checkpoints;
- TW control;
- strict mode;
- API inventory;
- known limitations;
- regression metrics;
- release version.

The CHANGELOG is specific and includes:

- added APIs;
- behavioral changes;
- fixed correctness defects;
- test inventory;
- certification environment;
- known limitations.

## 18.2 Residual issues

### “No External DLL”

The code uses Windows system libraries:

```text
kernel32
winmm.dll
```

The accurate claim is:

```text
no third-party DLL
no bundled native dependency
```

### Exact Application-state restore

The comparison table presents exact restoration as an unconditional capability.

Calculation restoration is subject to the documented stable-host invariant.

The table should use a footnote or narrower wording.

### Backwards-clock detection table

The introductory note correctly limits backward-clock detection to supported backends, but the capability table remains absolute.

Add a footnote for backend 1.

### T2/ET source headers

They state that Checkpoint updates the caches.

The implementation and CHANGELOG state the opposite.

### “Any Double vector”

This is stronger than the numerical contract actually supported.

### Documentation percentage

“76% documentation” is a manually maintained metric unless generated.

Generated metrics would be preferable.

## 18.3 Governance documents

The tagged root does not include:

- `SECURITY.md`;
- `CONTRIBUTING.md`;
- a formal release checklist;
- a machine-readable compatibility matrix.

These are not required for a personal repository, but they matter for a 9+ public component.

---

# 19. Maintainability review

## 19.1 Strengths

- consistent naming;
- extensive procedure contracts;
- named runtime errors;
- centralized elapsed reader;
- centralized native readers;
- centralized checkpoint reset;
- source-first repository;
- binary demo removed from Git history;
- late-bound dictionary;
- no opaque dependency framework.

## 19.2 Pressure points

### Class size

At 225,723 bytes, the class is large.

### Test-module size

At 362,006 bytes, the test module is difficult to navigate.

### Repeated documentation boilerplate

Detailed contracts are helpful, but repetitive headers increase diff noise.

### Manual inventories

API/test/documentation counts remain manually synchronized.

### Mixed responsibilities

Statistics and report rendering are separable from timing-session state.

## To reach 9+

- split tests by domain;
- extract stateless internals;
- generate inventories;
- keep concise contracts near code;
- move long conceptual material to docs;
- add source-header and version checks.

---

# 20. Security and platform assessment

## 20.1 Positive characteristics

No high-severity security vulnerability was identified.

The runtime has:

- no network access;
- no shell execution;
- no production filesystem mutation;
- no third-party DLL;
- no installer;
- no COM registration;
- narrow Windows API use;
- late-bound Scripting.Dictionary;
- internal support hidden from Macro dialog;
- explicit cleanup;
- no modal MsgBox-based runtime error path.

## 20.2 Dynamic invocation boundary

`MeasureProcedure` executes a caller-supplied procedure string through `Application.Run`.

This is intentional but should be treated as a trusted-input boundary.

A public library should state:

> Only pass trusted, fully qualified macro names.

## 20.3 Fault-injection exposure

Friend test hooks become callable from any module in the importing VBA project.

A malicious or accidental call can force the next timing read to fail.

The hooks are one-shot and not externally exposed through COM, so risk is low.

Conditional compilation would remove them from distribution builds.

## 20.4 Release workbook trust

The published digest improves integrity verification.

A stronger chain would add:

- build manifest;
- VBA project signature;
- source hash manifest;
- reproducible import/build instructions;
- immutable release policy.

---

# 21. Findings summary

| ID | Severity | Area | Finding |
|---|---|---|---|
| CPM120-P2-01 | **P2** | Status contract | Strict endpoint failures can raise before `LastReadStatus` records the failed read |
| CPM120-P2-02 | **P2** | Measurement harness | Non-strict worker read failure can enter a sample vector as an ordinary zero |
| CPM120-P2-03 | **P2** | Elapsed validation | Non-strict backward/negative elapsed clamp returns zero with OK read status |
| CPM120-P2-04 | **P2** | Statistics | “Any Double vector” exceeds the supported CV and numerical domain |
| CPM120-P2-05 | **P2** | Procedure invocation | Unqualified `Application.Run` names can resolve ambiguously |
| CPM120-P2-06 | **P2** | Platform assurance | 32-bit compatibility is source-designed but not release-certified |
| CPM120-P2-07 | **P2** | CI / release | No automated Excel gate, required status check or protected main |
| CPM120-P2-08 | **P2** | Artifact provenance | Asset digest exists, but embedded source/build equivalence is not independently traceable |
| CPM120-P2-09 | **P2** | TW lifecycle | Calculation control depends on a stable open-workbook set |
| CPM120-P3-01 | **P3** | Start transaction | Method-3 aligned-start timeout can retain a newly acquired timer-resolution request |
| CPM120-P3-02 | **P3** | Diagnostics | QPC T1/T2 are described as raw ticks despite Currency scaling |
| CPM120-P3-03 | **P3** | Documentation | T2/ET headers incorrectly say Checkpoint updates them |
| CPM120-P3-04 | **P3** | Formatter | `ElapsedTime` narrows whole seconds to `Long` |
| CPM120-P3-05 | **P3** | Resource control | Measurement iterations have no practical upper bound |
| CPM120-P3-06 | **P3** | Test surface | Friend fault-injection hooks ship in production source |
| CPM120-P3-07 | **P3** | Architecture | The public class concentrates too many independent subsystems |
| CPM120-P3-08 | **P3** | Modularity | Worksheet report output lives in the TW manager |
| CPM120-P3-09 | **P3** | Key allocator | Instance-key seed has a theoretical `Long` overflow |
| CPM120-P3-10 | **P3** | Documentation | “No external DLL” should be “no third-party DLL” |
| CPM120-P3-11 | **P3** | Governance | No SECURITY / CONTRIBUTING / formal release-policy documents |
| CPM120-P3-12 | **P3** | Release policy | GitHub release is not marked immutable |

No P1 finding was identified.

---

# 22. Detailed principal findings

## CPM120-P2-01 — Strict endpoint failure does not reliably update `LastReadStatus`

### Affected surface

```text
ElapsedSeconds, strict mode
ElapsedTime, strict mode
Checkpoint, strict mode
QPC endpoint failure
method-4 endpoint failure
method-4 wrong-format endpoint
```

### Root cause

The public operation resets status to OK.

The private elapsed function raises before returning its status.

The public operation cannot publish the failure status.

### Impact

After catching the strict error:

```text
Err.Number = correct native-read error
LastReadStatus may still equal cPM_ReadOK
```

The error path is not silent, but the status property is inconsistent.

### Recommended fix

Move strict/non-strict policy to the public operation:

```text
private reader:
    never raises for read failure
    returns status

public operation:
    assign LastReadStatus
    if strict and status != OK:
        raise
```

### Acceptance criteria

```text
[ ] strict QPC endpoint failure sets cPM_ReadQpcFailed
[ ] strict method-4 endpoint failure sets cPM_ReadSystemTimeFailed
[ ] strict wrong-format sets cPM_ReadSystemTimeFormatInvalid
[ ] error number remains unchanged
[ ] cached good endpoint remains unchanged
[ ] checkpoint count remains unchanged
```

---

## CPM120-P2-02 — Non-strict harness can store failed reads as valid zero samples

### Affected surface

```text
MeasureProcedure
MeasureOverhead_Samples
Stats_Min
Stats_Median
Stats_Mean
Stats_StdDev
Stats_CoefficientOfVariation
Stats_IsContaminated
```

### Root cause

The worker mirrors non-strict mode and returns zero on failed endpoint reads.

The harness stores the result without inspecting the worker status.

### Impact

A rare native failure can look like the fastest run.

It can materially distort the minimum and small-sample statistics.

### Recommended fix

Backward-compatible option:

```text
After every worker.ElapsedSeconds:
    if worker.LastReadStatus != cPM_ReadOK:
        raise ERR_CPM_MEASUREMENT_READ_FAILED
```

A fallback-to-method-2 status at session start is usable; endpoint failure is not.

### Acceptance criteria

```text
[ ] no non-OK endpoint sample enters the vector
[ ] warm-up failures are also detected
[ ] error identifies iteration and status
[ ] caller session remains untouched
[ ] test injects a worker failure through a controllable test adapter
```

---

## CPM120-P2-03 — Negative elapsed clamp is not status-bearing

### Root cause

`Elapsed_Validate` returns only a number.

Non-strict mode changes negative elapsed to zero without emitting status.

### Impact

```text
zero + OK
```

can represent an invalid backward-clock event.

Checkpoint can treat it as valid.

### Recommended fix

Add:

```text
cPM_ReadElapsedInvalid
```

or a broader operation status.

Pass status by reference through validation.

### Acceptance criteria

```text
[ ] non-strict backward source returns zero + non-OK status
[ ] cache policy is explicit
[ ] checkpoint abandons clamped result
[ ] harness rejects clamped result
```

---

## CPM120-P2-04 — Statistics contract exceeds implementation domain

### Root cause

The API is described as general `Double()` statistics, while algorithms and CV semantics assume positive timing data.

### Recommended release direction

Declare the supported domain as:

```text
finite, nonnegative timing observations
```

Then reject out-of-domain data.

This is sufficient and aligned with product purpose.

### Acceptance criteria

```text
[ ] negative sample behavior defined
[ ] nonfinite behavior defined
[ ] zero-mean CV defined
[ ] negative threshold rejected
[ ] README says timing vectors, not arbitrary numerical vectors
```

---

## CPM120-P2-05 — Macro target resolution is not deterministic

### Root cause

```vb
Application.Run ProcedureName
```

uses Excel’s macro resolution rules.

### Impact

A same-named procedure in another workbook or add-in may be selected.

### Recommended fix

Require or construct a fully qualified macro name.

### Acceptance criteria

```text
[ ] two workbooks can expose same procedure without ambiguity
[ ] target workbook is explicit
[ ] trusted-input warning documented
[ ] qualified names covered by regression
```

---

## CPM120-P2-06 — No published Office 32-bit execution evidence

### Current evidence

Source contains VBA7/Win64 conditional declarations.

Release execution certificate is 64-bit only.

### Recommended fix

Publish a periodic 32-bit certification artifact:

```text
Excel version/build
Office bitness
commit SHA
cases/assertions/failures
```

This need not run on every commit if infrastructure is difficult, but it should run before major releases.

---

## CPM120-P2-07 — Manual, unprotected release verification

### Current strengths

- PR;
- verified merge;
- tag;
- digest;
- detailed manual Excel certificate.

### Missing controls

- no workflow;
- no required checks;
- no protected main;
- no machine-readable CI artifact.

### Recommended fix

Implement hosted static checks first, then a self-hosted Excel runner.

---

## CPM120-P2-08 — Partial source-to-artifact provenance

### Current strengths

- exact tag;
- source hashes available;
- workbook digest published;
- environment published.

### Missing link

No build manifest proves the `.xlsm` contains exactly the tagged source.

### Recommended fix

Release manifest:

```text
tag SHA
class blob/SHA-256
module blob/SHA-256
demo/test blob/SHA-256
workbook SHA-256
Excel build
bitness
test counters
build script version
```

---

# 23. Prioritized remediation plan

## Gate 1 — Status correctness

**Priority:** v1.2.x maintenance or early v1.3  
**Expected score effect:** 8.9 → 9.0

- strict endpoint status propagation;
- validation-status enum;
- checkpoint/harness rejection of invalid clamped results.

## Gate 2 — Harness integrity

**Priority:** v1.3  
**Expected score effect:** 9.0 → 9.1

- fail on non-OK worker sample;
- qualify macro target;
- cap iterations;
- add dispatch-baseline helper.

## Gate 3 — Automated assurance

**Priority:** v1.3  
**Expected score effect:** 9.1 → 9.25

- static workflow;
- headless Excel suite;
- machine-readable artifacts;
- main protection;
- required checks.

## Gate 4 — Platform and provenance

**Priority:** v1.3 release  
**Expected score effect:** 9.25 → 9.4

- Office 32-bit certification;
- build manifest;
- reproducible workbook assembly;
- immutable/signature policy.

## Gate 5 — Statistics and modularity

**Priority:** v1.3 / v2  
**Expected score effect:** 9.4 → 9.5+

- timing-only sample contract or stable generic algorithms;
- extract statistics;
- extract measurement;
- extract reporting;
- conditional test hooks.

---

# 24. Concrete route from 8.9 to 9+

| Stage | Approx. score | Deliverable |
|---|---:|---|
| Published v1.2.0 | **8.9** | Strong source, fault-injected suite, manual 64-bit release certificate |
| Status + harness hardening | **9.0–9.1** | No invalid zero can masquerade as a valid sample |
| Automated Excel regression | **9.2** | Exact-commit compile/test evidence |
| Protected merge + provenance | **9.3** | Source-bound release artifact |
| 32/64-bit recurring certification | **9.4** | Platform claim supported by execution |
| Statistics/domain + modularity | **9.5+** | Premium maintainability and numerical contract |

---

# 25. Published-release assessment

## Suitable uses

The published release is suitable for:

- production-oriented Excel VBA timing;
- QPC-based benchmarking;
- ordinary workbook automation;
- structured workflow instrumentation;
- checkpoint reporting;
- strict-mode correctness-sensitive timing;
- non-strict timing when callers inspect `LastReadStatus`;
- shared TW suppression under the documented stable-host invariant;
- repeated measurement of trusted public procedures;
- statistics on ordinary nonnegative timing vectors;
- teaching Windows timing behavior in VBA.

## Important usage rules

1. Prefer method 5, QPC.
2. Prefer strict mode when failed measurements must never be accepted.
3. In non-strict mode, check `LastReadStatus`.
4. Keep the workbook set stable while Calculation suppression is active.
5. Call `ResetEnvironment` explicitly.
6. Do not use hard `End`.
7. Use fully qualified macro names where multiple projects are open.
8. Do not treat `MeasureOverhead_Samples` as a matched `Application.Run` baseline.
9. Treat the statistics layer as timing-oriented.
10. Verify the release workbook digest before distribution.

## Release classification

> **Production-capable public release with explicit platform and lifecycle constraints.**

The release is not fully reproducible or CI-certified, but its source design, fault-injected tests, tagged PR, digest and manual Excel certificate establish a strong assurance baseline for a VBA project.

---

# 26. Final verdict

`VBA-PERFORMANCE_MANAGER` v1.2.0 is a high-quality VBA runtime component.

Its most important achievements are not feature count but **failure semantics and state ownership**:

- start sessions are transactional;
- failed native reads are not timestamps;
- method-4 failures cannot become rollover-inflated durations;
- non-strict failure is observable;
- last known good caches survive failed reads;
- failed checkpoints are abandoned;
- shared Excel state is centrally coordinated;
- Calculation baseline validity is explicit;
- heap-address identity is avoided;
- difficult native failure paths are testable;
- release behavior is documented and certified.

The principal remaining work is infrastructure and contract refinement:

- status propagation in strict endpoint failures;
- non-strict harness sample integrity;
- validation status beyond native reads;
- deterministic macro targeting;
- a tighter statistics domain;
- 32-bit evidence;
- automated Excel CI;
- reproducible source-to-workbook provenance;
- internal modularization.

### Final scores

```text
Overall repository:             8.9 / 10
Production runtime:             9.0 / 10
Timing/session core:            9.3 / 10
TW state management:            8.9 / 10
Checkpoint/reporting:           9.3 / 10
Measurement harness:            8.4 / 10
Statistics:                     8.2 / 10
Regression assurance:           9.3 / 10
Documentation:                  9.0 / 10
CI/release engineering:         7.6 / 10
```

> **Classification: advanced, production-capable pure-VBA timing and execution-control release with excellent source-level correctness and regression coverage, but incomplete automated release assurance.**

> **9+ recommendation: make invalid elapsed outcomes status-bearing end-to-end, prevent failed harness samples, add Excel CI and 32-bit certification, and publish a complete source/artifact provenance manifest.**

---

# Appendix A — Public API inventory

## Core timing

```text
StartTimer
ElapsedSeconds
ElapsedTime
```

## Session and inspection

```text
T1
T2
ET
StrictMode
ActiveMethodID
HasActiveSession
MethodName
RunLabel
LastReadStatus
```

## Diagnostics

```text
OverheadMeasurement_Seconds
OverheadMeasurement_Text
Get_SystemTickInterval
QPC_Get_SystemTickInterval
QPC_FrequencyPerSecond
QPC_FrequencyPerSecond_Value
```

## Execution and TW

```text
Pause
ResetEnvironment
TW_Turn_OFF
TW_Turn_ON
TW_IsActive
TW_ActiveSessionCount
TW_CalculationExempted
```

## Checkpoints and reporting

```text
SetRunLabel
ClearCheckpoints
Checkpoint
CheckpointCount
ReportAsArray
ReportAsText
```

## Measurement and statistics

```text
MeasureProcedure
MeasureOverhead_Samples
Stats_Count
Stats_Min
Stats_Max
Stats_Mean
Stats_Median
Stats_Percentile
Stats_StdDev
Stats_CoefficientOfVariation
Stats_IsContaminated
Stats_Text
```

## Public enums

```text
TW_Enum
cPM_ReadStatus
```

The README reports:

```text
43 public members
24 methods
19 properties
```

---

# Appendix B — Recommended new regression cases

```text
01 Strict_QPC_EndFailure_SetsLastReadStatusBeforeRaise
02 Strict_SystemTime_EndFailure_SetsLastReadStatusBeforeRaise
03 Strict_SystemTime_FormatFailure_SetsLastReadStatusBeforeRaise
04 NonStrict_NegativeElapsed_SetsInvalidStatus
05 Checkpoint_NegativeElapsed_AbandonsCapture
06 MeasureProcedure_NonStrictFailedSample_Raises
07 MeasureOverhead_NonStrictFailedSample_Raises
08 MeasureProcedure_QualifiedWorkbookTarget
09 MeasureProcedure_DuplicateMacroName_IsDeterministic
10 MeasureProcedure_IterationLimit
11 Stats_RejectsNegativeTimingSample
12 Stats_RejectsNegativeCvThreshold
13 Stats_ZeroMean_NonzeroSpread_Defined
14 Stats_LargeFiniteMean
15 Stats_LargeFiniteMedian
16 ElapsedTime_LargeDuration
17 Method3_AlignedStartFailure_ReleasesNewTimerPeriod
18 TW_ApplicationSetterFailure_BeginRollback
19 TW_ApplicationSetterFailure_EndRollback
20 Workbookless_Begin_StrictAndNonStrict
21 WorkbookCloseBeforeRestore_StableHostViolation
22 ReleaseManifest_SourceHashesMatchTag
23 PublicApiInventory_MatchesREADME
24 TestCount_MatchesREADMEAndChangelog
```

---

# Appendix C — Suggested GitHub issues

1. `Set LastReadStatus before strict endpoint errors are re-raised`
2. `Do not admit failed non-strict measurements into sample vectors`
3. `Represent negative-elapsed clamping as a non-OK operation status`
4. `Require or construct fully qualified MeasureProcedure targets`
5. `Define Stats_* as timing-only or harden generic numeric algorithms`
6. `Reject negative CV thresholds`
7. `Add a dispatch-matched empty-procedure baseline helper`
8. `Cap measurement iterations and report allocation errors explicitly`
9. `Correct QPC T1/T2 unit documentation`
10. `Correct T2/ET checkpoint documentation`
11. `Roll back timeBeginPeriod acquisition when aligned start fails`
12. `Conditionalize Friend failure-injection seams`
13. `Move cPM_Report_WriteToRange into M_cPM_REPORTING`
14. `Extract M_cPM_STATISTICS behind compatibility wrappers`
15. `Add a headless Excel regression workflow`
16. `Protect main with required static and Excel checks`
17. `Publish Office 32-bit certification`
18. `Publish a complete release provenance manifest`
19. `Adopt an immutable or signed release-asset policy`
20. `Add SECURITY.md and CONTRIBUTING.md`

---

# Appendix D — Evidence confidence

| Conclusion | Confidence |
|---|---|
| Exact tag and commit | High |
| PR and merge topology | High |
| Tagged source hashes and sizes | High |
| Release asset size and SHA-256 | High |
| 63 registered cases | High |
| 431 assertions / 0 failures | Publisher-certified, not independently rerun |
| Excel version/build/64-bit certificate | Publisher-certified |
| Transactional start behavior | High from source |
| QPC/method-4 failure handling | High from source |
| Wrong-format regression coverage | High from source |
| Calculation baseline validity | High from source |
| Strict `LastReadStatus` gap | High from control flow |
| Non-strict harness zero-sample gap | High from control flow |
| Negative-clamp status gap | High from control flow |
| 32-bit runtime behavior | Not independently executed |
| Workbook embedded-source parity | Not independently inspected |
| Reproducible workbook build | Not demonstrated |
| Timing performance values | Not independently measured |

---

# Appendix E — 9+ acceptance checklist

## Correctness and status

```text
[ ] strict endpoint failure records LastReadStatus
[ ] non-strict negative elapsed records non-OK status
[ ] checkpoint abandons every invalid/clamped measurement
[ ] harness rejects every failed endpoint sample
[ ] method-3 failed start rolls back newly acquired timer resolution
```

## Measurement and statistics

```text
[ ] procedure targets are deterministic
[ ] matched dispatch baseline exists
[ ] iteration cap exists
[ ] timing-sample domain is explicit
[ ] CV zero/negative-mean semantics are defined
[ ] negative threshold is rejected
[ ] large-value behavior is tested
```

## TW lifecycle

```text
[ ] stable-host invariant remains explicit
[ ] workbook-less begin is certified
[ ] close-before-restore behavior is certified
[ ] setter failure rollback is injected
[ ] pending restore or lifecycle tracking is evaluated
```

## Regression

```text
[ ] core runner can execute headlessly
[ ] machine-readable counters are emitted
[ ] failure details are uploaded
[ ] Office 64-bit runs automatically
[ ] Office 32-bit evidence is published
[ ] registered case count is generated
```

## Release engineering

```text
[ ] hosted static workflow
[ ] self-hosted Excel workflow
[ ] main branch protected
[ ] required status checks
[ ] release built from exact tag
[ ] source/artifact manifest published
[ ] release immutability/signing policy
```

## Architecture and maintainability

```text
[ ] native timing internals extracted
[ ] statistics internals extracted
[ ] measurement internals extracted
[ ] worksheet reporting extracted
[ ] test seams conditionalized
[ ] public API remains backward-compatible
[ ] API inventory generated
```

## Documentation and governance

```text
[ ] T2/ET checkpoint wording corrected
[ ] QPC raw-unit wording corrected
[ ] “no external DLL” changed to “no third-party DLL”
[ ] capability table includes stable-host footnote
[ ] statistics domain accurately described
[ ] SECURITY.md added
[ ] CONTRIBUTING.md added
```
