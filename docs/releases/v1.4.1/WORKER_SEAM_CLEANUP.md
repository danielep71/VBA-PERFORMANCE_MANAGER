# Worker seam cleanup — issue #75

Status at preparation: real 64-bit Excel execution pending.

`MeasureProcedure` now enables its existing error/cleanup handler before
blank-name validation. The handler captures and re-raises the original number,
source and description, retains unpublished output sentinels and disarms both
worker-failure counters even when no worker was created.

`MeasureBaseline` has its own blank-name validation before delegation; it clears
both counters on that branch before raising its existing baseline-specific error.
Valid calls still delegate normally. Do not clear seams unconditionally at entry:
that would disable deliberately armed failures for valid measurements.

Audit of early paths:

| Entry point | Before worker creation | Cleanup |
|---|---|---|
| MeasureProcedure | Blank-name validation, method normalization, qualification, allocation | Existing handler now covers validation too |
| MeasureBaseline | Own blank-name validation; otherwise delegates | Explicit disarm for its own raise; delegated paths use MeasureProcedure |
| MeasureOverhead_Samples | Sentinel assignment and bounded count normalization before handler; method normalization/allocation after handler | Existing handler already covers the error-producing operations |

Errors converting caller arguments before a method is entered are outside this
method-exit guarantee. No public signature, error number, output sentinel or
ordinary measurement semantics changes. Existing tests still check deliberate
start/endpoint failures on valid calls, preventing a vacuous fix that disables
all fault injection.

New final case 87 runs seven scenarios: empty/space-only procedure name;
empty/space-only baseline name; strict invalid method through procedure,
baseline and overhead. Each arms both worker seams, checks the exact original
error and three sentinel fields, then checks full sample count, zero failed
reads, zero rejected samples and OK status on a valid same-instance QPC run.

The case adds 56 assertions, giving an expected 87 cases / 771 assertions from
the prior 86 / 715 baseline. These are expected counts, not execution evidence.
The new case is appended so existing case 86 timer-rollback red-check wiring is
unchanged. Additional work is 21 short workload cycles with no intentional pause.

## Execution evidence required

Run `tools/Certify-PerformanceManager.cmd` from a clean exact-candidate checkout
on the available 64-bit Excel host. Retain JSON and session log. Verify import
hashes, successful compile, 87 cases, actual assertion count, zero failures and
completed cleanup before closing #75 or integrating executable changes into main.
No 32-bit host is required. Existing runner trust and recovery requirements apply.

For an optional negative control, the existing runner's `-RedFile` option can
substitute the parent commit's class with `-RedComponent
src\classes\cPerformanceManager.cls -RedExpectedCase 87`; keep the new test module.
The four blank-name scenarios should expose leaked counters in case 87, while
other existing cases stay green. Keep red-run failure and cleanup evidence
separate from the source run; do not publish a mutated class as the candidate.
