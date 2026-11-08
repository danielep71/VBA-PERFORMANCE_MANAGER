# Measurement-example evidence (#77)

## Source review

All six harness calls in the four measurement examples now request
`FailedReadsOut`, `LastFailureStatusOut` and `RejectedSamplesOut` explicitly.
No production module or public class contract changes.

| Example | Evidence and reporting policy |
| --- | --- |
| `Example_Benchmark_WithStatistics` | Print failed/rejected/status; withhold statistics when either count is nonzero; retain the variance check after the evidence check. |
| `Example_Compare_TwoImplementations` | Keep separate evidence for both vectors; print both; withhold summaries and speedup if either run is incomplete. |
| `Example_DispatchMatchedBaseline` | Keep separate baseline/workload evidence; print both; withhold medians, subtraction and interpretation if either run is incomplete. |
| `Example_ReadStatus_Diagnostics` | Print failed reads and rejected samples; report a median only when both are zero. |

`FailedReadsOut` includes warm-up native-read failures. `RejectedSamplesOut`
counts measured cycles excluded after start fallback. Nonzero includes negative
unpublished-evidence sentinels; no such value qualifies as a clean run. These
examples deliberately require a complete run, although retained samples from
an incomplete run can be useful in a separately qualified analysis.

The diagnostic's direct-instance measurement remains distinct from the isolated
worker's harness evidence. Procedure headers describe the amended checks.

## Documentation search

README and the wiki at `78de82ca6200680d661c5e7038e3a092fccbbda6` contain no copies
of these four example procedures or the affected workload-call snippets.
They are unaffected. The changelog records the correction.

## Exact-candidate Excel verification

Pending at implementation time. Record the full candidate SHA and Office build
with the result; a later PR evidence comment supersedes this preparation record.

The ordinary `Certify-PerformanceManager.cmd` runner imports the four regression
dependencies, **not** `M_cPM_USAGE_EXAMPLES`. Its JSON alone cannot certify #77.
Run the ordinary regression and also perform this focused check:

1. From a clean checkout of the candidate, create a disposable `.xlsm` workbook.
   Import `src/modules/M_cPM_TIMEWASTERS.bas`,
   `src/classes/cPerformanceManager.cls`, and
   `demo/M_cPM_USAGE_EXAMPLES.bas` into its VBA project. Keep only this test
   workbook open to avoid unqualified workload-name ambiguity.
2. Create a worksheet named `DATA_cPM`. The examples write synthetic data into
   columns H and I; do not use a workbook containing useful data.
3. Run **Debug > Compile VBAProject**, then open the Immediate Window (Ctrl+G).
4. Run `Run_MeasurementUsageExamples`. Retain its entire Immediate Window output,
   including all four example banners, failed/rejected counts and conclusions.
   The example handlers print unexpected errors rather than re-raising them:
   merely returning from the macro is not a passing result. Any `Error ...`
   line requires investigation. A noisy-run refusal is legitimate, not a test
   failure. Ordinary successful measurements should show zero failed/rejected
   counts before statistics or a noise refusal.
5. Record successful compile, all four completed examples, Excel build/bitness,
   exact SHA and any observed errors. Close the disposable workbook normally.

Review the nonzero-evidence guards directly: they precede every statistical
conclusion and combine both evidence counters for every measured vector.
The existing core suite supplies deterministic backend-failure/fallback tests;
this manual smoke does not claim to force a hardware read failure in an example.

No 32-bit execution claim is made. Final #47 workbook certification remains a
separate smoke matrix.
