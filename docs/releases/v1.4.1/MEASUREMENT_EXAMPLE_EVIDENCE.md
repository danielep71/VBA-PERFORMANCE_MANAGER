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

Run `tools/Certify-MeasurementExamples.cmd` from the exact clean candidate.
It runs the full regression, imports the usage module, creates the disposable
DATA_cPM sheet, compiles and executes all four examples. Send the generated JSON:
it includes imported hashes, compile status, core regression results, the example
transcript, its verdict and cleanup state. See [Local Excel tools](../../LOCAL_EXCEL_TOOLS.md).

The ordinary `Certify-PerformanceManager.cmd` does not import the usage module;
its JSON alone still cannot certify #77. The prior 0a32b32 core-only result does
not certify the new capture/automation changes.

The focused result records unexpected errors even when the interactive example
handler swallows them. Noise/incomplete-measurement refusals remain legitimate
outputs and must be reviewed as such; a PASS is execution evidence, not a claim
that every run produced a usable performance comparison.

Review the nonzero-evidence guards directly: they precede every statistical
conclusion and combine both evidence counters for every measured vector.
The existing core suite supplies deterministic backend-failure/fallback tests;
this manual smoke does not claim to force a hardware read failure in an example.

No 32-bit execution claim is made. Final #47 workbook certification remains a
separate smoke matrix.
