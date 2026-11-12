# Strict error-code review — issue #76

Outcome: merged into `main` in increment 13 ([#136](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/136)). Excel: 87
cases, 771 assertions, 0 failures; #76 is closed. See the
[evidence record](EVIDENCE.md). The text below is the preparation record.

Status at preparation: implementation ready for exact-candidate 64-bit Excel
execution. Production code is unchanged. Tests require specific documented
errors instead of accepting unrelated nonzero errors.

## Review matrix

| Test / operation | Required error |
|---|---|
| Strict elapsed before StartTimer | ERR_CPM_NO_ACTIVE_SESSION |
| Strict elapsed backend mismatch | ERR_CPM_METHOD_MISMATCH |
| Strict invalid StartTimer method | ERR_CPM_INVALID_START_METHOD |
| TW blank-key BeginSession | ERR_TW_BEGIN_BLANK_KEY |
| TW blank-key EndSession | ERR_TW_END_BLANK_KEY |
| TW blank-key IsInstanceActive | ERR_TW_ISACTIVE_BLANK_KEY |
| Checkpoint before session | ERR_CPM_CHECKPOINT_NO_SESSION |
| SetRunLabel after capture | ERR_CPM_RUNLABEL_AFTER_CAPTURE |
| Strict worker QPC start fault | ERR_CPM_QPC_READ_FAILED |
| Legacy overhead text invalid backend | ERR_CPM_INVALID_START_METHOD |
| LastReadStatus strict QPC endpoint fault | ERR_CPM_QPC_READ_FAILED |
| LastReadStatus strict method-4 read fault | ERR_CPM_TIMEGETSYSTEMTIME_FAILED |
| LastReadStatus strict method-4 wrong format | ERR_CPM_TIMEGETSYSTEMTIME_FORMAT |
| LastReadStatus strict checkpoint QPC fault | ERR_CPM_QPC_READ_FAILED |

All 14 assertions use `Test_Assert_EqualLong`. Expected values were checked
against the class error enum, the relevant raise sites and
`Read_StatusToErrorNumber`, or the shared TW module's blank-key raises. Existing
status, description and exclusion assertions remain; a different nonzero error
no longer satisfies the strengthened assertion.

The three boolean statistics helpers (`Test_Stats_RaisesFor`,
`Test_Stats_MinRaisesFor`, `Test_Stats_MedianRaisesFor`) are used only with
negative samples in the current suite. Each now accepts only
`ERR_CPM_STATS_OUT_OF_DOMAIN`; their comments state this narrower contract.

## Deliberately retained checks

`Test_Measure_RaisesFor` is used only for the deliberately nonexistent macro
`cPM_NoSuchProcedure_XYZ`. That error originates in Excel's `Application.Run`;
the library propagates it and does not assign a specific CPM error number.
The helper intentionally checks that a failure propagates, without inventing a
library contract for Excel's external error number.

Other `Err.Number <> 0` / saved-error checks are cleanup decisions, fixture
setup guards or unexpected-failure reporting, not expected strict-raise
assertions. They remain generic so any unexpected failure is handled.
Specific error comparisons already present elsewhere remain unchanged.

## Validation boundary

No cases or assertion calls are added: expected source execution remains
**87 cases and 771 assertions**. These are expected totals until confirmed by
the new JSON. The previous #75 run does not certify these changed tests.
Use the existing `tools/Certify-PerformanceManager.cmd` from a clean checkout
of the exact candidate; retain import hashes, compile outcome, Excel build and
64-bit Office, cases/assertions/failures and completed cleanup. Keep #76 open
and the PR draft until that evidence is verified. No 32-bit host is required.
