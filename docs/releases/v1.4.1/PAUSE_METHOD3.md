# Pause method 3 coverage — issue #74

Outcome: merged into `main` in increment 12 ([#133](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/133)). Excel at
`ca66242`: 86 cases, 715 assertions, 0 failures; #74 is closed. See the
[evidence record](EVIDENCE.md). The text below is the preparation record.

Status: implementation prepared; real 64-bit Excel execution pending.
This is test coverage, not new release certification or a timing API change.

`Test_Pause_Method3` retains the existing one-second request, explicitly named
as the Timer-guard-only path, and adds a 2.5-second request with independent QPC
measurement. The source enters `Application.Wait` only with at least two seconds
remaining. On a quiet host the new request leaves 0.5 seconds of entry margin;
a severely descheduled process can consume that margin before the first wait.
This is a real-host timing regression, not deterministic branch instrumentation.

The long request must measure between 2.45 and 2.75 seconds. The lower allowance
is 50 ms for Timer quantization against QPC. The upper allowance is 250 ms for
Timer/DoEvents scheduling noise on the controlled certification host; it is
less than a whole-second API tick and rejects an additional coarse-second delay.
The trailing Timer guard must complete the requested duration even if the coarse
wait returns early. These tolerances do not promise an OS real-time deadline.
Do not certify a run interrupted by suspension, debugging or heavy host load;
retain failure evidence and investigate before rerunning, rather than loosening
the bound automatically.

One existing case now contains two assertions instead of one. `TotalSteps` and
the runner retain 86 cases. Relative to the latest supplied 714-assertion run,
the expected clean total is 715 (confirm from the actual JSON, never infer a
PASS). Nominal suite duration increases by approximately 2.5 seconds. Its
status-bar label matches the case log label. Production code is unchanged.

## Execution and evidence

Use a clean checkout of the exact PR candidate. Close Excel after saving work,
then run `tools/Certify-PerformanceManager.cmd`. The existing launcher imports
that commit, compiles and runs the full suite in desktop Excel; retain its JSON
and session log. It temporarily changes VBA-project access as already documented
by the runner and restores it; no new runner permissions are introduced here.

Before closing #74, record the full SHA, Excel build, 64-bit Office, completed
run/cleanup state, actual cases/assertions/failures and the case's elapsed values
from the regression log. A static pass does not execute Application.Wait.
If branch entry itself needs diagnosis, use a separate manual debug run with a
breakpoint on the existing Application.Wait line and the 2.5-second request;
never use a breakpoint-interrupted run as timing-certification evidence.

The one-second and 2.5-second assertions must both pass, the suite must finish
with zero failures and successful cleanup, and the evidence must identify the
same commit as the PR. No 32-bit host is required. Keep #74 open and do not
integrate this executable test change into main until that evidence is verified.
