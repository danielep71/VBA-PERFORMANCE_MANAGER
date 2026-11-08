# Excel certification with a double click

Use Windows with desktop Excel 64-bit and Git installed. No Python, command
parameters or administrator session is needed.

1. In your Git client, select `release/v1.4.1` and pull the latest changes.
2. Save your work and close every Excel window.
3. Open the repository's `tools` folder and double-click
   `Certify-PerformanceManager.cmd`.
4. Let Excel import, compile and run the suite. The console remains open and
   reports `RESULT=PASS` or `RESULT=FAIL`.
5. The evidence folder opens automatically. Send its
   `cpm-regression-*.json` file for review. Keep `session.log` as well.

The launcher certifies the commit actually checked out, not a hard-coded SHA.
Use the release branch for the pending #33 lifecycle coverage (85 cases at the
time this launcher was introduced). The main branch has a different test suite.
The working tree must be clean, including untracked files. Download ZIP is not
a Git checkout and is not supported by this launcher.

The existing runner temporarily enables VBA project access for the current
user and restores the previous value afterwards. It does not alter macro policy;
an organizational policy can still prevent execution. Do not run as administrator.

This is a supervised local runner, not the automated CI gate tracked in #11.
If a compile/runtime dialog appears, capture the error before dismissing it;
the underlying runner has no timeout for modal dialogs. A forced Excel cleanup
is a failed certification. No new Excel execution is claimed by adding this
launcher. It runs the source suite only, not the separate red check for #31.

## Timer rollback (#31)

On `fix/31-timer-period-rollback`, double-click `tools/Certify-TimerRollback.cmd`.
It runs all 86 cases, then repeats them using a temporary copy of the class
with only the rollback call disabled. The checkout is never edited. The source
run must pass; only case 86 must fail in the negative control, and both Excel
processes must exit cleanly. Send the generated JSON, which contains both runs.
The negative control is expected to report assertion failures; it is not source
certification. Do not merge #107 until this evidence is reviewed.
