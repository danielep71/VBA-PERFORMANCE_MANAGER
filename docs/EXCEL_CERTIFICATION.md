# Excel certification with a double click

Use Windows with desktop Excel 64-bit and Git installed. No Python, command
parameters or administrator session is needed.

1. In your Git client, check out the candidate to certify (`main` or
   `release/v1.4.1`) and pull the latest changes.
2. Save your work and close every Excel window.
3. Open the repository's `tools` folder and double-click
   `Certify-PerformanceManager.cmd`.
4. Let Excel import, compile and run the suite. The console remains open and
   reports `RESULT=PASS` or `RESULT=FAIL`.
5. The evidence folder opens automatically. Send its
   `cpm-regression-*.json` file for review. Keep `session.log` as well.

The launcher certifies the commit actually checked out, not a hard-coded SHA.
The working tree must be clean, including untracked files. Download ZIP is not
a Git checkout and is not supported by this launcher.

The existing runner temporarily enables VBA project access for the current
user and restores the previous value afterwards. It does not alter macro policy;
an organizational policy can still prevent execution. Do not run as administrator.

This is a supervised local runner (#11), not the automated CI gate tracked in #156.
If a compile/runtime dialog appears, capture the error before dismissing it;
the underlying runner has no timeout for modal dialogs. A forced Excel cleanup
is a failed certification. No new Excel execution is claimed by adding this
launcher. It runs the source suite only, not the separate red check for #31
described below. To certify the measurement examples, and for the workbook
sync and export launchers, see [Local Excel tools](LOCAL_EXCEL_TOOLS.md).

## Timer rollback (#31)

From a clean checkout, double-click `tools/Certify-TimerRollback.cmd`. It runs
the full suite, then repeats it using a temporary copy of the class with only
the rollback call disabled. The checkout is never edited. The source
run must pass; only case 86 must fail in the negative control, and both Excel
processes must exit cleanly. Send the generated JSON, which contains both runs.
The negative control is expected to report assertion failures; it is not source
certification.
