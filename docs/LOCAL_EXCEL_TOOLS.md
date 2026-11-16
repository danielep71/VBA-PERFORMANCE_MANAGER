# Local Excel tools

These launchers live in `tools/` and require Windows, Git and desktop Excel.
Use a clean checkout of the reviewed candidate. Save your work and close all
Excel windows first. No administrator rights or Python installation is needed.
The tools are supervised: an Excel compile/password/repair dialog can require
attention. They are not an unattended runner with a hard timeout (#156).

| Double-click launcher | Result |
| --- | --- |
| `Certify-PerformanceManager.cmd` | Existing full core regression, compile and JSON evidence. |
| `Certify-TimerRollback.cmd` | Core regression plus a negative control with the timer-period rollback disabled, in which only case 86 may fail. |
| `Certify-MeasurementExamples.cmd` | Core regression plus clean import/compile of the usage module and execution of all four measurement examples; JSON includes their output. |
| `Sync-PerformanceManager-VBA.cmd` | Select an `.xlsm`; create a new copy outside the checkout, replace the six managed modules, compile, save and reopen to verify code readback. |
| `Export-PerformanceManager-VBA.cmd` | Select an `.xlsm`; read all VBA components into a fresh directory with a SHA-256 manifest, without modifying the workbook. |
| `Build-DemoWorkbook.cmd` | Builds the demo workbook from the checked-out commit, smoke-tests every demo action, removes the build machine's metadata, and packages it with `LICENSE` and `NOTICE` in `..\cpm-demo-build` (#47, #72). |
| `Verify-SyncRefusals.cmd` | Evidence that the sync refuses a protected VBA project and a component type conflict, restores AccessVBOM and leaves no Excel process (#138). It creates both test workbooks in `..\cpm-sync-refusals`; locking the project is one manual step in the VBA editor, needed once. |

Each launcher opens its output folder. Send the JSON/manifest for review.
A failure keeps diagnostic output; do not distribute an incomplete sync copy.
[Excel certification with a double click](EXCEL_CERTIFICATION.md) describes the
two certification launchers step by step. To run several commits in one pass,
use `tools/Invoke-ExcelBatch.ps1` with a JSON manifest
(`tools/excel-batch.example.json`); see the README's
[Regression testing](../README.md#regression-testing) section.

## Building the demo workbook

`Build-DemoWorkbook.cmd` builds the release workbook from the clean checked-out
commit, without manual steps, into a new folder under `..\cpm-demo-build`:

1. imports the six managed components into a new workbook and compiles;
2. runs `cPM_CreateDemoSheets`, which builds `Home`, `DEMO_cPM`, `DATA_cPM`
   and `HELP_cPM`, and deletes every other sheet;
3. runs every demo action unattended and checks that each adds a
   results-log row and that none fails; failures are counted by the demo
   itself, so no message box blocks the run;
4. clears the results log and saves `PERFORMANCE MANAGER.xlsm`;
5. removes the metadata a managed machine can add (sensitivity label,
   document category or content status, the folder it was saved from, print
   headers and footers) and checks that none is left; these are the kinds
   `tools/release_provenance.py` refuses;
6. reopens the workbook with macros disabled and checks its sheets;
7. writes `VBA-Performance-Manager-v<VERSION>.zip` with the workbook and the
   checkout's `LICENSE` and `NOTICE`, and `demo-build.json`.

`RESULT=PASS` needs all seven steps. The build record names the commit, the
Excel build, every component hash, each action's rows and time, what metadata
was removed, and the workbook and package SHA-256. The smoke test is demo
evidence; it does not replace the regression suite.

## Verifying the sync refusals

`Verify-SyncRefusals.cmd` works in `..\cpm-sync-refusals`, next to the
checkout, and never changes the checkout.

- **`conflict.xlsm`:** created automatically, with a class module named
  `M_cPM_TIMEWASTERS`, which the sync expects as a standard module.
- **`protected.xlsm`:** created automatically with one standard module. Excel
  cannot lock a VBA project through automation, so the script opens it with
  the VBA editor and waits. In the editor: Tools > VBAProject Properties >
  Protection, lock the project with any password, save, close the workbook,
  then press Enter. Later runs reuse the locked file.

Each sync must then fail with its refusal, leave the input unchanged, report
cleanup `completed`, restore AccessVBOM to its prior value and leave no Excel
process. The sync manifest does not record AccessVBOM, so the script reads it
before and after each run. Invisible Excel instances are closed first; a
visible Excel window stops the script. The verdict and every check go to
`sync-refusals-<sha>-<UTC time>.json` in the work folder.

## Certifying the examples

`Certify-MeasurementExamples.cmd` reuses the existing certification runner and
adds `M_cPM_USAGE_EXAMPLES.bas`. It creates `DATA_cPM` in the disposable workbook,
runs `Certify_MeasurementUsageExamples`, and records the transcript under
`details.runs[0].examples`. The capture includes the four BEGIN/END pairs and all
measurement-example output. Unexpected errors counted by the example handlers
or missing/duplicate completion markers cause failure. Legitimate noise or
incomplete-measurement refusals remain visible in the transcript and do not
constitute an unexpected execution error.

A PASS establishes compile and execution of the examples, not forced hardware
failures, numerical performance guarantees or a rebuilt release workbook. The
JSON retains the ordinary regression counts separately from the example output.
The normal certification launcher still runs only the core suite.

## Synchronization boundary

Sync requires a clean Git checkout and supports `.xlsm` only. It copies the
input first and modifies only the copy. It replaces these named components:

- `M_cPM_TIMEWASTERS`, `cPerformanceManager`;
- `M_DEMO_BUILDER`, `M_cPM_DEMO`, `M_cPM_USAGE_EXAMPLES`;
- `M_cPM_Test`.

Other modules, document modules, sheets and controls are retained. Conflicting
component types and protected VBA projects are refused; passwords are never
removed. Module code is read back before saving and after reopening, ignoring
export metadata, line endings and trailing whitespace. Identifier case is
also ignored, because the VBE keeps one spelling per identifier across the
project and recases imported code to it; the manifest lists each component
that differed only in case under `case_normalized`. Any other difference fails.
Source and output hashes are recorded. This is a development workbook with
demo/test modules; it does
not replace the two-file production installation or #47's workbook rebuild.
It does not rebuild sheets, fix button assignments, refresh Ribbon XML, run
workbook startup, or validate arbitrary pre-existing code. Recompilation can
invalidate a VBA signature; re-sign only through the established release process.

## Export boundary

Export opens the workbook read-only and exports standard/class/form/document
components, including form binary companions where Excel emits them. Document
module `.cls` files are review exports, not drop-in replacement class imports.
Exports go into a unique directory outside the checkout; no source file is
overwritten and no workbook/module name is inferred as authority to update Git.
The input workbook's hash is checked again after close.

## Trust and recovery

Workbook sync/export opens Excel with macros and events disabled and does not
run workbook callbacks. The command launchers temporarily enable AccessVBOM
for the current user and restore its prior value in cleanup, like the existing
certifier. Managed policy may still refuse access; do not bypass such policy.
For an already provisioned environment the PowerShell entry point can omit
`-AllowVbomAccess`. Close normally if a modal prompt appears; retain failure
output. Do not terminate PowerShell during registry restoration.

Only the Excel process created by the operation is targeted for cleanup. Forced
termination is reported as failure. The original workbook must have an unchanged
hash for sync/export PASS. To roll back a sync, discard the new output directory
and keep using the untouched input workbook.

## Optional command-line use

```powershell
& .\tools\Certify-MeasurementExamples.cmd
& .\tools\Sync-PerformanceManager-VBA.cmd -Workbook 'C:\Work\Demo.xlsm'
& .\tools\Export-PerformanceManager-VBA.cmd -Workbook 'C:\Work\Demo.xlsm'
```

Run these tools only on workbooks you intend to inspect. Source exports can
contain private VBA strings or logic: review files before uploading them.
