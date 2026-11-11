# Local Excel tools

These launchers live in `tools/` and require Windows, Git and desktop Excel.
Use a clean checkout of the reviewed candidate. Save your work and close all
Excel windows first. No administrator rights or Python installation is needed.
The tools are supervised: an Excel compile/password/repair dialog can require
attention. They are not an unattended runner with a hard timeout (#11).

| Double-click launcher | Result |
| --- | --- |
| `Certify-PerformanceManager.cmd` | Existing full core regression, compile and JSON evidence. |
| `Certify-MeasurementExamples.cmd` | Core regression plus clean import/compile of the usage module and execution of all four measurement examples; JSON includes their output. |
| `Sync-PerformanceManager-VBA.cmd` | Select an `.xlsm`; create a new copy outside the checkout, replace the six managed modules, compile, save and reopen to verify code readback. |
| `Export-PerformanceManager-VBA.cmd` | Select an `.xlsm`; read all VBA components into a fresh directory with a SHA-256 manifest, without modifying the workbook. |

Each launcher opens its output folder. Send the JSON/manifest for review.
A failure keeps diagnostic output; do not distribute an incomplete sync copy.

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
