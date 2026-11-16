# #47 Demo workbook inventory

This is the first step of [#47](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/47):
an inventory of the workbook surface that v1.4.0 distributed. Each item is
marked to be retained, corrected, replaced or removed in the rebuild. It
records findings only; nothing here changes the workbook or the source.

## How it was taken

- **Input:** the published v1.4.0 asset `PERFORMANCE.MANAGER.xlsm`, SHA-256
  `5eddd6473e43a5c9fb12a97f3d00c07117843368165082c947c953003ed7b4a2`, the
  digest the v1.4.0 release notes gave at the time. The asset was re-issued
  on 2026-11-16 without its document metadata (#191); the sheets, names,
  controls and VBA project inventoried here are unchanged.
- **Method:** the file was read as an Office Open XML package (workbook,
  sheets, shared strings, drawings and names), without opening it in Excel and
  without running any macro. The demo source was read at `release/v1.4.1`
  (`8b8df34`).
- **Not covered:** behaviour from a clean open, button clicks and
  compilation. Those need Excel and belong to the smoke matrix of #47.

## Package

| Part | Content | Decision |
|---|---|---|
| Sheets | `Home`, `DEMO_cPM`, `DATA_cPM`, `HELP_cPM`, `REGRESSION_cPM`, all visible | See below |
| Tables | One, on `DEMO_cPM` (the results log) | Retain |
| Drawings | One, on `DEMO_cPM`: nine buttons | See [Controls](#controls) |
| Images | One, on `Home` (2.2 MB PNG) | Replace (see [Home](#home)) |
| VBA project | `xl/vbaProject.bin`, 2.8 MB | Rebuild from the exact candidate source |
| Document properties | Creator, dates and category | Reset in the rebuild; the maintainer reviews them before publication |

## Workbook names

| Name | Refers to | Decision |
|---|---|---|
| `BoolList` | `DEMO_cPM!$AA$2:$AA$3` | Retain |
| `cPM_MethodID`, `cPM_AlignToNextTick`, `cPM_StrictMode`, `cPM_PauseSeconds`, `cPM_PauseMethod`, `cPM_Iterations`, `cPM_RangeRows`, `cPM_TWMode` | `DEMO_cPM!$D$5` … `$D$12` | Retain |
| `cPM_ValueFill`, `cPM_ValueFillValue`, `cPM_FormulaFill`, `cPM_FormulaFillExpr` | `DATA_cPM!$D$5`, `$D$6`, `$G$5`, `$G$6` | Retain |
| `_xleta.LOG` | `#NAME?` | **Remove.** A broken name that no source creates. #47 requires no broken references. |

## Controls

`cPM_BuildDemoSheet` creates all nine: eight through `DEMO_Add_ButtonGrid`
and the regression button through `DEMO_Add_DemoButton`. Every assigned macro
exists in the source.

| Shape | Caption | Macro | Decision |
|---|---|---|---|
| `btn_cPM_Basic` | Run Basic Timing | `cPM_RunBasicTiming` | Retain |
| `btn_cPM_AllMethods` | Run All Methods | `cPM_RunAllMethods` | Retain |
| `btn_cPM_Aligned` | Run Aligned Demo | `cPM_RunAlignedDemo` | Retain |
| `btn_cPM_Diagnostics` | Run Diagnostics | `cPM_RunDiagnostics` | Retain |
| `btn_cPM_Overhead` | Run Overhead Demo | `cPM_RunOverheadDemo` | Correct (see finding 4) |
| `btn_cPM_Pause` | Run Pause Demo | `cPM_RunPauseDemo` | Retain |
| `btn_cPM_TW` | Run TW Comparison | `cPM_RunTWComparison` | Retain |
| `btn_cPM_ClearLog` | Clear Results Log | `cPM_ClearResultsLog` | Retain |
| `Btn_cPM_Regression` | Run Regression Tests | `Run_cPerformanceManager_RegressionSuite` | Replace (see finding 3) |

## Sheets

### Home

Holds only the image `images/cPM_Home.png`; it has no cell text, and no source
builds the sheet. The image is an illustration, not a screenshot:

- it lists "GetSystemTimeAsFileTime" among the timing methods, which the
  component does not use (method 4 is `timeGetSystemTime`);
- its dashboard figures are invented.

**Decision:** replace. The landing page becomes a built sheet with a
recommended first action (#47, "landing page and user journey").

### DEMO_cPM

| Area | Content | Decision |
|---|---|---|
| Control panel | Method ID 5, AlignToNextTick 0, StrictMode 1, Pause Seconds 0.25, Pause Method 1, Iterations 100, Range Rows 100000, TW Mode "All Off" (list: None, All Off, Keep ScreenUpdating, Keep EnableEvents, Keep ScreenUpdating+EnableEvents) | Retain, with explanations beside the advanced options |
| Demo actions | The buttons above | See [Controls](#controls) |
| Results log | Run ID, Timestamp, Scenario, Method ID, Method Name, Elapsed Seconds, Elapsed Time, T1, T2, ET, Notes | Retain; add read-status and failure columns where an action can report them |

### DATA_cPM

Value-fill target (`I6`, value 1), formula-fill target (`J6`, `=RAND()`) and a
benchmark target area. **Decision:** retain.

### HELP_cPM

The text is built by `cPM_BuildHelpSheet` and matches the source.

| Cell | Text | Decision |
|---|---|---|
| D5–D7 | What the three demo sheets contain | Correct: it omits `Home` and `REGRESSION_cPM` |
| D10–D13 | Recommended first steps | Retain; revise with the new landing page |
| D16–D21 | Method map 1–6 | Retain |
| D24, D25, D27 | Pause methods 1, 2 and 4 | Retain |
| D26 | "3 = Application.Wait: whole-second granularity; simple and deterministic; may overshoot slightly." | **Correct.** Since v1.2.0 the coarse loop is bounded by the same elapsed measurement as the trailing guard, so it does not overshoot by up to a second. Below about two seconds it issues no coarse wait and behaves like method 2 (CHANGELOG, v1.2.0). |
| D30–D31 | What TW suppression is | Retain |
| D32 | "…requires the companion module M_cPM_TimeWasters.bas" | Correct: the file is `M_cPM_TIMEWASTERS.bas` |
| D35–D37 | Method 5, `ElapsedSeconds`/`ElapsedTime`, `ResetEnvironment` | Retain |
| — | No text on `MeasureProcedure`, `MeasureBaseline`, `MeasureOverhead_Samples`, the evidence outputs and their `-1` sentinel, retained vectors, checkpoints, strict versus non-strict behaviour, or `LastReadStatus` | Add, with progressive disclosure (#47, "current-contract coverage") |

### REGRESSION_cPM

Holds the log of the v1.4.0 certification run: 80 case rows, all PASS, with
their timestamps. **Decision:** ship the rebuilt workbook with an empty log,
or label the rows with the exact run and SHA. Rows from an earlier run in a
newer workbook would read as evidence for it.

## Findings

1. **No source builds the `Home` sheet.** A rebuild from source does not
   reproduce the distributed workbook, which #47's "deterministic rebuild"
   criterion requires.
2. **The landing image is inaccurate** (see [Home](#home)).
3. **A test-only control sits among the product actions.**
   `Btn_cPM_Regression` runs the full regression suite from the demo panel;
   #47 asks for no test-only controls presented as ordinary features. Move it
   to `REGRESSION_cPM`, or label the group as maintainer tooling.
4. **A demo failure ends in a VBA error dialog.** `cPM_RunOverheadDemo`, like
   the other seven actions of `M_cPM_DEMO`, re-raises any error after cleanup. Since v1.4.0
   `OverheadMeasurement_Seconds` can raise `ERR_CPM_MEASURE_NO_VALID_SAMPLES`,
   so a host whose backend fails shows a runtime-error dialog with a Debug
   button, not a message on the sheet. #47 asks for failures visible without
   opening the VBE.
5. **The demo does not show the measurement harness.** No action calls
   `MeasureProcedure`, `MeasureBaseline`, `MeasureOverhead_Samples`, the
   statistics or checkpoints. `M_cPM_USAGE_EXAMPLES` covers them (#77), but no
   button reaches it.
6. **Two help statements are out of date** (D26 and D32 above).
7. **One broken workbook name** (`_xleta.LOG`).
8. **The regression log ships with rows from an earlier run.**

## What is still needed for #47

- From a clean open in 64-bit Excel: every control and assigned macro, startup
  behaviour, and Application state after success and after a handled failure.
- The redesign itself: landing page, user journey, help text, and the
  findings above.
- The rebuild from the exact candidate source, compilation, the smoke matrix,
  and the workbook hash tied to the candidate SHA.
