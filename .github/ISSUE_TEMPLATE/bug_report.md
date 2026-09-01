---
name: "🐛 Bug report"
about: "Report incorrect behavior or an unreliable measurement"
title: "[Bug] "
labels: "bug"
assignees: ""
---

<!--
Thank you for helping improve VBA Performance Manager.

Before submitting:
- Search open and closed issues for an existing report.
- Reduce the problem to the smallest reproducible workbook or code sample.
- Record the exact release tag or full 40-character commit SHA you tested.
- Remove passwords, client data, workbook internals, and other sensitive content.

Security vulnerabilities must not be reported publicly. Follow SECURITY.md instead.

Keep the headings below. Replace the prompts and examples with your evidence.
Screenshots are useful context, but they do not replace copyable text.
-->

## 🔍 Summary

<!--
Describe one observable problem in a few sentences.
State what fails, when it fails, and why it matters. Avoid diagnosing the cause
unless you have evidence for it.
-->


## 🎯 Expected and actual behavior

**Expected**

<!-- What should have happened? Link the documented contract when possible. -->


**Actual**

<!-- What happened instead? Include exact values, status text, and error details. -->


## 🔁 Minimal reproduction

**Frequency**

- [ ] Every time
- [ ] Intermittent — approximately <!-- e.g. 3 of 10 runs -->
- [ ] Happened once

**Steps**

1. <!-- First step from a clean Excel session or known workbook state. -->
2. <!-- Next step. -->
3. <!-- The action that exposes the problem. -->

**Minimal VBA**

<!--
Prefer a new workbook with only the two runtime source files:
- src/modules/M_cPM_TIMEWASTERS.bas
- src/classes/cPerformanceManager.cls

Capture LastReadStatus immediately after the operation being diagnosed; a later
call can replace it. Always restore Excel state with ResetEnvironment, including
on an error path.
-->

```vb
Option Explicit

Public Sub ReproduceIssue()
    Dim PM As cPerformanceManager
    Dim ElapsedS As Double

    On Error GoTo CleanFail
    Set PM = New cPerformanceManager

    PM.StartTimer cPM_MethodQPC

    ' Minimal workload that exposes the problem.

    ElapsedS = PM.ElapsedSeconds
    Debug.Print "ElapsedSeconds="; ElapsedS
    Debug.Print "ActiveMethodID="; PM.ActiveMethodID
    Debug.Print "LastReadStatus="; PM.LastReadStatus

CleanExit:
    If Not PM Is Nothing Then PM.ResetEnvironment
    Exit Sub

CleanFail:
    Debug.Print "Err.Number="; Err.Number
    Debug.Print "Err.Source="; Err.Source
    Debug.Print "Err.Description="; Err.Description
    Resume CleanExit
End Sub
```

**Cleanup result**

<!--
After the reproduction, did ResetEnvironment complete? Did Excel settings return
to their original values? State whether VBA's End command or an Excel restart
was required; both can hide cleanup defects and shared-state leaks.
-->


## 📦 Exact source identity

<!--
"Latest" is not reproducible. Use a release tag or the full output of
`git rev-parse HEAD`. If the source was edited, link or attach a minimal diff.
-->

| Field | Value |
| --- | --- |
| Release tag or full 40-character commit SHA | <!-- e.g. v1.4.0 or 0123... --> |
| Source origin | <!-- Official release / release branch / fork / copied files --> |
| Local modifications | <!-- None, or describe/link the diff --> |
| `VERSION` in the `cPerformanceManager.cls` header | <!-- Exact value --> |
| `VERSION` in the `M_cPM_TIMEWASTERS.bas` header | <!-- Exact value --> |
| Imported files | <!-- Exact filenames --> |
| `Debug > Compile VBAProject` | <!-- Pass / fail, with exact error --> |


## 💻 Host environment

| Field | Value |
| --- | --- |
| Excel version and full build | <!-- File > Account > About Excel --> |
| Office bitness | <!-- 32-bit / 64-bit --> |
| Windows version and build | <!-- Exact version/build --> |
| Workbook host | <!-- .xlsm / .xlam / PERSONAL.XLSB / other --> |
| Relevant add-ins | <!-- None, or list names and versions --> |
| Regional settings | <!-- If date, number, or string formatting may matter --> |

<!--
The source contains VBA7 declarations for both 32-bit and 64-bit Office, but
published execution evidence covers 64-bit Excel only. A 32-bit report is
valuable; please identify it explicitly.
-->


## ⚠️ Error, timer, and result evidence

<!--
Use "N/A" for fields that truly do not apply. Preserve the raw values.
Capture LastReadStatus immediately after the failing timer read.
-->

| Field | Value |
| --- | --- |
| `Err.Number` | <!-- Decimal and/or hexadecimal --> |
| `Err.Source` | |
| `Err.Description` | |
| Returned value or output | |
| Requested timing method | <!-- Enum name preferred; IDs are Timer=1, TickCount=2, timeGetTime=3, system time=4, QPC=5, Now=6 --> |
| `ActiveMethodID` | |
| `HasActiveSession` | |
| `StrictMode` | <!-- True by default --> |
| `LastReadStatus` | <!-- OK / QPC failed / system-time failed / format invalid / fallback to method2 / elapsed invalid --> |

<!--
In non-strict mode, a failed high-resolution read may fall back. Report both the
requested method and ActiveMethodID; they answer different questions.
-->


## 📊 Measurement-harness evidence

<!--
Complete this section for MeasureProcedure, MeasureBaseline, or overhead/sample
collection problems. Otherwise write "Not applicable".

MeasureProcedure exposes worker-read evidence through FailedReadsOut and
LastFailureStatusOut. The caller object's LastReadStatus describes the caller's
own last read, not the internal measurement worker. Current MeasureBaseline and
MeasureOverhead_Samples APIs do not expose equivalent worker outputs, so state
that limitation rather than inferring "no failures".

Application.Run executes the supplied procedure name. Use only a trusted target
and include the exact qualification used, including workbook/module names.
-->

| Field | Value |
| --- | --- |
| API used | <!-- MeasureProcedure / MeasureBaseline / MeasureOverhead_Samples --> |
| Exact `ProcedureName` | <!-- Include full qualification and apostrophes --> |
| Host workbook filename | |
| Iterations / warm-up iterations | |
| Requested timing method | |
| `Stats_Count` returned | |
| `FailedReadsOut` | <!-- MeasureProcedure, or N/A --> |
| `LastFailureStatusOut` | <!-- MeasureProcedure, or N/A --> |
| Sample vector concern | <!-- Mixed backend / invalid value / count / other --> |


## 🧹 Shared Excel-state evidence

<!--
Complete this section for TimeWasters, nesting, cleanup, pause/resume, or leaked
Excel-state problems. The TimeWasters coordinator is process-wide, so record any
other open workbooks or add-ins that may participate.
-->

| Field | Before | After reproduction / cleanup |
| --- | --- | --- |
| `TW_IsActive` | | |
| `TW_ActiveSessionCount` | | |
| `TW_CalculationExempted` | | |
| `Application.ScreenUpdating` | | |
| `Application.EnableEvents` | | |
| `Application.DisplayAlerts` | | |
| `Application.Calculation` | | |
| `Application.Cursor` | | |
| Open workbook set | | |


## 🧪 Validation evidence

<!--
Static CI checks source shape and policy; it does not launch Excel. If you ran the
optional regression host, import demo/M_DEMO_BUILDER.bas and test/M_cPM_Test.bas
and report the exact result below. Do not copy the project's published numbers
unless you ran that suite against the source identity reported above.
-->

- `Debug > Compile VBAProject`: <!-- Pass / fail -->
- Static lint: <!-- Not run / pass / fail; include command and output -->
- Regression suite: <!-- Not run / pass / fail -->
- Tested source SHA: <!-- Full SHA -->
- Cases / assertions / failures: <!-- Exact counts -->
- Excel build and bitness used for the run: <!-- Exact values -->
- Failing case names and first failure: <!-- Copyable text -->
- Targeted checks performed: <!-- Describe focused validation -->


## 📎 Logs, screenshots, and additional context

<!--
Paste short logs as text. Attach a minimal workbook only if it contains no
sensitive data and the code sample above cannot reproduce the issue. Explain any
relevant concurrency, workbook lifecycle, midnight rollover, system clock,
sleep/resume, or VM/remote-session conditions.
-->


## ✅ Reporter checklist

- [ ] I searched open and closed issues for this problem.
- [ ] I supplied an exact release tag or full commit SHA, not "latest."
- [ ] I reduced the report to a minimal, repeatable example where possible.
- [ ] I captured statuses and error details before another call could replace them.
- [ ] I described cleanup and any Excel application-state changes.
- [ ] I removed credentials, client data, and other sensitive information.
- [ ] This is not a security vulnerability; security reports follow `SECURITY.md`.
