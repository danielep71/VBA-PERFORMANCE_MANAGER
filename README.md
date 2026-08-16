<div align="center">

# ⚙️ Class Performance Manager

### A session-bound, benchmark-grade timing and execution-control class for pure Excel VBA

**Six timing backends · Monotonic high-resolution QPC · Session-bound validation · Distribution-aware statistics · Structured checkpoints · Shared Excel state suppression**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-32%20%2F%2064--bit-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://github.com/danielep71/vba-performance_manager)
[![Pure VBA](https://img.shields.io/badge/Implementation-Pure_VBA-00599C?style=for-the-badge)](https://github.com/danielep71/vba-performance_manager)
[![No External DLL](https://img.shields.io/badge/External_DLL-None-555555?style=for-the-badge)](#-installation)
[![QPC](https://img.shields.io/badge/Default_Backend-QueryPerformanceCounter-c2185b?style=for-the-badge)](#-timing-backends)
[![Regression](https://img.shields.io/badge/Regression-72_Cases_·_511_Assertions-d97706?style=for-the-badge)](#-testing-and-validation)
[![Statistics](https://img.shields.io/badge/Statistics-Median_·_P95_·_CV-0f766e?style=for-the-badge)](#-measurement-and-statistics)
[![Version](https://img.shields.io/badge/Version-1.3.0-4c1d95?style=for-the-badge)](CHANGELOG.md)

<br>

[![License](https://img.shields.io/github/license/danielep71/vba-performance_manager?style=flat-square&color=2ea44f)](LICENSE)
[![Stars](https://img.shields.io/github/stars/danielep71/vba-performance_manager?style=flat-square&logo=github&color=6f42c1)](https://github.com/danielep71/vba-performance_manager/stargazers)
[![Forks](https://img.shields.io/github/forks/danielep71/vba-performance_manager?style=flat-square&logo=github&color=0969da)](https://github.com/danielep71/vba-performance_manager/network/members)
[![Issues](https://img.shields.io/github/issues/danielep71/vba-performance_manager?style=flat-square&color=d73a49)](https://github.com/danielep71/vba-performance_manager/issues)
[![Last commit](https://img.shields.io/github/last-commit/danielep71/vba-performance_manager?style=flat-square&color=orange)](https://github.com/danielep71/vba-performance_manager/commits/main)

<br>

**No add-in · No installer · No COM reference · Two files to import**

[Read the Wiki](https://github.com/danielep71/vba-performance_manager/wiki)
&nbsp;·&nbsp;
[Quick start](#-quick-start)
&nbsp;·&nbsp;
[Measurement and statistics](#-measurement-and-statistics)
&nbsp;·&nbsp;
[Public API](#-public-api)
&nbsp;·&nbsp;
[Changelog](CHANGELOG.md)

</div>

---

<p align="center">
  <img src="images/performance_manager_banner2.png"
       alt="Class Performance Manager — benchmark-grade timing for Excel VBA"
       width="100%">
</p>

---

> [!IMPORTANT]
> **This is a timing *session manager*, not a `Timer()` wrapper.**
>
> A backend is bound when the session starts, and every elapsed-time read is validated against that same backend. Rollover arithmetic, unsigned 32-bit conversion, timer-resolution lifecycle, and shared Excel state coordination are all handled explicitly rather than left to the caller. Backwards-clock detection is applied to backends 2, 3, 4 and 6 — see the limitations on backend 1.

<div align="center">

### 📐 At a glance

| | |
|---:|:---|
| **6** | timing backends behind one session-bound interface |
| **44** | public members — 25 methods and 19 properties |
| **72** | regression cases running **511** deterministic assertions, all green |
| **30** | named error constants, exposed as public enums; no bare error numbers anywhere |
| **76 %** | of the class is documentation, at procedure level |
| **2** | files to import. No reference, no add-in, no DLL |

</div>

---

<p align="center">
  <img width="1536" height="1024" alt="cPM_Home_reduced" src="https://github.com/user-attachments/assets/c4137fcb-2886-4d38-9cb8-e3349112c258" />
</p>

---

## ✨ What this project is

**Class Performance Manager** (`cPerformanceManager`) is a self-contained timing and execution-control component for Excel VBA on Windows.

It provides:

- six timing backends under one consistent, session-bound interface;
- low-overhead numeric elapsed-time measurement;
- human-readable elapsed-time formatting, with or without a second timing sample;
- a repeated-measurement harness returning the **full per-run sample vector**;
- distribution-aware statistics — median, minimum, percentiles, spread, contamination detection;
- structured named checkpoints with delta and cumulative reporting;
- shared, reference-counted suppression of expensive Excel Application behaviour;
- explicit diagnostics for QPC frequency, tick interval, and measurement overhead.

It is designed for:

- ⏱️ benchmarking VBA procedures and comparing optimisation strategies;
- 📊 instrumenting long-running workbook automation;
- 🧪 controlled execution environments that need deterministic Excel state;
- 🚀 general workbook performance work, *even when nothing is being timed*;
- 🎓 teaching how Windows timing sources actually behave.

> **Positioning**
>
> A production-oriented timing component for pure Excel VBA, built for correctness within an explicitly documented domain. It emphasises session safety, honest statistics, and predictable failure behaviour rather than attempting to match the resolution or tooling of compiled profilers.

---

## 🌟 Why this repository is different

| Capability | `Timer()` inline | Typical VBA timer helper | This project |
|---|:---:|:---:|:---:|
| Multiple timing backends | — | Sometimes | ✅ |
| Backend bound to the session and validated on read | — | — | ✅ |
| High-resolution monotonic QPC as the default | — | Sometimes | ✅ |
| Correct unsigned 32-bit counter conversion | — | Rarely | ✅ |
| Rollover-safe elapsed arithmetic | — | Rarely | ✅ |
| Backwards-clock detection | — | — | ✅ |
| Transactional session start | — | — | ✅ |
| `timeBeginPeriod` lifecycle balanced and guarded | — | Rarely | ✅ |
| Repeated measurement returning every sample | — | Rarely | ✅ |
| Median, percentile, and spread statistics | — | — | ✅ |
| Explicit contamination warning on noisy runs | — | — | ✅ |
| Structured named checkpoints | — | Rarely | ✅ |
| Overlapping instances cannot corrupt Excel state | — | — | ✅ |
| Exact Application-state restore on the last session | — | Rarely | ✅ |
| Strict and non-strict error policies | — | — | ✅ |
| Named error constants throughout | — | Rarely | ✅ |
| Deterministic regression suite | — | Rarely | ✅ |
| No external dependency | Excel runtime | Usually | ✅ |

---

## 🧭 Core ideas

<table>
<tr>
<td width="33%" valign="top">

### ⏱️ Session-bound timing

`StartTimer` binds a backend. Every elapsed read is validated against it.

Measuring with one clock and reading with another — the most common VBA timing bug — is structurally impossible.

</td>
<td width="33%" valign="top">

### 📊 Honest statistics

One sample tells you nothing about stability.

The harness returns every run, reports **median and minimum first**, and warns when variance says the number should not be trusted.

</td>
<td width="33%" valign="top">

### 🧹 Shared state safety

Excel Application settings are process-global.

A reference-counted scope manager aggregates every instance's request, so overlapping instances coexist and the original state is restored exactly once.

</td>
</tr>
<tr>
<td width="33%" valign="top">

### 🛡️ Explicit contracts

Every procedure documents its purpose, inputs, behaviour, and error policy.

Strict mode raises. Non-strict mode falls back predictably and records the reason in `LastReadStatus`, so a returned zero is always distinguishable from a real measurement of zero.

</td>
<td width="33%" valign="top">

### 🧱 Structured checkpoints

Name the phases of a workflow and get delta and cumulative timings per phase.

Export as a 2-D array, readable text, or straight to a worksheet range.

</td>
<td width="33%" valign="top">

### 📦 Frictionless deployment

Two files. Import, compile, use.

No add-in, no installer, no COM reference, no external numerical runtime.

</td>
</tr>
</table>

---

# ⏰ Timing backends

Select a backend with the first argument to `StartTimer`. The default is **5 (QPC)**.

| # | Backend | Resolution | Monotonic | Notes |
|:-:|---|---|:-:|---|
| **1** | `Timer` | ~10 ms | — | Seconds since midnight; wraps every 24 h |
| **2** | `GetTickCount` / `GetTickCount64` | ~10–16 ms | ✅ | Win64 uses the 64-bit variant; Win32 wraps at ≈49.7 days |
| **3** | `timeGetTime` | 1 ms | ✅ | Requests 1 ms multimedia resolution; wraps at ≈49.7 days |
| **4** | `timeGetSystemTime` | 1 ms | ✅ | `MMTIME` / `TIME_MS`; wraps at ≈49.7 days |
| **5** | **`QueryPerformanceCounter`** | **sub-µs** | ✅ | **Default and recommended** |
| **6** | `Now() * 86400` | ~1 s | — | Wall clock; diagnostics only |

> [!TIP]
> Use **5 (QPC)** unless you have a specific reason not to. It is the only backend with sub-microsecond resolution, and it is the one path the class optimises for.

<details>
<summary><strong>Why the rollover handling matters</strong></summary>

<br>

Backends 2, 3 and 4 are 32-bit millisecond counters on Win32. VBA stores them in a signed `Long`, so once the counter passes 2³¹ ms — roughly **24.9 days of uptime** — the value reads as negative.

`UInt32ToDouble` maps those values back into the logical unsigned range, and `RolloverSeconds` supplies the correct wrap period per backend. On Win64, backend 2 uses `GetTickCount64` and reports no rollover at all.

This is the class of bug that appears only on machines with long uptime, which is precisely when nobody is testing. Both functions are covered by dedicated boundary tests.

</details>

<details>
<summary><strong>Why negative elapsed times raise</strong></summary>

<br>

After rollover correction, a negative elapsed time can only mean the timing source moved backwards. That makes the measurement invalid, not merely imprecise.

`Elapsed_Validate` raises in strict mode and clamps to zero otherwise. Backend 5 skips the check deliberately — QPC is monotonic, and it is the default hot path.

</details>

---

# ⚡ Quick start

## 1. Import the two required files

| File | Path | Required |
|---|---|:-:|
| `M_cPM_TimeWasters` | `src/modules/M_cPM_TIMEWASTERS.bas` | ✅ |
| `cPerformanceManager` | `src/classes/cPerformanceManager.cls` | ✅ |

> [!WARNING]
> Import the **module first**. The class calls `PM_TW_NewInstanceKey`, so importing the class alone will not compile.

## 2. Time a block of code

```vb
Dim cPM As cPerformanceManager
Set cPM = New cPerformanceManager

cPM.StartTimer 5                       'QPC
    '... your workload ...
Debug.Print cPM.ElapsedTime

cPM.ResetEnvironment
Set cPM = Nothing
```

```text
0:00:02 - 431 milliseconds - 908 microseconds - 200 nanoseconds
```

## 3. Benchmark properly, with statistics

```vb
Dim cPM As cPerformanceManager
Dim S() As Double

Set cPM = New cPerformanceManager

S = cPM.MeasureProcedure("RebuildPivotCache", 30, 3)
Debug.Print cPM.Stats_Text(S, "Pivot rebuild")

cPM.ResetEnvironment
Set cPM = Nothing
```

```text
TIMING STATISTICS | Pivot rebuild
Samples           : 30
Median            : 0.412883100 s
Min               : 0.401220400 s
P95               : 0.498771300 s
Max               : 0.612004900 s
Mean              : 0.421994700 s
StdDev            : 0.038112600 s
CoeffOfVariation  : 0.0903
```

## 4. Instrument a workflow with checkpoints

```vb
cPM.StartTimer 5, False, "Import run"

LoadSourceData
cPM.Checkpoint "Load"

TransformRecords
cPM.Checkpoint "Transform", "22 400 rows"

WriteResults
cPM.Checkpoint "Write back"

Debug.Print cPM.ReportAsText
```

```text
CHECKPOINT REPORT | RunLabel=Import run
Seq | Checkpoint | DeltaSeconds | CumulativeSeconds | MethodName | Note
1 | Load | 1.204881200 | 1.204881200 | QPC |
2 | Transform | 3.771003400 | 4.975884600 | QPC | 22 400 rows
3 | Write back | 0.918224100 | 5.894108700 | QPC |
```

## 5. Speed up a heavy procedure without timing anything

```vb
cPM.TW_Turn_OFF                        'suppress Excel time-wasters
    '... heavy workbook automation ...
cPM.TW_Turn_ON                         'restore exactly as found
```

Exempt anything you still need:

```vb
cPM.TW_Turn_OFF Except:=TW_Enum.EnableEvents Or TW_Enum.Cursor
```

---

# 📊 Measurement and statistics

> [!IMPORTANT]
> **Report the median and the minimum. Treat the mean with suspicion.**
>
> Timing distributions are right-skewed with fat tails. A single OS scheduler preemption during a run inflates the mean while barely moving the median. A mean on its own cannot distinguish a clean measurement from a contaminated one.

## The harness

| Member | Returns | Purpose |
|---|---|---|
| `MeasureProcedure(Name, Iterations, Warmup, Method, [FailedReadsOut], [LastFailureStatusOut])` | `Double()` | Runs a named `Public Sub` N times via `Application.Run`; returns every per-run value, and reports how many reads failed |
| `MeasureBaseline(EmptyProcName, Iterations, Warmup, Method)` | `Double()` | An empty procedure through the **same dispatch path**, so its median can legitimately be subtracted |
| `MeasureOverhead_Samples(Iterations, Warmup, Method)` | `Double()` | Per-cycle overhead of the backend itself; its **minimum is your observed empty-cycle floor** |

## The statistics

All routines accept a **timing observation vector** — a `Double()` of finite,
non-negative values. Anything else is rejected at a single gate, naming the
offending index. The vector is yours to keep, export and post-process; it simply
has to remain a set of elapsed times.

| Member | Returns |
|---|---|
| `Stats_Median(S)` | Median — the headline statistic |
| `Stats_Min(S)` / `Stats_Max(S)` | Best and worst observed run |
| `Stats_Percentile(S, P)` | Nearest-rank percentile; every value returned was actually observed |
| `Stats_Mean(S)` | Arithmetic mean, retained for continuity |
| `Stats_StdDev(S)` | Sample standard deviation (N−1 divisor) |
| `Stats_CoefficientOfVariation(S)` | Dimensionless spread, comparable across benchmarks |
| `Stats_IsContaminated(S, [Threshold])` | `True` when CV exceeds 0.25 — re-run before trusting the result |
| `Stats_Count(S)` | Element count; returns 0 for an uninitialised array |
| `Stats_Text(S, [Caption])` | Formatted multiline summary |

<details>
<summary><strong>Reading the numbers</strong></summary>

<br>

| CV | Interpretation |
|---|---|
| **< 0.05** | Clean run. Trust the median. |
| **0.05 – 0.25** | Normal for real workloads. Compare medians, not means. |
| **> 0.25** | Contaminated. `Stats_IsContaminated` returns `True`. Close background work and re-run. |

Compare **P95 against the median** to see the tail. If P95 is close to the median, the run was stable. If it is several multiples higher, something interrupted you.

</details>

<details>
<summary><strong>Constraints of the harness</strong></summary>

<br>

- `Application.Run` reaches only **Public procedures in standard modules**. It cannot call class methods, `Private` procedures, or anything in a module declared `Option Private Module`.
- `Application.Run` adds a dispatch cost to every sample. Measure it on your own machine rather than assuming a figure — it varies by Excel version, bitness and load. The harness suits work measured in milliseconds or longer.
- `MeasureOverhead_Samples` measures the backend timing cycle only; it does **not** dispatch through `Application.Run`, so it is not a matched baseline. Use **`MeasureBaseline`** for that, and subtract medians rather than means.
- Measurement runs on an isolated worker instance, so your own session, checkpoints, and run label are never disturbed.

</details>

---

# 🧹 Time-waster suppression

Excel Application settings are **global to the process**, not to your object. Suppressing them naively from two places corrupts state for both.

`cPerformanceManager` delegates to a shared, reference-counted scope manager:

1. the **first** active session captures the original Application baseline;
2. each instance registers its own disable-mask;
3. the effective state is the **OR** of every active mask;
4. the **last** session to end restores the original baseline, exactly once.

| Flag | `TW_Enum` | Suppressed value |
|---|---|---|
| Screen updating | `ScreenUpdating` | `False` |
| Event handling | `EnableEvents` | `False` |
| Alert dialogs | `DisplayAlerts` | `False` |
| Recalculation | `Calculation` | `xlCalculationManual` |
| Mouse cursor | `Cursor` | `xlWait` |

```vb
cPM.TW_Turn_OFF                                    'suppress everything
cPM.TW_Turn_OFF Except:=TW_Enum.ScreenUpdating     'keep the screen live
cPM.TW_Turn_OFF Except:=TW_Enum.EnableEvents Or TW_Enum.Calculation
```

> [!NOTE]
> A hard `End` statement clears module globals without running `Class_Terminate`, which can leave Excel visibly suppressed. Run `PM_TW_EndAllSessions` to recover.

---

# 🧬 Public API

<details open>
<summary><strong>⏱️ Core timing</strong></summary>

<br>

| Member | Description |
|---|---|
| `StartTimer([Method], [AlignToNextTick], [RunLabel])` | Starts a session and binds the backend |
| `ElapsedSeconds([Method])` | Numeric elapsed seconds. The low-overhead primitive |
| `ElapsedTime([Method], [ElapsedSecondsIn])` | Formatted elapsed time; formats an existing value when supplied |

</details>

<details>
<summary><strong>📊 Measurement and statistics</strong></summary>

<br>

| Member | Description |
|---|---|
| `MeasureProcedure(Name, [Iterations], [Warmup], [Method], [FailedReadsOut], [LastFailureStatusOut])` | Repeated measurement of a named procedure |
| `MeasureBaseline(EmptyProcName, [Iterations], [Warmup], [Method])` | Dispatch-matched baseline |
| `MeasureOverhead_Samples([Iterations], [Warmup], [Method])` | Per-cycle backend overhead samples |
| `Stats_Median` · `Stats_Min` · `Stats_Max` · `Stats_Mean` | Central tendency and extremes |
| `Stats_Percentile` · `Stats_StdDev` · `Stats_CoefficientOfVariation` | Distribution shape |
| `Stats_IsContaminated` · `Stats_Count` · `Stats_Text` | Quality signal, size, summary |

</details>

<details>
<summary><strong>🧱 Checkpoints and reporting</strong></summary>

<br>

| Member | Description |
|---|---|
| `Checkpoint(Name, [Note])` | Captures one named checkpoint |
| `CheckpointCount` | Number captured in the current session |
| `SetRunLabel(Label)` | Labels the run — call **after** `StartTimer` |
| `RunLabel` | Current run label |
| `ClearCheckpoints` | Clears checkpoint state without ending the session |
| `ReportAsArray` | 2-D array with header row |
| `ReportAsText` | Readable multiline report |

</details>

<details>
<summary><strong>🔎 Session and state inspection</strong></summary>

<br>

| Member | Description |
|---|---|
| `T1` · `T2` · `ET` | Raw start, raw end, cached elapsed |
| `ActiveMethodID` · `HasActiveSession` | Current session binding |
| `MethodName(Index)` | Human-readable backend label |
| `StrictMode` | Get/Let the error policy |
| `LastReadStatus` | Outcome of the most recent native timing read |

</details>

<details>
<summary><strong>📉 Diagnostics</strong></summary>

<br>

| Member | Description |
|---|---|
| `OverheadMeasurement_Seconds([Method], [Iterations])` | Mean near-empty overhead |
| `OverheadMeasurement_Text([Method], [Iterations], [In])` | Formatted overhead report |
| `QPC_FrequencyPerSecond` · `QPC_FrequencyPerSecond_Value` | QPC frequency, text and numeric |
| `QPC_Get_SystemTickInterval` · `Get_SystemTickInterval` | Tick interval diagnostics |

</details>

<details>
<summary><strong>🧭 Execution control and environment</strong></summary>

<br>

| Member | Description |
|---|---|
| `Pause(Seconds, [Method])` | Four pause strategies; capped at 3 600 s |
| `ResetEnvironment` | Releases timer resolution and ends this instance's TW session |
| `TW_Turn_OFF([Except])` · `TW_Turn_ON` | Shared suppression control |
| `TW_IsActive` · `TW_ActiveSessionCount` | Suppression state |
| `TW_CalculationExempted` | Whether Calculation control could not be honoured |

</details>

---

# 🛡️ Strict mode

`StrictMode` defaults to **`True`**.

| Condition | Strict | Non-strict |
|---|---|---|
| Invalid method passed to `StartTimer` | Raises | Falls back to 5 |
| QPC requested but unavailable | Raises | Falls back to 2 |
| `ElapsedSeconds` before `StartTimer` | Raises | Returns 0 |
| Read method ≠ session method | Raises | Uses the session method |
| Negative elapsed after rollover correction | Raises | Clamps to 0 |
| QPC read failure | Raises | Returns 0 |
| Tick alignment exceeds the spin guard | Raises | Returns a current timestamp |
| Native read fails during a session start | Raises | Falls back to backend 2 before committing |
| Native read fails during an elapsed read | Raises | Returns 0 and records the reason in `LastReadStatus` |

Every raised error number is a named constant, so no bare `vbObjectError` offset appears anywhere in the code, and they are exposed as the public `cPM_Error` and `cPM_TWError` enums so a caller can trap a condition by name.

Backends and pause strategies are named too — `cPM_TimerMethod` and `cPM_PauseMethod`. The two share the numbers 1–4 and mean entirely different things, so separate types make them non-interchangeable at compile time.

---

# 🚦 Reading `LastReadStatus`

A returned zero means nothing on its own. The status is what qualifies it.

| Value | Meaning |
|---|---|
| `cPM_ReadOK` | The value is a real measurement |
| `cPM_ReadQpcFailed` | `QueryPerformanceCounter` failed |
| `cPM_ReadSystemTimeFailed` | `timeGetSystemTime` failed |
| `cPM_ReadSystemTimeFormatInvalid` | `timeGetSystemTime` returned an unexpected format |
| `cPM_ReadFallbackToMethod2` | The requested backend was unreadable; backend 2 was bound instead |
| `cPM_ReadElapsedInvalid` | The timing source moved backwards; the value was clamped |

`cPM_ReadFallbackToMethod2` is not a failure — the session is valid and the
measurement usable. It records that the requested backend was coerced, so the
coercion is visible rather than silent.

> [!IMPORTANT]
> `LastReadStatus` describes **this instance's own reads**. The measurement
> harness runs on an isolated worker released before the sample vector is
> returned, so its outcome arrives through `FailedReadsOut` and
> `LastFailureStatusOut` instead.

---

# 📉 What the harness does with a failed read

In non-strict mode a failed read returns zero. Storing that zero would make the
failure look like the fastest run in the set, so it is **excluded** rather than
recorded.

| Consequence | |
|---|---|
| `Stats_Count(Samples)` | May be less than `Iterations` |
| `Iterations - Stats_Count` | Is the measured-failure count |
| `FailedReadsOut` | Counts every failed read, warm-up included, so it can exceed the shortfall |
| No measured read succeeded | Raises `ERR_CPM_MEASURE_NO_VALID_SAMPLES` — an empty vector is not a measurement |

---

# 🧪 Testing and validation

```vb
Run_cPerformanceManager_RegressionSuite
```

**72 cases · 511 assertions**, written to a dedicated worksheet log and summarised in the Immediate Window.

Last certified run: **0 failures**, 2026-08-16, on Excel for Microsoft 365 MSO
Version 2606 Build 16.0.20131.20152, 64-bit.

> [!NOTE]
> A run certifies one bitness. On 64-bit Office, backend 2 compiles to
> `GetTickCount64`, so the 32-bit wrap-correction path is compiled out and not
> exercised.
>
> The *semantics* of that path are shared: the signed-to-unsigned
> reinterpretation is `UInt32ToDouble`, verified at all four boundary values on
> whichever bitness the suite runs. What remains bitness-specific is the
> declaration binding, the API return, and the `RolloverSeconds` constant — a
> materially smaller surface than the whole branch, but not zero.
>
> **32-bit is therefore supported by construction and by shared arithmetic, but
> is not execution-certified.** Certification before a major release is tracked
> as a release step.

Coverage includes:

- every backend across start, elapsed, aligned-start, and formatted output;
- strict and non-strict behaviour on every failure path;
- `UInt32ToDouble` at all four 32-bit boundaries;
- `RolloverSeconds` per backend, with the Win64 branch handled conditionally;
- checkpoint storage integrity across 1 000 captures and ~10 `ReDim Preserve` cycles;
- `Class_Terminate` releasing a shared TW session with all five Application flags restored;
- 75 instance create/destroy cycles proving no stale TW registration survives;
- statistics against a hand-computed vector, plus order independence and boundary behaviour;
- **injected native-read failures** on both backends, covering strict-mode raises, non-strict fallback to backend 2, cache preservation, and checkpoint abandonment;
- an injected **wrong-format** `timeGetSystemTime` result, proving it is reported distinctly from an outright read failure;
- Calculation baseline validity, deliberate exemption, overlapping scopes, and proof that no synthetic baseline is ever written.

---

# 🏗️ Repository contents

```text
src/classes/cPerformanceManager.cls      Required — the class
src/modules/M_cPM_TIMEWASTERS.bas        Required — shared TW manager
test/M_cPM_Test.bas                      Regression suite
demo/M_cPM_DEMO.bas                      Interactive demo
demo/M_cPM_USAGE_EXAMPLES.bas            Worked examples
demo/M_DEMO_BUILDER.bas                  Presentation helpers
CHANGELOG.md                             Version history
```

> [!NOTE]
> The demo workbook is distributed as a **[Release asset](https://github.com/danielep71/vba-performance_manager/releases)** rather than versioned in the repository.

---

# 🧩 Requirements

| Requirement | Detail |
|---|---|
| Host | Microsoft Excel on **Windows** |
| Bitness | 32-bit and 64-bit, via `VBA7` / `Win64` conditional compilation |
| References | **None.** `Scripting.Dictionary` is late-bound |
| Dependencies | No add-in, installer, DLL or COM component. Both source files are required together; the regression suite additionally needs `M_DEMO_BUILDER` |

---

# 🛠️ Installation

1. Open the VBE with <kbd>Alt</kbd> + <kbd>F11</kbd>
2. **File → Import File…** → `src/modules/M_cPM_TIMEWASTERS.bas`
3. **File → Import File…** → `src/classes/cPerformanceManager.cls`
4. **Debug → Compile VBAProject**

Optionally import `test/M_cPM_Test.bas` and the `demo/` modules to run the suite and the examples.

---

# 🧠 Design notes

<details>
<summary><strong>Transactional session start</strong></summary>

<br>

`StartTimer` captures every timestamp into locals first and commits class state only after all fallible operations succeed. A failed QPC read under strict mode therefore leaves the previous session completely untouched.

</details>

<details>
<summary><strong>Collision-proof instance keys</strong></summary>

<br>

Shared TW keys are issued by a module-level counter, not derived from `ObjPtr`. VBA reuses heap addresses, so an address-based key could let a new instance silently inherit a session left behind by a destroyed one.

The counter and the session store share a module-global lifetime, so a project reset clears both together and a key can never be reissued while a stale registration survives.

</details>

<details>
<summary><strong>Currency arithmetic for QPC</strong></summary>

<br>

QPC ticks and frequency are stored as `Currency`, which VBA holds as a scaled 64-bit integer. Because both the tick delta and the frequency carry the same scaling, it cancels in the division and the elapsed value is exact.

</details>

<details>
<summary><strong>Single-site backend dispatch</strong></summary>

<br>

All elapsed-time dispatch lives in one private reader. `ElapsedSeconds` delegates and caches; `Checkpoint` delegates without caching, so taking a checkpoint never overwrites an explicit measurement you are still holding.

</details>

---

# ⚠️ Known limitations

- **Windows only** for backends 2–5. Backends 1 and 6 are conceptually portable.
- **Backend 1 cannot distinguish a backward clock adjustment from midnight rollover.** Both appear as a negative raw delta, and both get 24 hours added. Use backend 5 where this matters.
- **Rollover correction handles one wrap.** A session spanning more than a full 32-bit millisecond wrap on backends 2, 3 or 4 cannot be recovered by a single addition.
- **`Application.Run` dispatch cost** is included in every `MeasureProcedure` sample. Subtract a baseline from `MeasureBaseline`, not from `MeasureOverhead_Samples`, which never dispatches.
- **An unqualified procedure name** resolves against the workbook hosting the class, not the active workbook. Pass a qualified name such as `'Other.xlsm'!Proc` to measure elsewhere.
- **`MeasureProcedure` executes the name it is given.** It reaches
  `Application.Run`, so a name taken from a worksheet cell, a configuration
  sheet or a file lets that source run arbitrary code inside your project. Pass
  literals, or names your own code controls. Validating the string would not
  help — any valid procedure name is executable by design.
- **`timeBeginPeriod(1)`** affects system-wide timer resolution while held and is released by `ResetEnvironment`, with `Class_Terminate` as the fallback.
- **A hard `End` statement** bypasses `Class_Terminate`; recover with `PM_TW_EndAllSessions`.
- **Calculation control requires a stable open-workbook set** for the life of a suppression scope. Where that does not hold, the flag is exempted and `TW_CalculationExempted` reports it.
- **`Pause` methods 3 and 4** issue no coarse wait for requests under roughly two seconds, to guarantee they never overshoot.
- **Nanosecond display precision** is presentational and does not imply measurement resolution at that scale.
- **Statistics are descriptive, not inferential.** `Stats_IsContaminated` is a heuristic on the coefficient of variation; it flags runs worth repeating, it does not prove a result wrong.

---

# 📚 Wiki

Full technical documentation lives in the **[Wiki](https://github.com/danielep71/vba-performance_manager/wiki)**:

| Page | Contents |
|---|---|
| [Installation and Import](https://github.com/danielep71/vba-performance_manager/wiki/Installation-and-Import) | Import order and setup |
| [Quick Start](https://github.com/danielep71/vba-performance_manager/wiki/Quick-start) | First measurement in five minutes |
| [Timer Methods](https://github.com/danielep71/vba-performance_manager/wiki/Timer-Methods) | Backend characteristics and selection |
| [Core API](https://github.com/danielep71/vba-performance_manager/wiki/Core-API) | Session model and the main members |
| [Statistics and Measurement](https://github.com/danielep71/vba-performance_manager/wiki/Statistics-and-Measurement) | The harness and reading the numbers |
| [Checkpoint and Reporting](https://github.com/danielep71/vba-performance_manager/wiki/Checkpoint-and-Reporting) | Structured instrumentation |
| [Time-Waster Suppression](https://github.com/danielep71/vba-performance_manager/wiki/Time-Waster-Suppression) | The shared scope model |
| [Benchmarking Guidance](https://github.com/danielep71/vba-performance_manager/wiki/Benchmarking-Guidance) | Getting trustworthy numbers |
| [Architecture and Internal Design](https://github.com/danielep71/vba-performance_manager/wiki/Architecture-and-Internal-Design) | How it works inside |
| [Testing and Validation](https://github.com/danielep71/vba-performance_manager/wiki/Testing-and-Validation) | The regression suite |
| [Known Limitations](https://github.com/danielep71/vba-performance_manager/wiki/Known-Limitations) | Documented boundaries |
| [Version History](https://github.com/danielep71/vba-performance_manager/wiki/Version-History) | Release record |

---

<div align="center">

## 📄 License

Released under the [MIT License](LICENSE).

## 👤 Author

**Daniele Penza**

[![GitHub](https://img.shields.io/badge/GitHub-danielep71-181717?style=for-the-badge&logo=github)](https://github.com/danielep71)

<br>

*If this component saved you time, consider starring the repository.*

⭐

</div>
