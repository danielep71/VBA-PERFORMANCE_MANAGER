---
name: 🐛 Bug report
about: Something behaves incorrectly or produces a wrong measurement
title: "[Bug] "
labels: bug
assignees: ''
---

> [!IMPORTANT]
> **Timing behaviour cannot be diagnosed from a description alone.**
>
> Several code paths are selected by `#If Win64`, so a defect may exist on one
> bitness and not the other. A report saying *"the elapsed time is wrong"*
> without the environment section below usually cannot be acted on.
>
> Everything asked for here takes under a minute to gather.

---

## 🔍 What happened

<!-- What you expected, and what you got instead. -->



**Expected:**

**Actual:**

---

## 🧪 Minimal reproduction

<!--
The smallest procedure that shows the problem. Please include the StartTimer
call and how the result is read — the interaction between them is where most
defects live.
-->

```vb
Public Sub Repro()

    Dim cPM As cPerformanceManager
    Set cPM = New cPerformanceManager

    cPM.StartTimer cPM_MethodQPC
    '... your workload ...
    Debug.Print cPM.ElapsedSeconds

    cPM.ResetEnvironment
    Set cPM = Nothing

End Sub
```

---

## 💻 Environment

| | |
|---|---|
| **Component version** | <!-- VERSION stamp in cPerformanceManager.cls, or the release tag --> |
| **Excel version and build** | <!-- File → Account → About Excel. Paste the whole line --> |
| **Bitness** | <!-- 32-bit or 64-bit, shown at the end of the same line --> |
| **Windows version** | <!-- optional --> |

> [!NOTE]
> The **build number** matters, not just the year. `Version 2606 Build
> 16.0.20131.20152` is the useful form; `Microsoft 365` on its own is not.

---

## ⏱️ Session state

| | |
|---|---|
| **Backend requested** | <!-- the value passed to StartTimer, e.g. cPM_MethodQPC (5) --> |
| **`ActiveMethodID`** | <!-- Debug.Print cPM.ActiveMethodID --> |
| **`StrictMode`** | <!-- True (default) / False --> |
| **`LastReadStatus`** | <!-- Debug.Print cPM.LastReadStatus --> |

<details>
<summary><strong>Why each of these matters</strong></summary>

<br>

**`ActiveMethodID` differing from what you requested** means a non-strict
fallback fired. The class bound a different clock than you asked for, and that
is very often the entire explanation.

**`StrictMode = False`** substitutes documented fallbacks instead of raising, so
a returned `0` may be a suppressed failure rather than a fast operation.

**`LastReadStatus`** distinguishes a real measurement from a failed read. Note
that it does **not** apply to samples returned by `MeasureProcedure`, which
measures on an isolated instance that is released before the vector comes back.

</details>

---

## ✅ Import checklist

- [ ] `M_cPM_TIMEWASTERS.bas` was imported **before** `cPerformanceManager.cls`
- [ ] **Debug → Compile VBAProject** completes with no errors
- [ ] The class and module are the same version

> [!WARNING]
> Import order is the single most common cause of unexplained behaviour. The
> class calls `PM_TW_NewInstanceKey`, so importing it first will not compile.

---

## 🧾 Regression suite

<!-- If you imported test/M_cPM_Test.bas. A failing suite narrows this down a lot. -->

- [ ] All cases pass
- [ ] Some cases fail — names below
- [ ] Not run
- [ ] Test module not imported

**Failing cases:**

---

## 📎 Anything else

<!--
Workbook size, whether the machine was under load, regional settings if the
problem involves formatted output, screenshots of the Immediate Window.
-->
