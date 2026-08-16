---
name: ✨ Feature request
about: Propose a new capability or a change to the public surface
title: "[Feature] "
labels: enhancement
assignees: ''
---

> [!TIP]
> **Start from the problem, not the solution.**
>
> A clear description of what you cannot do today usually leads to a better
> design than the one either of us had in mind first. Several members of this
> class exist in a different shape than originally proposed, because the
> underlying problem turned out to be something adjacent.

---

## 🎯 The problem

<!--
What you are trying to measure or control, and where the current API stops you.
Describe the situation rather than the feature.
-->



---

## 🧩 Proposed surface

<!--
What you would like to call, and what it would return. A rough signature is more
useful than a precise one — the shape matters more than the spelling.
-->

```vb
' e.g.
S = cPM.MeasureProcedure("Transform", 30, 3, cPM_MethodQPC, _
                         SetupProcedureName:="ResetSourceData")
```

---

## 🔧 Current workaround

<!--
How you handle this today and what it costs — a wrapper procedure per parameter
set, a hand-written loop, giving up on the measurement entirely.

If there is no workaround, say so. That is useful information.
-->



---

## ⚖️ Backwards compatibility

- [ ] **Purely additive** — new members, or new optional arguments on existing ones
- [ ] **Changes existing behaviour** — a caller relying on today's behaviour would notice
- [ ] Not sure

<details>
<summary><strong>Why this question decides the release it lands in</strong></summary>

<br>

Additive changes ship in a **minor** release. Changes to existing behaviour
normally require a **major** one.

There is one documented exception: v1.3.0 changed
`Stats_CoefficientOfVariation` from returning `0` to raising, because the old
return value could certify a failed measurement as clean. Preserving it for
compatibility would have preserved a correctness defect.

That bar is deliberately high. "The old behaviour was wrong in a way that
misleads" clears it; "the new behaviour is nicer" does not.

</details>

---

## 🔀 Alternatives considered

<!--
Other shapes this could take, and why you prefer the one above. Naming the
trade-offs usually shortens the discussion considerably.
-->



---

## 📐 Scope check

<!-- Optional, but helps place the request. -->

- [ ] Fits the existing session-bound timing model
- [ ] Needs new Windows API surface
- [ ] Affects the shared time-waster scope
- [ ] Affects the measurement harness or statistics
- [ ] Documentation or examples only

---

## 📎 Anything else

<!--
Prior art in other timing libraries, a link to the workflow that prompted this,
or a sketch of how you would test it.
-->
