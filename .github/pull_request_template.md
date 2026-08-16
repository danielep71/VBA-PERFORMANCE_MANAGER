<!--
The static checks run automatically and gate this pull request. They analyse the
sources as text and cannot execute a single test, so the regression evidence
below is the only record that the code actually runs.
-->

## What this changes

<!-- One paragraph. What behaviour is different, and why. -->

Closes #

## Regression suite

<!--
Run `Run_cPerformanceManager_RegressionSuite` against the exact code in this
branch, not an earlier build. Leave this section as-is if the change touches no
VBA source.
-->

| | |
|---|---|
| Cases | |
| Assertions | |
| Failures | |
| Excel version and build | |
| Bitness | |

<!--
On 64-bit Office, backend 2 compiles to GetTickCount64, so the 32-bit
wrap-correction branch of RolloverSeconds is compiled out and is not exercised.
Worth noting if this change touches rollover arithmetic.
-->

## Checklist

- [ ] Compiles cleanly — Debug → Compile VBAProject
- [ ] Regression suite run against this branch, result recorded above
- [ ] `TotalSteps` matches the number of cases, if cases were added
- [ ] `CHANGELOG.md` updated under `[Unreleased]`, if the public surface changed
- [ ] Version stamps left alone — they move once, at release
- [ ] README and wiki updated, if a documented guarantee changed

## Anything reviewers should look at closely

<!--
Judgement calls, trade-offs you were unsure about, or behaviour that a passing
suite would not catch. Both v1.2.0 blocker defects passed every automated check;
they were found by review. This section is where that happens.
-->
