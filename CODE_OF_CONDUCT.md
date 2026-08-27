<div align="center">

# 🧭 Code of Conduct

### Respectful, evidence-led technical collaboration for VBA-PERFORMANCE_MANAGER

[![Applies to](https://img.shields.io/badge/Applies_to-Everyone-217346?style=for-the-badge)](#scope)
[![Spaces](https://img.shields.io/badge/Spaces-Issues_PRs_Wiki_Releases-0969da?style=for-the-badge)](#scope)
[![Standard](https://img.shields.io/badge/Standard-Respectful_%2B_Evidence--Led-6f42c1?style=for-the-badge)](#technical-discussion-standards)
[![Enforcement](https://img.shields.io/badge/Enforcement-Maintainer-d97706?style=for-the-badge)](#enforcement)

<br>

**Technical rigor · Respectful disagreement · Reproducible evidence · Honest benchmarking · Privacy-aware collaboration**

</div>

---

**VBA-PERFORMANCE_MANAGER** is a focused open-source Excel/VBA timing and
execution-control project.

This Code of Conduct exists to keep interaction around the project respectful,
technical, constructive, and welcoming — especially when a problem depends on
Windows timing behavior, Excel Application state, Office bitness, process
uptime, benchmark methodology, or conditions that are difficult to reproduce.

People should feel comfortable:

- reporting defects;
- asking basic or advanced VBA questions;
- challenging timing or statistical claims;
- questioning a benchmark methodology;
- proposing safer or simpler alternatives;
- identifying behavior that differs by Office version, bitness, backend, or
  Excel session state;
- saying that an earlier assumption, benchmark, or implementation decision was
  wrong;
- contributing even when they are unfamiliar with the project's conventions.

A technically demanding project benefits from disagreement.

It does not benefit from hostility.

---

<a id="our-pledge"></a>

## 🤝 Our pledge

Everyone who participates — by opening an issue, submitting a pull request,
commenting, reviewing, editing documentation or the Wiki, discussing a release,
or representing the project elsewhere — is expected to help create a
harassment-free experience for all.

That expectation applies regardless of:

- experience level;
- professional or academic background;
- age;
- disability;
- ethnicity;
- gender identity or expression;
- nationality;
- race;
- religion;
- sexual orientation;
- socioeconomic status;
- or any other personal characteristic unrelated to the technical contribution.

Technical rigor and respectful interaction are complementary requirements.

Neither excuses the absence of the other.

---

<a id="expected-behavior"></a>

## ✅ Expected behavior

Participants are expected to:

- be respectful and assume good faith unless evidence shows otherwise;
- focus criticism on code, behavior, documentation, tests, architecture,
  measurements, or process rather than on the person who produced them;
- distinguish **observed fact**, **measurement**, **inference**, **hypothesis**,
  and **platform limitation**;
- describe Excel/VBA and Windows timing behavior precisely;
- provide reproduction steps, logs, screenshots, sample vectors, regression
  evidence, or minimal examples where practical;
- state the timing backend, Office bitness, and relevant measurement conditions
  when making performance claims;
- acknowledge uncertainty rather than presenting an assumption as verified;
- correct mistakes openly when new evidence changes the conclusion;
- give and receive constructive review comments professionally;
- respect privacy, confidentiality, client restrictions, and security boundaries;
- help newcomers understand the repository's workflow and vocabulary;
- allow maintainers reasonable time to investigate host-specific or
  timing-sensitive behavior;
- respect that a contribution may be adopted, adapted, deferred, split into a
  follow-up issue, or declined to preserve correctness, compatibility,
  maintainability, or release discipline.

### Useful disagreement

A useful technical disagreement is specific enough that another person can test
it:

> "`MeasureProcedure` requested QPC, but one measured cycle resolved to method 2
> after a forced start failure. The vector was then summarized as if every
> observation used QPC. The exact SHA, Excel build, bitness, requested method,
> resolved method, and regression output are below."

That can be investigated.

### Unhelpful disagreement

A personal judgment cannot be tested:

> "This benchmark is wrong because the author does not understand timing."

Both statements may arise from frustration with the same problem.

Only the first helps fix it.

---

<a id="unacceptable-behavior"></a>

## 🚫 Unacceptable behavior

Unacceptable behavior includes:

- personal attacks, insults, ridicule, or derogatory comments;
- harassment in public or private;
- discriminatory, demeaning, or sexualized language or imagery;
- threats, intimidation, or encouragement of violence;
- publishing another person's private information without permission;
- deliberate misrepresentation of another participant's work or statements;
- knowingly fabricating benchmark results, regression output, screenshots,
  provenance evidence, reproduction steps, or implementation claims;
- cherry-picking measurements while presenting them as representative;
- sustained disruption of technical discussion;
- repeated bad-faith argument after the technical decision and evidence have
  been explained;
- spam, unrelated promotion, or commercial solicitation;
- attempts to pressure maintainers into unsafe disclosure, unverifiable claims,
  or an unverified release;
- public disclosure of a suspected vulnerability before reasonable coordinated
  remediation;
- retaliation against someone who reports misconduct, a security concern, a
  benchmark-validity concern, or a technical failure.

Disagreement is allowed.

Abuse is not.

---

<a id="technical-discussion-standards"></a>

## 🧪 Technical discussion standards

VBA-PERFORMANCE_MANAGER interacts with timing and Excel surfaces whose behavior
can depend on the host process and execution context.

Relevant surfaces include:

```text
Timer
GetTickCount / GetTickCount64
timeGetTime
timeGetSystemTime
QueryPerformanceCounter
Now
timeBeginPeriod / timeEndPeriod
Application.ScreenUpdating
Application.EnableEvents
Application.DisplayAlerts
Application.Calculation
Application.Cursor
Application.Run
DoEvents
```

A technical report is therefore more useful when its environment is explicit.

### For timing and runtime behavior

Where relevant, include:

- exact repository tag, branch, or commit SHA;
- Excel version and build;
- Office 32-bit or 64-bit;
- Windows version;
- requested timing backend;
- `ActiveMethodID` / resolved backend when different;
- `StrictMode`;
- `LastReadStatus`;
- whether aligned start was requested;
- process uptime where a 32-bit rolling counter may matter;
- whether another `cPerformanceManager` instance is active;
- whether TW suppression is active;
- whether `Application.Calculation` control was exempted;
- reproduction steps;
- observed behavior;
- expected behavior;
- regression evidence, logs, screenshots, or minimal workbooks when safe to
  share.

Do not infer a backend, failure cause, or source-resolution behavior merely from
a plausible numeric result when the project exposes a status or diagnostic that
can be inspected directly.

Where the platform makes a fact unavailable or impractical to verify, say so and
classify the evidence appropriately:

```text
automated regression
manual Excel certification
source inspection
Windows / VBA platform contract
benchmark observation
unresolved hypothesis
```

That distinction is part of this project's quality standard.

---

## ⏱️ Benchmark and measurement standards

Performance claims require more discipline than ordinary functional claims.

When presenting a benchmark, identify enough information for another person to
understand what was actually measured.

Where relevant, include:

```text
commit / version
Excel build
Office bitness
timing backend
requested iterations
warm-up iterations
retained sample count
failed / excluded reads
fallbacks or resolved-backend changes
median
minimum
P95 or another tail measure
coefficient of variation / contamination result
baseline method, if one was subtracted
```

### Do not confuse clock properties with harness results

The minimum returned by an empty timing cycle is not automatically the clock's
resolution.

A measurement harness may include:

```text
class dispatch
start read
end read
loop overhead
Application.Run dispatch
warm-up behavior
Excel scheduling noise
OS scheduling noise
```

Use the project's diagnostic functions when discussing the timing source itself,
and benchmark observations when discussing the whole measurement path.

### Matched baselines

A baseline is meaningful only when it matches the execution path whose overhead
is being estimated.

`MeasureBaseline` exists to use the same `Application.Run` dispatch path as
`MeasureProcedure`.

Do not describe `MeasureOverhead_Samples` as a matched `Application.Run`
baseline: it measures a different path.

### Failed or fallback reads are not ordinary samples

A failed timing read is not a zero-duration measurement.

A fallback backend is not automatically equivalent to the backend originally
requested.

When a harness exposes failure counts, statuses, or resolved-method information,
include them in the interpretation rather than discarding them.

### Representative reporting

Do not present a single exceptionally low observation as representative without
saying that it is a minimum.

For optimization claims, prefer the full sample distribution or at least a
summary that makes variability visible.

---

## ✅ Regression evidence

The current regression harness is:

```text
test/M_cPM_Test.bas
```

Current public entry point:

```vb
Run_cPerformanceManager_RegressionSuite
```

The suite covers both the class and its required shared-state companion module.

For production changes, contributors should normally compile first:

```text
VBA Editor → Debug → Compile VBAProject
```

and then run the regression suite against the exact source being proposed.

A useful report includes at least:

```text
Cases
Assertions
Failures
Excel version/build
Office bitness
commit SHA
```

For release-quality evidence:

```text
Failures = 0
```

is required.

The count alone is not enough. A statement such as:

```text
511 assertions passed
```

is incomplete if it omits:

- whether the suite finished;
- whether failures occurred elsewhere;
- which commit was tested;
- which Excel build and bitness executed the run.

If a future headless Excel regression gate publishes machine-readable evidence,
that artifact should be preferred over manually transcribed counts.

Static source checks are useful, but they do not substitute for importing,
compiling, and executing the VBA in Excel.

---

## 🔍 Review standards

Review comments should be actionable whenever possible.

A strong review comment identifies:

1. **where** the concern exists;
2. **what** behavior or invariant is at risk;
3. **why** it matters;
4. **what evidence** supports the concern;
5. whether the requested change is:
   - required;
   - recommended;
   - optional;
6. whether the concern affects:
   - correctness;
   - timing validity;
   - benchmark methodology;
   - compatibility;
   - Excel state safety;
   - diagnostics;
   - test validity;
   - documentation;
   - release provenance;
   - security;
   - maintainability;
   - style.

Example:

> "`OverheadMeasurement_Seconds` averages requested iterations even when a
> non-strict endpoint read fails and returns zero. That lowers the reported mean.
> The compatibility wrapper should delegate to the validated sample path so the
> denominator reflects retained observations."

This is preferable to:

> "I don't like this benchmark API."

---

## 🧹 Excel Application-state discussions

The component can change process-wide Excel settings for performance.

Those settings affect more than the calling procedure.

When discussing TW behavior, distinguish at least:

```text
baseline capture
baseline validity
per-instance request
aggregate effective state
overlapping instances
Calculation exemption
strict vs non-strict policy
last-session restoration
emergency cleanup
```

Do not describe global Application-state manipulation as harmless merely because
the calling macro is local to one workbook.

When reporting a restoration problem, include:

- the starting Application state;
- which TW flags were requested;
- how many manager instances were active;
- whether workbooks were opened or closed during the scope;
- whether `TW_CalculationExempted` became true;
- whether cleanup completed normally or through recovery.

The project prefers:

```text
explicit state ownership
```

over:

```text
implicit global side effects
```

and:

```text
reported inability to restore/control a state
```

over:

```text
guessing a replacement value
```

Technical discussion should preserve those distinctions.

---

## 📦 Source-first and release-evidence discussions

The repository is source-first.

Authoritative implementation lives in exported source such as:

```text
src/classes/cPerformanceManager.cls
src/modules/M_cPM_TIMEWASTERS.bas
```

The demo workbook distributed through GitHub Releases is a convenience artifact,
not the authoritative source tree.

When discussing a release, distinguish:

```text
tag / commit identity
source hashes
static-check evidence
Excel regression evidence
Office bitness
workbook digest
source-to-workbook provenance
```

A SHA-256 digest proves the identity of a file.

It does not, by itself, prove how that file was produced.

Do not claim reproducibility, execution certification, or source provenance
beyond what the published evidence actually establishes.

---

<a id="privacy-and-confidentiality"></a>

## 🔒 Privacy, confidentiality, and safe reproductions

Do not upload confidential business material merely to demonstrate a timing,
state-management, or benchmark problem.

That includes:

- client workbooks;
- proprietary VBA;
- credentials;
- private signing keys;
- personal data;
- internal URLs;
- connection strings;
- production data extracts;
- non-public add-ins or modules you are not authorized to distribute.

Create a sanitized minimal reproduction instead.

Excel files can contain more than visible worksheet values, including:

```text
defined names
external links
connection strings
Power Query metadata
cached values
comments
document properties
hidden sheets
VBA
```

A workbook that appears anonymized may still disclose information elsewhere in
the package.

If a private reproduction is genuinely necessary, coordinate privately with the
maintainer before sharing it.

---

## 🔑 Security issues are different from ordinary bugs

A dedicated `SECURITY.md` is not currently part of the repository.

Until one is added, suspected vulnerabilities should be reported **privately**
to the maintainer rather than opened as a normal public issue when disclosure
could expose users, credentials, signing material, or an exploitable condition.

Do **not** publish exploit details, credentials, private keys, or a working proof
of concept in a normal public issue before coordinated disclosure.

The Code of Conduct reporting channel is for participant behavior.

A security report is about software risk.

If an incident involves both, use the private channel and say so.

---

<a id="scope"></a>

## 🛠️ Scope

This Code of Conduct applies to:

- this GitHub repository;
- issues;
- pull requests;
- review comments;
- GitHub Discussions, if enabled;
- release threads;
- the Wiki;
- project-related email;
- project-related private communication between participants;
- public spaces where someone is representing the project.

Examples of project representation include:

- speaking on behalf of the project;
- using an official project account;
- presenting oneself as a maintainer or contributor in a project-related forum;
- moderating a project discussion.

The standards apply to both maintainers and contributors.

---

<a id="reporting-unacceptable-behavior"></a>

## 📣 Reporting unacceptable behavior

Report unacceptable behavior **privately** to the maintainer:

```text
danielep71@gmail.com
```

Where available, include:

- what happened;
- where it happened;
- dates or approximate times;
- links, screenshots, or quoted text;
- whether the behavior is ongoing;
- whether another participant witnessed it;
- any immediate safety, privacy, or confidentiality concern.

Do not post sensitive personal information in a public issue.

Reports will be handled as discreetly as reasonably possible.

Information will be shared only as needed to:

```text
understand the report
protect participants
enforce this policy
comply with platform or legal requirements where applicable
```

A good-faith report will not be treated as misconduct merely because the
maintainer ultimately concludes that no violation occurred.

---

<a id="enforcement"></a>

## ⚖️ Enforcement

The maintainer is responsible for interpreting and enforcing this Code of
Conduct.

Responses depend on seriousness, frequency, context, prior behavior, and risk to
participants or the project.

Possible actions include:

1. clarification or a private reminder;
2. a formal warning;
3. editing or removing comments or contributions;
4. closing or locking a discussion;
5. rejecting or reverting a contribution;
6. temporary restriction from project participation;
7. permanent blocking;
8. escalation to GitHub or another relevant platform.

Enforcement aims to be:

```text
proportionate
consistent
documented where appropriate
protective of participants
protective of the technical record
```

Retaliation against a reporter, witness, or participant in an investigation is
itself a violation.

---

## 🧩 Conflicts of interest

Participants should disclose a material conflict when it could reasonably affect
technical review.

Examples:

- ownership of a competing implementation;
- commercial interest in a dependency, benchmark, or integration being
  proposed;
- employment/client restrictions that materially limit what can be disclosed;
- inability to establish the origin or license of copied code;
- evaluating a contribution one personally authored under another identity or
  organization.

A conflict is not automatically disqualifying.

Undisclosed material influence is the problem.

---

## 📜 Source and licensing integrity

Contributors must have the right to submit the code, documentation, screenshots,
benchmark data, or other material they provide.

Do not submit:

- proprietary code copied from an employer or client;
- code with an incompatible license without clear attribution and discussion;
- screenshots containing confidential data;
- benchmark evidence that has been altered or selectively presented to hide
  contrary results;
- generated code presented as independently authored when its provenance or
  license is uncertain;
- binary workbook changes that bypass the repository's source-first architecture.

If material was adapted from another source, identify that source and its license
clearly enough for review.

---

## 🧱 Maintainer decisions

A maintainer may decline a contribution even when it is technically valid.

Reasons can include:

```text
scope
compatibility
maintenance burden
testability
platform risk
API stability
benchmark integrity
release timing
duplication
architectural direction
```

A declined contribution is not a judgment about the contributor.

When practical, the technical reason should be recorded so that the same design
question does not need to be rediscovered repeatedly.

Participants may challenge a decision respectfully with new evidence.

Repeatedly reopening the same argument without new evidence is not constructive.

---

## 🙏 Project scale and response expectations

VBA-PERFORMANCE_MANAGER is maintained by one person.

That affects response capacity, not the seriousness of this policy.

Response times are best-effort.

Complex reports may take longer when they require:

- a particular Office bitness;
- a clean Excel process;
- long process uptime;
- a specific timing backend;
- fault-injection behavior;
- overlapping manager instances;
- workbook-lifecycle transitions;
- a second Excel host or VM;
- Windows API behavior that cannot be reproduced in hosted Linux CI.

Reasonable delay is not dismissal.

Repeatedly demanding immediate action is not an acceptable substitute for
technical evidence.

Where appropriate, GitHub's platform policies and community standards also
apply.

---

## 🧭 Practical principle

The standard for this project can be summarized as:

```text
be precise about the software
be honest about the measurement
be generous toward the person
show the evidence
state the uncertainty
protect user data
```

That is the environment in which difficult Excel/VBA timing and state-management
problems are most likely to be solved well.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
