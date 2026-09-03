<div align="center">

# 🔒 Security Policy

### Security, trust boundaries, responsible disclosure, and safe deployment for VBA-PERFORMANCE_MANAGER

[![Reporting](https://img.shields.io/badge/reporting-private-d97706?style=for-the-badge)](#-reporting-a-vulnerability)
[![Support](https://img.shields.io/badge/support-latest_tagged_release-217346?style=for-the-badge)](#-supported-versions)
[![Platform](https://img.shields.io/badge/platform-Excel_VBA_%2F_Windows-0078D6?style=for-the-badge)](#-security-model)
[![Scope](https://img.shields.io/badge/scope-runtime_release_and_automation-6f42c1?style=for-the-badge)](#-scope)
[![Automation](https://img.shields.io/badge/automation-least_privilege-d73a49?style=for-the-badge)](#-repository-automation-and-runner-security)

<br>

**Source-first trust · Explicit timing status · Trusted macro dispatch · Shared-state ownership · Private disclosure · Release provenance**

<br>

[Supported versions](#-supported-versions)
&nbsp;·&nbsp;
[Report privately](#-reporting-a-vulnerability)
&nbsp;·&nbsp;
[Security scope](#-scope)
&nbsp;·&nbsp;
[Runtime boundaries](#-runtime-security-boundaries)
&nbsp;·&nbsp;
[Supply chain](#-supply-chain-and-release-integrity)
&nbsp;·&nbsp;
[Automation](#-repository-automation-and-runner-security)
&nbsp;·&nbsp;
[Verify a release](#-verifying-a-release)

</div>

---

**VBA-PERFORMANCE_MANAGER** distributes reviewable Excel/VBA source and, for
convenience, macro-enabled release binaries.

The production component runs with the privileges already granted to Microsoft
Excel and the current Windows user.

There is no:

- background service;
- privileged installer;
- automatic updater;
- production network client;
- production shell execution;
- bundled third-party DLL;
- credential store in the runtime component;
- elevated Windows service or driver;
- package manager required by the VBA runtime.

The component does call Windows **system** APIs from `kernel32` and `winmm` for
timing. It also uses Excel process-wide Application state and can dynamically
dispatch trusted VBA procedures through `Application.Run`.

The attack surface is therefore relatively small, but it is not zero.

Security-relevant or integrity-sensitive surfaces include:

```text
VBA macros
Application.Run
QueryPerformanceCounter / QueryPerformanceFrequency
GetTickCount / GetTickCount64
timeGetTime
timeGetSystemTime
timeBeginPeriod / timeEndPeriod
Application.ScreenUpdating
Application.EnableEvents
Application.DisplayAlerts
Application.Calculation
Application.Cursor
DoEvents
macro-enabled .xlsm release artifacts
release provenance tooling
GitHub Actions automation
future self-hosted Excel regression runners
```

> [!IMPORTANT]
> VBA-PERFORMANCE_MANAGER is **not a security boundary**.
>
> It does not enforce authorization, workbook permissions, segregation of duties,
> data-access controls, macro trust, process isolation, or Windows security.
>
> Its safety mechanisms — transactional timer starts, explicit read status,
> strict/non-strict policy, failure-aware measurement, balanced timer-resolution
> cleanup, and reference-counted Excel state ownership — are designed to prevent
> accidental corruption, misleading measurements, and cross-instance
> interference.
>
> They are not a sandbox against malicious VBA already running with the same
> Excel/VBA privileges.

---

## 🧭 Security model

The project assumes:

```text
Excel itself is trusted
the VBA project containing the component is trusted
the current Windows user is authorized to run that VBA project
macros are enabled through an approved trust mechanism
procedure names passed to dynamic measurement APIs come from trusted code
```

The project does **not** assume:

```text
every workbook open in the Excel process is cooperative
every Excel Application setting is local to one workbook
a returned numeric zero always means a valid zero-duration measurement
a requested timing backend is necessarily the backend eventually bound
all Windows timing branches behave identically across Office bitness
a release workbook is proven to have been built from source merely because both
  source and workbook hashes are published
a self-hosted Excel runner is safe to expose to arbitrary pull-request code
```

That distinction explains several design choices:

- `StartTimer` commits a timing session only after obtaining a valid start value;
- non-strict start fallback is reported through status and the resolved method;
- failed endpoint reads do not enter elapsed arithmetic;
- `LastReadStatus` distinguishes failure from a genuine zero result;
- the measurement harness excludes failed reads rather than storing them as
  ordinary zero-duration observations;
- statistics reject values outside the timing-sample domain;
- TW Application-state ownership is shared and reference-counted across manager
  instances;
- `Application.Calculation` baseline validity is tracked separately from its
  value;
- release provenance distinguishes source identity, execution certification,
  and workbook identity rather than treating them as the same claim.

---

## 📦 Supported versions

Security fixes are normally applied to the **latest tagged release**.

| Source state | Security support |
|---|---|
| **Latest tagged release** | ✅ Supported |
| `main` | ⚠️ Development branch / best effort |
| Older tagged releases | ❌ Normally unsupported; upgrade first |
| Modified third-party forks/copies | ❌ Unsupported unless the issue reproduces in official source |
| Unofficial binary mirrors | ❌ Unsupported |

The current latest tagged release is:

```text
v1.4.0
```

Its retained execution certification covers Microsoft 365 Excel 64-bit. Real
Office 32-bit execution remains unverified; source-level conditional branches
must not be presented as execution evidence.

When reporting, identify the exact affected state with one of:

```text
release tag
full 40-character commit SHA
```

Do not report only:

```text
latest
current
main from yesterday
```

because those descriptions can point to different source after the report is
submitted.

### Security-fix policy

A confirmed vulnerability may result in:

- a controlled fix branch;
- regression or fault-injection coverage;
- a security advisory;
- a corrected tagged release;
- updated release artifacts;
- guidance to remove or stop using an affected artifact;
- changes to release or runner permissions.

Older releases are not normally patched in place.

---

## 📣 Reporting a vulnerability

**Do not open a public GitHub issue for a suspected vulnerability.**

Do not publish:

```text
exploit code
weaponized macro-enabled workbooks
credentials or personal access tokens
private signing keys
sensitive workstation details
client workbooks
proprietary VBA
proof-of-concept runner escapes
```

in a public issue, pull request, discussion, Wiki page, or release thread.

### Option 1 — GitHub private vulnerability reporting

If private vulnerability reporting is enabled for the repository:

```text
Repository
→ Security
→ Report a vulnerability
```

submit the report there.

### Option 2 — email the maintainer

```text
danielep71@gmail.com
```

Suggested subject:

```text
Private security report — VBA-PERFORMANCE_MANAGER
```

### Include, where relevant

- affected release tag or full commit SHA;
- exact file/module/procedure;
- Excel version/build;
- Office 32-bit or 64-bit;
- Windows version;
- requested timing backend;
- resolved `ActiveMethodID`;
- `StrictMode`;
- `LastReadStatus`;
- whether TW suppression was active;
- whether `TW_CalculationExempted` was true;
- whether a second manager instance was active;
- whether the report concerns source or an official release workbook;
- minimal reproduction steps;
- observed behavior;
- expected behavior;
- confidentiality, integrity or availability impact;
- whether exploitation requires already-trusted malicious VBA;
- whether a repository workflow or runner is involved;
- any proposed mitigation;
- whether public disclosure has already occurred.

### Safe reproduction material

Prefer:

```text
sanitized workbook
minimal reproduction module
plain-text steps
regression output
small sample vectors
screenshots with sensitive values removed
hashes and exact release metadata
```

Do not attach a production workbook merely because it reproduces the problem.

---

## ⏱️ What to expect

This is a solo-maintained open-source project.

Response times are **best effort**, not a contractual SLA.

The expected process is:

1. acknowledge the report;
2. identify the affected source, artifact, and environment;
3. reproduce where practical;
4. distinguish correctness from security impact;
5. determine affected versions/artifacts;
6. develop remediation;
7. add regression/fault-injection coverage where appropriate;
8. validate on the relevant Excel/Windows host;
9. prepare a corrected release or mitigation;
10. disclose publicly after users have had reasonable time to update.

Credit can be included in release notes or an advisory if the reporter wants it.

Anonymous credit is also acceptable.

Please allow reasonable remediation time before public disclosure.

---

## 🎯 What qualifies as a security issue

When uncertain, report privately.

The maintainer can safely reclassify a report as an ordinary defect.

### 1. Unexpected code execution / macro-dispatch issues

Examples:

- official source or a release workbook executes code outside the documented
  component behavior;
- untrusted worksheet/file/network-derived input can reach `Application.Run`
  without a trusted caller deliberately selecting that behavior;
- procedure qualification unexpectedly redirects an unqualified target to
  attacker-controlled code;
- a release workbook contains undisclosed macros, external links, connections,
  embedded payloads, or executable behavior;
- a published release asset materially differs from what the release claims it
  contains.

### 2. Integrity issues

Examples:

- a native-read failure is transformed into a valid-looking timing result in a
  way that bypasses the documented status/failure model and can influence a
  security-sensitive automated decision;
- a non-owner manager instance can restore or suppress process-wide Excel state
  outside the documented ownership model;
- TW cleanup reports successful restoration while leaving security- or
  integrity-sensitive Application state in an unknown condition;
- a release manifest claims source/tag/artifact relationships it did not
  actually verify;
- a repository workflow modifies release/source metadata outside its documented
  permission boundary;
- a crafted input causes the component to operate on a different macro target
  than the documented contract permits.

### 3. Confidentiality issues

Examples:

- diagnostics or release tooling unexpectedly publishes workbook data,
  workstation paths, credentials, secrets, or client-specific information;
- a release workbook contains confidential or machine-specific information;
- a workflow artifact or log exposes a token, signing secret, environment
  variable, or private runner detail;
- future Excel-CI automation uploads an unsanitized workbook containing private
  data.

### 4. Availability issues

Examples:

- crafted input causes persistent Excel hangs, repeated crashes, or an
  uncontrolled timing loop without a practical recovery path;
- timer-resolution ownership becomes unbalanced in a way that persistently
  degrades the host environment;
- TW suppression can be driven into a state that persistently disables events,
  calculation, alerts, or screen updating outside the documented recovery path;
- a self-hosted Excel runner can be persistently compromised or left with
  orphaned Excel processes that affect later jobs.

### 5. Supply-chain issues

Examples:

- tampered GitHub Release assets;
- compromised repository workflow;
- compromised or unexpectedly changed third-party action;
- mismatch between release notes/source tag and published workbook;
- malicious redirect in download instructions;
- unauthorized replacement of a published release artifact;
- forged or misleading release provenance;
- compromise of a self-hosted Windows/Excel runner used to certify releases.

### 6. Credential / automation issues

Examples:

- repository or runner secrets exposed in logs/artifacts;
- workflow permissions broadened beyond what the job requires;
- untrusted pull-request code gaining access to a persistent self-hosted runner;
- release-signing keys becoming available to ordinary build/test jobs;
- a future PAT or token being made available to arbitrary repository scripts.

---

## 🐞 Ordinary bugs

A serious defect is not automatically a security vulnerability.

Use a public issue for ordinary defects such as:

- an elapsed value is slightly inaccurate but bounded and honestly reported;
- a timing backend has lower resolution than expected;
- a percentile or formatting routine produces the wrong ordinary numeric result;
- `Pause` overshoots within a documented bounded behavior;
- a benchmark is noisy;
- a regression case is missing;
- a source comment is stale;
- a statistic uses an undesirable numerical algorithm but no trust boundary is
  crossed;
- a TW setting is restored incorrectly in a normal reproducible scenario but
  does not create a security-relevant impact;
- a documentation example is wrong;
- a release checklist step is confusing.

Some ordinary defects can become security-relevant when they cross a concrete
confidentiality, integrity, availability, credential, or supply-chain boundary.

If unsure, report privately first.

---

## 🛠️ Scope

### In scope — production source

```text
src/classes/cPerformanceManager.cls
src/modules/M_cPM_TIMEWASTERS.bas
```

Future production modules introduced under `src/` are also in scope.

### In scope — regression and demo source

```text
test/M_cPM_Test.bas
demo/M_DEMO_BUILDER.bas
demo/M_cPM_DEMO.bas
demo/M_cPM_USAGE_EXAMPLES.bas
```

A defect in demo/test code is security-relevant when it can:

```text
misrepresent release-quality evidence
damage unrelated workbook/process state
leak sensitive data
execute unexpected code
compromise a release runner
```

### In scope — repository tooling

```text
tools/vba_lint.py
tools/release_provenance.py
```

Future committed release/build/regression tooling is also in scope.

### In scope — repository automation

Current workflow:

```text
.github/workflows/static-checks.yml
```

Future Windows/Excel execution workflows, self-hosted-runner configuration, and
release-attestation/signing automation are in scope.

### In scope — official release artifacts

Official GitHub Release assets, including:

```text
PERFORMANCE.MANAGER.xlsm
release-manifest.json
```

where published for the relevant release.

Release notes, hashes, provenance claims, and execution-certification claims are
also in scope.

### In scope — runtime integrations

- dynamic macro execution through `Application.Run`;
- Windows timing APIs;
- timer-resolution lifecycle;
- QPC/tick-count read status;
- strict/non-strict failure behavior;
- process-wide Excel Application-state suppression;
- Calculation baseline/exemption behavior;
- multi-instance TW ownership;
- emergency cleanup/recovery;
- benchmark sample validity where it affects integrity claims.

### Out of scope

- vulnerabilities in Microsoft Excel, Office, Windows, GitHub, Python, or the
  VBA runtime themselves;
- organization-controlled macro-security configuration;
- malicious VBA not supplied by this repository;
- unrelated add-ins in the host process;
- user modifications that do not reproduce in official source;
- unofficial mirrors or repackaged workbooks;
- old unsupported tags where the issue does not affect a supported state;
- social engineering unrelated to project content;
- a trusted caller deliberately passing an untrusted procedure name to
  `Application.Run` contrary to the documented trusted-input contract;
- using timing output as an authorization, authentication, randomness, or
  cryptographic primitive;
- ordinary performance or correctness bugs without concrete security impact.

A vulnerability in Excel, Windows or GitHub should be reported to the relevant
vendor/platform.

---

## 🛡️ Runtime security boundaries

### 1. Dynamic procedure execution with `Application.Run`

The measurement harness accepts a procedure name and executes it.

That is executable input.

The API is intended for **trusted procedure names supplied by trusted VBA code**.

Do not pass a procedure name directly from:

```text
worksheet cells
downloaded files
untrusted CSV/JSON
network responses
external user input
```

without an application-level allowlist or other trusted resolution policy.

An unqualified procedure name is qualified by the component to the workbook
containing the component.

An already qualified name may be honored as a deliberate caller choice.

Qualification narrows accidental resolution ambiguity.

It does **not** turn an untrusted macro name into safe data.

> [!IMPORTANT]
> `MeasureProcedure` and related dispatch helpers are benchmarking utilities, not
> a sandbox or command-authorization layer.

A report that simply demonstrates that trusted VBA can ask `Application.Run` to
execute another trusted macro is not a vulnerability.

A report showing that untrusted data reaches an unexpected macro target without
the host application deliberately allowing that behavior may be.

---

### 2. Windows timing APIs

The component uses Windows timing APIs including:

```text
QueryPerformanceCounter
QueryPerformanceFrequency
GetTickCount / GetTickCount64
timeGetTime
timeGetSystemTime
timeBeginPeriod
timeEndPeriod
```

These APIs run inside the existing Excel process and Windows user session.

They do not elevate privileges.

Security-sensitive changes include:

- changing `PtrSafe` declarations;
- changing `Long` / `LongPtr` / `Currency` representations used for native data;
- changing 32-bit / 64-bit conditional compilation;
- changing signed-to-unsigned conversion;
- changing rollover periods;
- bypassing status-bearing native-read wrappers;
- adding new system DLLs;
- adding native APIs that read/write memory, files, processes, registry, network,
  or credentials.

A native declaration defect is often a correctness or availability issue rather
than a privilege-escalation vulnerability.

The report should distinguish those categories.

---

### 3. Timer-resolution ownership

Backend 3 can request a 1 ms multimedia timer period through:

```text
timeBeginPeriod(1)
```

The request must be balanced with:

```text
timeEndPeriod(1)
```

under the component's ownership model.

Cleanup is handled explicitly through normal lifecycle methods and
`ResetEnvironment`, with `Class_Terminate` as a best-effort safety net.

A balanced timer-period request is a performance/power-management concern, not a
security boundary.

An unbounded or persistent leak that materially degrades availability may be
security-relevant.

---

### 4. Excel Application-state suppression

The TW subsystem can change process-wide Excel properties:

```text
Application.ScreenUpdating
Application.EnableEvents
Application.DisplayAlerts
Application.Calculation
Application.Cursor
```

These are **Excel-process state**, not workbook-local state.

The component therefore uses shared ownership across manager instances:

```text
capture baseline on first participating session
aggregate active requests
restore baseline on the last session
```

`Application.Calculation` is handled specially because a valid baseline may be
unavailable depending on workbook state.

The component tracks:

```text
baseline value
baseline validity
Calculation exemption
strict/non-strict policy
```

The documented stable-host invariant requires the open-workbook set to remain
stable while Calculation suppression is active.

> [!IMPORTANT]
> TW state management is an operational safety mechanism, not an authorization
> boundary.
>
> Malicious VBA already running in the Excel process can change these
> Application properties directly.

When reporting TW issues, identify:

- the initial Application state;
- which suppression flags were requested;
- how many manager instances were active;
- whether workbooks opened/closed during the scope;
- whether `TW_CalculationExempted` became true;
- whether cleanup completed normally;
- whether `PM_TW_EndAllSessions` or `ResetEnvironment` was used.

---

### 5. Read status and timing-result integrity

A numeric zero has more than one possible meaning in timing code.

It can be:

```text
a real zero-duration observation
an unresolved duration below the backend/harness floor
a non-strict failure result
a clamped invalid elapsed result
```

The project exposes status so callers do not have to infer validity from the
number alone.

Direct operations expose:

```text
LastReadStatus
```

Repeated-measurement APIs expose or are evolving toward explicit failure
evidence for worker reads.

For any security-, audit-, release-, or control-relevant use of measurements:

```text
do not discard status/failure evidence
do not treat fallback data as equivalent without checking the resolved backend
do not treat failed reads as ordinary zero observations
```

The component is intended for performance measurement.

It is **not** intended as:

```text
a cryptographic timer
a random-number source
an authentication mechanism
an authorization mechanism
an anti-tamper clock
a trusted time authority
```

---

### 6. Strict and non-strict modes

Strict mode raises when a documented invalid condition cannot be accepted.

Non-strict mode may:

```text
fall back
return a neutral numeric result
record a status
exclude invalid samples
```

depending on the operation.

Non-strict behavior is designed for resilient performance instrumentation.

It is not permission to ignore status.

When a measurement drives a release certificate or another high-integrity
decision, prefer:

```text
StrictMode = True
```

or explicitly verify every non-strict status/failure output.

---

### 7. Statistics are descriptive, not security controls

The `Stats_*` surface is designed for timing observations.

It can help identify:

```text
variance
outliers
unstable benchmark runs
undefined coefficient-of-variation conditions
```

It cannot establish that code is safe, authentic, uncompromised, or
cryptographically trustworthy.

`Stats_IsContaminated` is a benchmark-quality heuristic.

It is not an intrusion detector or security attestation.

---

### 8. Fault-injection seams

The project includes internal/Friend fault-injection hooks so deterministic
regression tests can exercise native-read failures.

These seams are for testing.

They are not authentication-protected capabilities.

Code already executing inside the same trusted VBA project can generally invoke
project-internal behavior and has the same Excel privileges as the component
itself.

A malicious macro already present in the trusted project is outside the security
boundary of these hooks.

Changes should nevertheless keep test seams:

```text
non-public where practical
one-shot or explicitly resettable
deterministic
unable to leak into later tests silently
documented as test infrastructure
```

---

## 📦 Macro-enabled release artifacts

The repository is source-first.

The authoritative implementation is the tagged exported source.

A macro-enabled workbook such as:

```text
PERFORMANCE.MANAGER.xlsm
```

is executable Office content and must be treated accordingly.

> [!WARNING]
> A familiar filename is not proof of authenticity.
>
> Confirm the release tag and published SHA-256 before treating a workbook as an
> official certified artifact.

The release workbook is a convenience distribution artifact.

Where organizational policy requires maximum transparency, import and compile
the tagged exported source in a controlled workbook rather than relying solely
on the prebuilt workbook.

---

## 🔗 Supply-chain and release integrity

### Trusted distribution

Obtain source and release assets from the official GitHub repository and Releases
page.

Do not rely on:

```text
third-party mirrors
files forwarded by email without provenance
renamed macro-enabled workbooks
unofficial package sites
blog attachments
binary copies whose hashes do not match the official release
```

### Source-first review

For controlled deployments, review the relevant tagged source:

```text
src/classes/cPerformanceManager.cls
src/modules/M_cPM_TIMEWASTERS.bas
```

and, where relevant:

```text
tools/release_provenance.py
tools/vba_lint.py
release documentation
```

before enabling macro-enabled artifacts.

### What the release manifest proves

The provenance tooling can record:

```text
tag / commit identity
source hashes
release-asset hashes
Excel build
Office bitness
regression counts
```

The tooling verifies tagged source relationships when used as documented.

A workbook digest establishes **file identity**.

It does not automatically establish that the workbook was reproducibly built
from those source files.

The VBE can reformat imported modules, and the current release process does not
automatically assemble the workbook from source.

Therefore distinguish:

```text
tagged source identity
Excel execution certification
workbook file identity
source-to-workbook build provenance
```

These are separate claims.

Do not collapse them into one.

### Signing and attestations

At the time this policy was introduced, release trust is primarily
hash/provenance based rather than a complete signed/reproducible build chain.

Future improvements may include:

```text
signed annotated tags
release attestations
VBA project/workbook signing
immutable release-asset policy
stronger source-to-artifact provenance
```

Until those controls are actually deployed, do not describe a release as signed
or reproducibly built merely because hashes exist.

---

## 🔍 Verifying a release

For controlled use:

1. obtain source/assets from the official repository;
2. record the release tag and relevant commit SHA;
3. inspect the release notes and `CHANGELOG.md`;
4. inspect the tagged source relevant to the deployment;
5. treat `.xlsm` as executable Office content;
6. compute the release workbook's SHA-256:

   ```text
   certutil -hashfile "PERFORMANCE.MANAGER.xlsm" SHA256
   ```

7. compare it with the SHA-256 published on the official Release page;
8. inspect `release-manifest.json` when the release provides one;
9. verify that manifest tag/SHA/build/bitness fields describe the intended
   release;
10. scan the workbook under organizational policy where required;
11. compile the VBA source:

   ```text
   VBA Editor → Debug → Compile VBAProject
   ```

12. run:

   ```vb
   Run_cPerformanceManager_RegressionSuite
   ```

13. record the actual:

   ```text
   commit SHA
   cases
   assertions
   failures
   Excel version/build
   Office bitness
   ```

14. require zero regression failures for release-quality evidence;
15. verify the static-check run for the same tag candidate;
16. where the deployed artifact is the release workbook, smoke-test that exact
   workbook rather than assuming source-only tests certify packaging.

A hash mismatch means the file is not the certified artifact.

Treat that as a supply-chain concern, not merely a download inconvenience.

---

## 🧰 Safe-use guidance

### 1. Preserve Excel macro security

- keep Excel macro security at the organization-approved level;
- do not disable Protected View globally;
- do not weaken Trust Center settings solely to use this component;
- use Trusted Locations or signed VBA only where organizational policy supports
  them;
- unblock downloaded macro-enabled files only after establishing provenance.

### 2. Treat dynamic procedure names as executable input

Prefer literal or code-controlled procedure names.

Do not route untrusted external values directly into:

```text
MeasureProcedure
MeasureBaseline
Application.Run
```

without an application-level allowlist.

### 3. Prefer strict measurement policy for high-integrity evidence

For release certification or other high-integrity measurements, prefer strict
mode or explicitly inspect all non-strict failure/fallback outputs.

Do not certify a benchmark from:

```text
unknown backend
unknown read status
unknown failed-read count
mixed backend samples
```

### 4. Use QPC for precision, not security

QPC is the recommended high-resolution backend.

It is still a performance timer, not a trusted security clock.

### 5. Keep an environment-recovery path

Know the distinction between:

```text
normal timing/session cleanup
TW_Turn_ON / normal scope exit
ResetEnvironment
PM_TW_EndAllSessions
Excel restart
```

A full Excel restart is the safest process-level reset after an unrecoverable VBA
project reset or uncertain global state.

### 6. Run regression tests in a controlled workbook

The current test harness manipulates real Excel state and creates a regression
worksheet.

Do not run release-validation tests in an unsaved production workbook containing
irreplaceable data.

### 7. Protect signing and environment secrets

Never store private signing keys, passwords, PATs, or client credentials inside
the workbook, VBA source, test fixtures, or release manifest.

---

## 🔑 Repository automation and runner security

The repository currently contains one software-quality workflow:

```text
.github/workflows/static-checks.yml
```

It:

- runs on hosted Ubuntu;
- checks exported source as text;
- uses repository Python tooling;
- declares `contents: read`;
- uploads a machine-readable static-check artifact;
- does not require repository secrets;
- does not execute Excel/VBA.

This is intentionally a low-privilege automation surface.

### Current workflow boundary

Static analysis can establish source consistency properties.

It cannot establish:

```text
VBE import success
VBA compile success
Excel object-model behavior
Windows API runtime behavior
Office bitness behavior
regression pass/fail
release workbook packaging
```

Do not present a green static workflow as proof that Excel executed the component.

---

## 🖥️ Future self-hosted Windows/Excel runner

A real Excel regression gate requires a Windows host with Office installed.

For a **public repository**, a persistent self-hosted runner is a materially
different security boundary from GitHub-hosted static CI.

Untrusted pull-request code must not be allowed to execute arbitrarily on a
long-lived workstation containing:

```text
Office credentials
personal files
browser sessions
signing certificates
release secrets
other repositories
network-mounted drives
developer credentials
```

### Required runner principles

A future Excel runner should follow these principles:

```text
least privilege
isolated Windows account
no unnecessary secrets
no personal data
no developer browser/session state
finite job timeout
deterministic Excel cleanup
orphan EXCEL.EXE cleanup
trusted-trigger policy
rebuildable or disposable runner where practical
separate release-signing context
machine-readable output
exact commit-SHA binding
```

### Pull requests from forks

Do not automatically execute arbitrary fork pull-request code on a persistent
self-hosted Windows/Excel runner.

Safer approaches include:

```text
manual approval before runner execution
trusted branches only
workflow_dispatch for reviewed SHAs
ephemeral disposable Windows runners
a quarantined runner with no credentials or sensitive network access
```

The exact policy should be documented before the headless Excel gate becomes a
required release control.

### Release secrets

A future test runner should not automatically receive:

```text
release-signing private keys
high-privilege GitHub PATs
unrelated repository secrets
developer credentials
```

Testing and signing are separate trust stages.

A test job compromise should not automatically become a release-signing
compromise.

---

## 🔐 Secret-handling rules

Never commit:

```text
personal access tokens
GitHub tokens
private signing keys
PFX / P12 / PVK files
passwords
API credentials
client secrets
private certificates
production connection strings
```

The repository `.gitignore` can help prevent accidental additions.

`.gitignore` is not a security control.

A secret committed once must be considered compromised even if the commit is
later deleted.

If a credential is exposed:

1. revoke or rotate it immediately;
2. determine the scope of access;
3. remove it from current source;
4. inspect workflow/release activity;
5. assess whether history cleanup is useful;
6. assume copies may remain in clones, caches, artifacts, or logs.

---

## 🧾 Logging, diagnostics, and evidence

Diagnostics should reveal enough to reproduce a problem without exposing
unnecessary user or workstation data.

Prefer recording:

```text
procedure / stage
error number
error description
timing backend
read status
counts
commit SHA
Excel build
Office bitness
sanitized runner metadata
```

Avoid publishing:

```text
worksheet data
credentials
connection strings
private environment variables
signing-key paths
user profile contents
arbitrary workbook dumps
sensitive local paths where unnecessary
```

Future machine-readable regression artifacts should contain release evidence,
not workstation secrets.

---

## 🧪 Regression harness security considerations

The regression harness manipulates real Excel process state.

It can:

```text
create/rebuild a regression worksheet
exercise all timing backends
change Application-wide TW state
request/release multimedia timer resolution
inject deterministic native-read failures
exercise cleanup and fallback behavior
run procedures through Application.Run
```

Run it in a controlled workbook/Excel instance.

A passing assertion count is not sufficient if:

```text
the runner did not finish
cleanup failed
the wrong commit was tested
the wrong workbook/source was loaded
the result was manually transcribed incorrectly
```

A future headless runner should emit its own machine-readable completion/result
state so release evidence cannot be inferred merely from partial console output.

---

## 📣 Disclosure coordination

Please avoid public disclosure while:

- exploitability is still being assessed;
- a release fix is being prepared;
- users have not had reasonable time to update;
- a credential or signing secret remains active;
- a malicious release artifact remains downloadable;
- a vulnerable self-hosted runner remains reachable.

The maintainer may ask for:

- additional environment detail;
- a sanitized reproduction;
- confirmation against a candidate fix;
- verification on a second Office bitness;
- a reasonable embargo period.

The project does not require a reporter to surrender ownership of their research.

The goal is to reduce preventable user harm.

---

## 🧭 Security review checklist for maintainers

For a security-sensitive Performance Manager change, review:

```text
[ ] Trust boundary stated
[ ] Caller-controlled input identified
[ ] Application.Run target control assessed
[ ] Requested vs resolved backend behavior assessed
[ ] Failed-read/status behavior assessed
[ ] No sentinel enters timing arithmetic
[ ] No invalid read becomes an ordinary benchmark sample
[ ] Strict/non-strict behavior documented
[ ] Win32 / Win64 declarations reviewed
[ ] timeBeginPeriod ownership remains balanced
[ ] Excel Application-state ownership assessed
[ ] Calculation baseline validity/exemption assessed
[ ] Multi-instance cleanup assessed
[ ] Original error evidence is preserved
[ ] Recovery behavior documented
[ ] Regression/fault injection added where deterministic
[ ] Manual host validation recorded where required
[ ] Source-to-artifact claims remain exact
[ ] Release-workbook impact assessed
[ ] Workflow permissions remain least-privilege
[ ] Self-hosted runner exposure assessed
[ ] Secrets/signing keys isolated from untrusted code
[ ] Documentation updated
```

A security review should distinguish:

```text
code correctness
measurement integrity
operational safety
security impact
release trust
```

They overlap.

They are not identical.

---

## 📚 Related policies and documentation

- [Project README](README.md)
- [Releasing Guide](RELEASING.md)
- [Code of Conduct](CODE_OF_CONDUCT.md)
- [Changelog](CHANGELOG.md)
- [MIT License](LICENSE)
- [Project Wiki](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/wiki)

---

## 👤 Maintainer

Maintained by **Daniele Penza**.

Private security reports:

```text
danielep71@gmail.com
```

---

<div align="center">

## 🛡️ Security principle

**Trust the source you run. Treat macro names as executable input. Keep shared Excel state explicitly owned. Separate measurement evidence from release trust. Keep untrusted code away from privileged runners and signing material.**

</div>
