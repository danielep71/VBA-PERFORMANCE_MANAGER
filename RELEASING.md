<div align="center">

# 🚀 Release Guide

### Canonical maintainer procedure for publishing VBA-PERFORMANCE_MANAGER

[![Scope](https://img.shields.io/badge/Scope-Maintainer_release_control-6f42c1?style=flat-square)](#-release-control-model)
[![Source](https://img.shields.io/badge/Source-Tag_is_authoritative-217346?style=flat-square)](#-release-control-model)
[![Static](https://img.shields.io/badge/Static-GitHub_Actions-0969da?style=flat-square)](#-final-static-source-gate)
[![Excel](https://img.shields.io/badge/Excel-Execution_evidence-d97706?style=flat-square)](#-final-excel-execution-certification)
[![Provenance](https://img.shields.io/badge/Provenance-SHA--256_%2B_tag_binding-0f766e?style=flat-square)](#-generate-release-provenance)

<br>

**Exact tag target · Clean source · Frozen history · Real Excel evidence · Explicit bitness · Verified artifacts · No silent replacement**

<br>

[Readiness](#-release-readiness)
&nbsp;·&nbsp;
[Merge](#-merge-the-release-branch-to-main)
&nbsp;·&nbsp;
[Static gate](#-final-static-source-gate)
&nbsp;·&nbsp;
[Excel gate](#-final-excel-execution-certification)
&nbsp;·&nbsp;
[Artifact](#-build-and-validate-the-release-workbook)
&nbsp;·&nbsp;
[Tag](#-tag-the-exact-release-commit)
&nbsp;·&nbsp;
[Provenance](#-generate-release-provenance)
&nbsp;·&nbsp;
[Publish](#-publish-the-github-release)
&nbsp;·&nbsp;
[Verify](#-post-publication-verification)

</div>

---

> [!IMPORTANT]
> This document is for **maintainers publishing an official release**.
>
> It is not an installation guide.
>
> Users and integrators should follow:
>
> [INSTALLATION.md](INSTALLATION.md)

> [!IMPORTANT]
> **Tagged exported source is authoritative.**
>
> The macro-enabled workbook is a convenience Release asset. Its digest proves
> file identity, not an automated or reproducible source-to-workbook build.

> [!WARNING]
> The final release certification must describe the **exact commit that will be
> tagged**.
>
> Testing a release branch and then creating a different merge commit on `main`
> is useful pre-merge validation, but it is not SHA-bound certification of the
> final tag target.
>
> Final static and Excel evidence therefore come **after the release branch is
> merged to `main` and the final release SHA is known**.

---

# 🧭 Release control model

A valid release is a chain of separate controls:

```text
milestone / scope complete
        ↓
release branch prepared
        ↓
pre-merge source checks
        ↓
release branch merged to main
        ↓
final tag-target SHA recorded
        ↓
static source checks green on that exact SHA
        ↓
real Excel compile / regression evidence on that exact SHA
        ↓
release workbook assembled from that exact source
        ↓
actual final workbook smoke-tested
        ↓
lowercase vX.Y.Z tag created on that exact SHA
        ↓
source/tag provenance verified
        ↓
Release assets hashed
        ↓
GitHub Release published
        ↓
published tag and downloads independently re-verified
```

Do not collapse those into one claim.

In particular:

```text
green static checks
≠
Excel execution

release-branch regression
≠
final tagged-commit certification

Excel source regression
≠
release workbook packaging validation

workbook SHA-256
≠
reproducible build

source-compatible API
≠
unchanged physical import package
```

---

## Current repository control boundary

The repository currently has:

```text
.github/workflows/static-checks.yml
tools/vba_lint.py
tools/release_provenance.py
```

The hosted workflow runs static analysis on GitHub-hosted Ubuntu with:

```text
contents: read
```

It does **not** execute Excel/VBA.

The live `protect-main` ruleset currently protects the default branch against:

```text
deletion
non-fast-forward updates
```

It does **not** currently enforce:

```text
a pull request
VBA source consistency as a required status check
an up-to-date branch before merge
```

This release guide is intentionally stricter than the live branch rules.

Release policy requires the maintainer to verify the final candidate even where
GitHub does not technically block an unsafe merge.

---

# ✅ Release readiness

Before changing version stamps or creating a tag:

```text
[ ] Target milestone is the intended release
[ ] Every release-blocking issue is closed
[ ] No open P1 issue affects the release
[ ] Deferred work is explicitly moved to a later milestone
[ ] Release branch contains only intended release work
[ ] Working tree is clean
[ ] Local refs have been fetched from origin
[ ] Release support policy is explicit
[ ] Installation/deployment changes are understood
[ ] Known limitations are documented honestly
```

### Release branch

Use a dedicated branch where practical:

```text
release/vX.Y.Z
```

For example:

```text
release/v1.4.0
```

The release branch is preparation space.

The official release tag should resolve to the final commit on `main` after the
release branch has been merged.

> [!CAUTION]
> Do not tag `main` while the release branch is still unmerged.
>
> That can produce a tag containing the pre-release state even when the release
> branch itself is correct.

---

# 1. Fetch and synchronize

Before release preparation:

```text
GitHub Desktop → Fetch origin
```

or:

```bash
git fetch --all --tags --prune
```

Confirm:

```bash
git status
git branch --show-current
git log --oneline -5
```

The working tree should be clean.

Do not rely on an old local "up to date" indication without fetching origin
first.

---

# 2. Freeze release scope

Review the milestone one final time.

For every open item, choose one of:

```text
complete it
move it to a later milestone
explicitly remove it from release scope
```

Do not publish with an issue silently half-in / half-out of the milestone.

### Compatibility decisions

Before release, explicitly settle any change affecting:

```text
public method signatures
public enum/error values
minimum Office/VBA support
required source-file/import package
Office 32-bit support
reporting-module availability
release artifact format
```

If the physical source package changes, update `INSTALLATION.md` in the same
release.

---

# 3. Sync version stamps

Move release version stamps **once**, after feature work is complete.

For the current source layout, synchronize:

| File | Release field |
|---|---|
| `src/classes/cPerformanceManager.cls` | `' VERSION: X.Y.Z` |
| `src/modules/M_cPM_TIMEWASTERS.bas` | `' VERSION` block |
| `test/M_cPM_Test.bas` | `' VERSION` block |
| `README.md` | version badge / current-version references |

Run the linter after the update.

The static `version stamps agree` check exists so this is verified mechanically,
not by visual inspection alone.

### Future module splits

If a later release adds required production modules, those modules must join the
version/source inventory.

The generated source inventory planned for the repository should become the
authoritative list rather than maintaining the count manually forever.

---

# 4. Prepare `CHANGELOG.md`

The project follows Keep a Changelog / Semantic Versioning.

For the new release:

1. convert the current `[Unreleased]` section into `[X.Y.Z] — YYYY-MM-DD`;
2. add a fresh empty `[Unreleased]` section above it;
3. add/update the comparison link references at the foot of the file;
4. ensure every material user-visible change is represented;
5. ensure breaking/behavioral changes are clearly called out.

> [!IMPORTANT]
> **Do not rewrite previously released CHANGELOG sections.**
>
> The static linter verifies released sections against their tagged history.
>
> If an old release note contained an overclaim or wording error, correct the
> current contract in `[Unreleased]`, current documentation, or an explicit
> erratum. Do not make the historical tag appear to have said something it did
> not say.

---

# 5. Sync current documentation

Review all current-contract documents:

```text
README.md
INSTALLATION.md
SECURITY.md
CODE_OF_CONDUCT.md
CHANGELOG.md
Wiki pages
issue / pull request templates where relevant
```

Not every release requires edits to every file.

Update a file only when its contract changed.

At minimum verify:

```text
[ ] README version/current capabilities are accurate
[ ] INSTALLATION required source files and import order are accurate
[ ] README and INSTALLATION do not disagree
[ ] SECURITY trust/runner/release claims remain accurate
[ ] Wiki matches current supported behavior
[ ] Known limitations are not stronger than the implementation
[ ] Regression counts are not stale where intentionally displayed
[ ] Static-check count is not stale where intentionally displayed
[ ] DLL wording distinguishes Windows system APIs from bundled/third-party DLLs
```

### Repository controls are documentation too

If branch rules, Discussions, workflows, or release settings changed, verify that
documentation and issue-template links describe the **live** configuration.

---

# 6. Run pre-merge static checks

From the release branch, run:

```bash
python tools/vba_lint.py --json vba-lint-results.json
```

Required result:

```text
all checks pass
```

Push the prepared release branch and confirm:

```text
Actions
→ Static checks
→ run for the release-branch candidate
```

This is a **pre-merge quality gate**.

It catches problems before the release branch reaches `main`.

It is not yet the final release certification because the final `main` SHA may
be different after merge.

### Optional pre-merge Excel run

A full Excel regression on the release branch is useful when the release is
large or risky.

Treat it as pre-merge validation.

The final release evidence must still be tied to the exact tag-target SHA after
merge.

---

# 7. Review and merge the release branch

Recommended workflow:

```text
release/vX.Y.Z
        ↓
review / pull request
        ↓
main
```

A pull request is recommended even though the current repository rules do not
require one.

Before merge:

```text
[ ] release diff contains only intended changes
[ ] static checks are green
[ ] no unresolved review point remains
[ ] milestone scope still matches the release
```

Merge the release branch.

---

# 🔀 Merge the release branch to `main`

After merge:

```text
GitHub Desktop → Current branch → main
GitHub Desktop → Fetch origin
GitHub Desktop → Pull origin
```

or:

```bash
git fetch origin
git checkout main
git pull --ff-only
```

Confirm:

```bash
git status
git log --oneline -5
```

Require:

```text
working tree clean
all release changes present
HEAD = intended release commit
```

Do not tag yet.

---

# 8. Record the exact release SHA

Capture:

```bash
git rev-parse HEAD
```

Save the full 40-character SHA in the release working notes.

This is now the **tag target**.

Everything that follows must refer to this commit.

If any source or documentation change is made after this point:

```text
the release SHA changed
```

Repeat every affected final gate.

---

# 🔎 Final static source gate

The merge to `main` should trigger the hosted static workflow.

Confirm:

```text
Actions
→ Static checks
→ run whose commit = exact tag-target SHA
```

The run must be green.

Download/inspect the `vba-lint-results` artifact.

Verify:

```text
"passed": true
commit field = exact tag-target SHA
```

> [!IMPORTANT]
> The live branch rules do not currently require this status check.
>
> **Release policy does.**

The implementation currently runs **12** static checks, covering source
consistency including:

```text
conflict-marker absence
procedure/block balance
reserved-word identifier checks
error-source consistency
no bare production error offsets
local-callee consistency
test wiring
TotalSteps consistency
version-stamp agreement
native API call-site invariants
released CHANGELOG immutability
```

A green static run does **not** establish:

```text
VBE import success
VBA compile success
Excel object-model behavior
Windows API runtime behavior
Office bitness behavior
regression pass/fail
release-workbook packaging
```

Those require real Excel.

---

# 🧪 Final Excel execution certification

The final Excel certification must execute the source from the exact tag-target
SHA recorded above.

## Current process — until the headless Excel gate is implemented

Use a controlled validation workbook/project.

Import the files from the final `main` checkout:

```text
src/modules/M_cPM_TIMEWASTERS.bas
src/classes/cPerformanceManager.cls
demo/M_DEMO_BUILDER.bas
test/M_cPM_Test.bas
```

Then:

```text
Debug → Compile VBAProject
```

Run:

```vb
Run_cPerformanceManager_RegressionSuite
```

Record:

| Evidence | Required |
|---|---|
| Exact tag-target SHA | ✅ |
| Excel version/build | ✅ |
| Office bitness | ✅ |
| Cases | ✅ |
| Assertions | ✅ |
| Failures | **must be 0** |
| Run date/time | ✅ |
| Cleanup/result completion | ✅ |

Do not certify from:

```text
an earlier release-branch run
an old validation workbook
a different commit
a partial run
a run whose final cleanup/result is uncertain
```

### One bitness run certifies one bitness

A 64-bit run does not execute the 32-bit conditional branches.

A 32-bit run does not execute the Win64 branches.

A release may claim execution certification only for environments actually run.

### Bitness evidence

Do not treat any of these as execution certification for a bitness you did not
run:

```text
compiled by construction
shared arithmetic is tested
the other bitness passed
```

They are useful source-level facts and nothing more.

Where a bitness cannot be provisioned, say so explicitly before release: name
the environments actually certified, state that the other remains unverified
rather than unsupported, and record the tracking issue. Deferring evidence is
legitimate; implying it exists is not.

> [!NOTE]
> **v1.4.0 decision, 2026-08-31.** 32-bit Office could not be provisioned, since
> it cannot coexist with 64-bit Office on one Windows installation. v1.4.0
> shipped certified on 64-bit against its exact tag target, with 32-bit recorded
> as unverified and tracked in [#29](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/29).

---

## Future process — after the headless Excel gate lands

Once the executable gate is implemented, the workflow artifact for the **exact
tag-target SHA** becomes the primary Excel execution record.

It should include at least:

```text
commit SHA
explicit run/completion state
Excel version/build
Office bitness
cases
assertions
failures
cleanup outcome
timestamp
```

A persistent self-hosted runner must follow the trust model documented in
`SECURITY.md`; arbitrary unreviewed fork code must not execute automatically on
a privileged long-lived Excel host.

Manual Excel runs may remain useful as supplemental validation.

They should not replace or contradict the automated SHA-bound record once that
gate becomes release policy.

---

# 📦 Build and validate the release workbook

The Release workbook is a convenience executable artifact.

It is not generated automatically today.

## Assemble from the exact tag-target source

Use the final `main` source at the recorded release SHA.

Do not start from an old workbook and assume its code matches.

Where practical:

1. use a controlled workbook;
2. remove stale component copies;
3. import the final production source from the tag-target checkout;
4. import demo/test source required by the release workbook;
5. compile;
6. run the applicable regression/demo validation;
7. save the workbook **outside the Git-tracked source tree**.

Use the release's intended filename.

The convention, used by v1.3.0 and v1.4.0, is:

```text
PERFORMANCE.MANAGER.xlsm
```

> [!NOTE]
> The local build is named `PERFORMANCE MANAGER.xlsm` with a space. GitHub
> replaces the space with a dot on upload, so the published asset carries the
> dotted name. Record the published name.

If a future release changes the asset name, the Release page and documentation
must use the exact new name consistently.

## Package-level smoke test

After the final save:

1. close the workbook;
2. reopen the **actual file that will be uploaded**;
3. satisfy macro/trust policy;
4. run a basic timing smoke test;
5. run the relevant demo/startup checks;
6. confirm the workbook contains the intended version;
7. confirm no unintended external links, credentials, or local-only data were
   introduced.

> [!IMPORTANT]
> If the workbook is modified after validation, it is a new artifact.
>
> Re-run the relevant package-level validation and recompute its hash.

Source regression and package smoke are separate evidence.

---

# 🏷️ Tag the exact release commit

Only after the final static and Excel gates pass, create:

```text
vX.Y.Z
```

Use a lowercase `v`, consistent with the existing release series.

Examples:

```text
v1.3.0
v1.4.0
v1.4.1
v1.5.0
```

Git tags are case-sensitive.

```text
v1.4.0
≠
V1.4.0
```

Create the tag on the exact recorded `main` SHA.

Use an annotated tag where the chosen workflow supports it.

Do not claim cryptographic signing unless the tag/attestation is actually
verified as signed under the implemented release-trust model.

Push the tag:

```bash
git push origin vX.Y.Z
```

Verify immediately:

```bash
git rev-parse vX.Y.Z
git log vX.Y.Z --oneline -1
```

The resolved SHA must equal the release SHA recorded before the final gates.

---

# 🔐 Generate release provenance

Run provenance **after the tag exists** and while `HEAD` is still the exact
tagged commit.

The tool verifies the local source against the tag and hashes the Release asset.

Windows Command Prompt example:

```bat
python tools\release_provenance.py ^
  --version X.Y.Z ^
  --tag vX.Y.Z ^
  --asset "C:\path\to\PERFORMANCE.MANAGER.xlsm" ^
  --out release-manifest.json ^
  --excel "Microsoft 365 MSO, Version ####, Build ##.##.#####.#####" ^
  --bitness 64-bit ^
  --cases N ^
  --assertions N ^
  --failures 0
```

A one-line command is equally valid.

### Required checks

The command must complete without source/tag verification problems.

Inspect both console output and `release-manifest.json`.

Require:

```text
[ ] commit = tagged release commit
[ ] tag_verified = vX.Y.Z
[ ] no TODO required fields
[ ] no tag problems
[ ] source hashes present
[ ] workbook asset hash present
[ ] Excel build explicit
[ ] bitness explicit
[ ] cases/assertions explicit
[ ] failures = 0
```

### Current manifest bitness boundary

`release_provenance.py` currently records **one certification bitness per
manifest invocation**.

If a release has separate 32-bit and 64-bit execution evidence, do not squeeze
both into one `--bitness` field.

Until the provenance schema is extended, retain the second bitness evidence as a
separate machine-readable execution artifact / certification record and identify
both runs explicitly in the Release notes.

The manifest proves exactly what its schema records — no more.

---

## What provenance establishes

| Claim | Evidence |
|---|---|
| Source files match the tag | `--tag` source/blob comparison |
| Downloaded workbook is the published file | workbook SHA-256 |
| Exact repository commit | manifest commit SHA |
| Excel environment reported by releaser | certification fields |
| Real suite result | releaser/executable-gate evidence |

It does **not** establish:

```text
deterministic workbook build
byte-for-byte source round-trip through the VBE
cryptographic tag authenticity when the tag is unsigned
automatic Office execution when the run was manual
```

Tagged exported source remains authoritative.

---

# 📝 Draft the GitHub Release

Use:

```text
GitHub → Releases → Draft a new release
```

Select the **existing** tag:

```text
vX.Y.Z
```

Do not create a second tag from the Release form.

Confirm the commit shown for the tag is the exact release SHA.

### Recommended title

```text
vX.Y.Z — <one-line release theme>
```

### Release body

Include:

```text
short release summary
material fixes/features
behavioral/breaking changes
certification table
Office bitness actually executed
known limitations
installation/upgrade changes
provenance block / manifest reference
links to CHANGELOG / INSTALLATION where useful
```

Avoid claims such as:

```text
fully tested on 32/64-bit
reproducibly built
signed
immutable
```

unless the evidence for that exact release establishes them.

---

#  Attach release assets

At minimum, where applicable:

```text
PERFORMANCE.MANAGER.xlsm
release-manifest.json
```

Also attach the machine-readable Excel execution artifact once the headless gate
exists, and any separate 32-bit certification artifact required by the release.

Before publication:

```text
[ ] asset filenames are correct
[ ] asset hashes match the provenance evidence
[ ] no temporary/local file was attached accidentally
[ ] workbook opened successfully after final save
[ ] release-manifest.json is from the final tagged commit
```

---

# 📣 Publish the GitHub Release

Publish only after all prior gates are complete.

> [!WARNING]
> **Do not silently replace a published asset.**
>
> GitHub Releases are currently technically mutable, but release policy should
> treat published assets as immutable evidence.
>
> If a published asset is wrong, prefer an explicit correction / patch release
> rather than replacing the file under the same version without disclosure.

---

# 🔎 Post-publication verification

Do not stop at the green **Published** banner.

## 1. Verify the tag again

```bash
git fetch --tags --force
git log vX.Y.Z --oneline -1
git show vX.Y.Z:src/classes/cPerformanceManager.cls | findstr "VERSION:"
git show vX.Y.Z:test/M_cPM_Test.bas | findstr TotalSteps
```

Confirm the tag still resolves to the intended release commit.

---

## 2. Verify GitHub Release metadata

Check:

```text
[ ] Release tag = vX.Y.Z
[ ] Release is not accidentally marked prerelease
[ ] Title/version is correct
[ ] Release notes match the intended release
[ ] Uploaded assets are present
[ ] Manifest is present
```

---

## 3. Verify the downloaded workbook

Download the workbook **from the published Release**, not from the staging
folder.

Compute SHA-256.

PowerShell:

```powershell
Get-FileHash -Algorithm SHA256 ".\PERFORMANCE.MANAGER.xlsm"
```

or:

```text
certutil -hashfile "PERFORMANCE.MANAGER.xlsm" SHA256
```

Compare with the published provenance/manifest.

A mismatch means the downloaded file is not the certified asset.

---

## 4. Smoke-test the downloaded artifact

Open the actual downloaded release workbook in a controlled Excel instance.

Confirm:

```text
[ ] macro/trust policy satisfied
[ ] workbook opens normally
[ ] version shown is correct
[ ] basic QPC timing works
[ ] demo/basic workflow works
[ ] no packaging-only error appears
```

---

## 5. Verify links

Check:

```text
README release/download links
CHANGELOG comparison/release link
INSTALLATION.md links
Wiki release/download references
SECURITY.md references where changed
```

A technically correct release with broken download/documentation links is still
an incomplete publication.

---

# 🧹 Close out the release

After post-publication verification:

```text
[ ] Record static-check run URL / identifier
[ ] Record Excel execution evidence / workflow run
[ ] Record Office bitness(es) actually certified
[ ] Record release manifest/hash evidence
[ ] Close the milestone
[ ] Delete the merged release branch when no longer needed
[ ] Confirm CHANGELOG release link resolves
[ ] Confirm README points to the intended current release
[ ] Confirm INSTALLATION reflects the published source package
[ ] Confirm no temporary release files are tracked
```

Do not delete the release branch until the published tag and assets have been
verified.

---

# 🧯 Release recovery

## Problem found before publication

If the release is still a draft:

```text
fix the source/docs/artifact
commit the correction
rerun affected final checks on the new SHA
recreate/move the unpublished tag if necessary
regenerate provenance
```

Do not publish evidence from the superseded candidate.

---

## Wrong tag found before publication

Delete/recreate the unpublished tag so it resolves to the intended commit.

Then rerun provenance with `--tag`.

---

## Problem found after publication

Do not silently move the tag or replace an asset under the same version.

Choose a transparent recovery appropriate to severity:

```text
publish a patch release
mark the affected release clearly
publish a security advisory if applicable
withdraw the release only with an explicit public record where necessary
```

If a published release must be deleted because it was created in error, record
what happened before removing the evidence.

Release trust is damaged more by silent history rewriting than by an openly
documented correction.

---

# 📋 Release evidence record

For every release, retain a compact record containing:

```text
version
tag
full commit SHA
release date
final static-check run
static-check artifact/result
Excel version/build
Office bitness
regression cases
regression assertions
regression failures
Excel execution run/artifact when automated
workbook filename
workbook SHA-256
release-manifest filename/hash where relevant
known certification gaps
```

This can live in:

```text
GitHub Release notes
release-manifest.json
workflow artifacts
```

The repository source does not need to duplicate every transient artifact.

---

# ✅ Final release checklist

```text
SCOPE
[ ] Milestone scope finalized
[ ] No release-blocking issue remains
[ ] Support / bitness policy explicit

RELEASE BRANCH
[ ] Working tree clean
[ ] Branch synced with origin
[ ] Version stamps synchronized
[ ] CHANGELOG current release prepared
[ ] Released CHANGELOG history untouched
[ ] README / INSTALLATION / Wiki consistent
[ ] Pre-merge vba_lint passes
[ ] Release-branch Static checks green

MAIN / FINAL SHA
[ ] Release branch merged to main
[ ] Local main fetched/pulled and clean
[ ] Full tag-target SHA recorded
[ ] Final Static checks green on that exact SHA
[ ] vba-lint-results commit matches tag-target SHA

EXCEL
[ ] Exact tag-target source imported
[ ] Debug → Compile VBAProject passes
[ ] Regression suite completes
[ ] Failures = 0
[ ] Cases / assertions recorded
[ ] Excel version/build recorded
[ ] Office bitness recorded
[ ] 32-bit evidence present where required
[ ] Cleanup/completion state valid

ARTIFACT
[ ] Final PERFORMANCE.MANAGER.xlsm assembled from tag-target source
[ ] Actual final-saved workbook reopened
[ ] Package smoke passes
[ ] No post-validation binary edits

TAG
[ ] Lowercase vX.Y.Z tag created on exact tag-target SHA
[ ] Tag pushed
[ ] Tag independently re-resolves to recorded SHA

PROVENANCE
[ ] release_provenance.py run after tag exists
[ ] Source matches tag
[ ] No required TODO fields
[ ] Workbook SHA-256 recorded
[ ] release-manifest.json generated from final tag/artifact
[ ] Additional bitness evidence retained separately where required

PUBLICATION
[ ] Existing tag selected in GitHub Release
[ ] Release notes accurate
[ ] Certification claims limited to actual evidence
[ ] Workbook attached
[ ] Manifest attached
[ ] Additional execution evidence attached where applicable

POST-PUBLISH
[ ] Published tag independently verified
[ ] Downloaded workbook hash matches
[ ] Downloaded workbook smoke-tested
[ ] README / CHANGELOG / INSTALLATION links checked
[ ] Milestone closed
[ ] Release branch removed only after verification
```

---

<div align="center">

## 📌 Release principle

**Merge first, certify the exact tag target, test the artifact you publish, and claim only what the evidence proves.**

</div>
