# Releasing

The v1.2.0 release went wrong three times before it was correct, and every one of
those failures was mechanics rather than code: a version stamp that drifted, a
tag created against an unmerged `main`, and a tag named `V1.2.0` where the two
existing tags were lowercase.

This file exists so those are checked rather than remembered.

---

## Before you start

- [ ] Every issue in the milestone is closed
- [ ] `main` is green — the **VBA source consistency** check passes
- [ ] Local `main` is up to date. **Pull, do not assume**

> [!WARNING]
> A stale local `main` caused two of the three v1.2.0 failures. GitHub Desktop
> reports "already up to date" against your *local* copy, not the remote.

---

## 1. Sync the version stamps

Version stamps move **once**, here. They stay untouched during development so a
half-released state is impossible.

Set the same version in all four:

| File | Where |
|---|---|
| `src/classes/cPerformanceManager.cls` | `' VERSION: x.y.z` in the header |
| `src/modules/M_cPM_TIMEWASTERS.bas` | `' VERSION` block |
| `test/M_cPM_Test.bas` | `' VERSION` block |
| `README.md` | the version badge |

The linter's **version stamps agree** check will fail if any of them disagree,
so push and let it confirm rather than checking by eye.

---

## 2. Sync the documentation

- [ ] README case and assertion counts match the suite
- [ ] README public-member count matches the class
- [ ] Any guarantee that changed is corrected in README **and** the wiki
- [ ] `CHANGELOG.md`: rename `[Unreleased]` to the version with today's date
- [ ] `CHANGELOG.md`: add a fresh empty `[Unreleased]` section above it
- [ ] `CHANGELOG.md`: add the link reference at the foot of the file

---

## 3. Run and record the suite

Run `Run_cPerformanceManager_RegressionSuite` against **the exact code you are
about to tag**, not an earlier build.

Record from the summary and from **File → Account → About Excel**:

| | |
|---|---|
| Cases | |
| Assertions | |
| Failures | must be **0** |
| Excel version and build | |
| Bitness | |

> [!NOTE]
> A run certifies one bitness. Say which was tested rather than implying both.

### For a major release, certify both bitnesses

- [ ] Suite run and recorded on **64-bit** Office
- [ ] Suite run and recorded on **32-bit** Office

Office bitnesses cannot coexist on one machine, so this needs a VM or a second
host. It is not required for a minor release: the arithmetic shared between the
two paths is verified on whichever bitness runs, and only the declaration
binding, the API return and the `RolloverSeconds` constant are bitness-specific.

Where only one bitness was run, say so in the provenance block rather than
leaving the reader to assume both.

---

## 4. Generate the provenance block

```bash
python3 tools/release_provenance.py --version 1.3.0 \
    --asset "demo/PERFORMANCE MANAGER.xlsm" \
    --excel "Microsoft 365 MSO, Version 2606, Build 16.0.20131.20152" \
    --bitness 64-bit --cases 69 --assertions 447
```

The script exits non-zero and marks any missing field as `TODO`, so an
incomplete block cannot slip through unnoticed. Keep the output for step 6.

> [!IMPORTANT]
> Run this **after** the final commit. It hashes the working tree and reports
> the current commit, so running it earlier publishes hashes that do not match
> the tag.

---

## 5. Tag

Push everything first, then:

- [ ] **Confirm the Static checks run for this exact commit is green.**
      Actions → Static checks → the run for your tag target. Download the
      `vba-lint-results` artifact and check `"passed": true` and the `commit`
      field matches.

      > Branch protection does not require this check, so a commit can reach
      > `main` without it having run. That trade is deliberate for daily work
      > and wrong at release time, which is why this step is manual and explicit.

- [ ] **Releases → Draft a new release**
- [ ] Tag: `vX.Y.Z` — **lowercase `v`**, matching `v1.0.0`, `v1.1.0`, `v1.2.0`
- [ ] Target: `main`
- [ ] **Check the commit shown next to the target.** It must be the merge you
      intend to release

> [!CAUTION]
> Git tags are case sensitive. `V1.2.0` and `v1.2.0` are different refs, and the
> CHANGELOG links point at the lowercase form.
>
> If you use a pull request, **merge it before tagging**. Tagging `main` while
> the release branch is unmerged captures the pre-release state — this happened
> in v1.2.0 and shipped a tag containing none of the release's work.

---

## 6. Publish

- [ ] Title: `vX.Y.Z — <one-line theme>`
- [ ] Body: the release notes, with the provenance block from step 4 appended
- [ ] **Attach `PERFORMANCE MANAGER.xlsm`**
- [ ] Publish

> [!WARNING]
> The README, the wiki and the release notes all tell readers the demo workbook
> comes from Releases. Publishing without the attachment breaks that promise in
> three places at once.

---

## 7. Verify the published tag

Do not skip this. It is the check that failed in v1.2.0.

```bash
git fetch --tags --force
git log vX.Y.Z --oneline -1          # is this the commit you meant?
git show vX.Y.Z:test/M_cPM_Test.bas | grep TotalSteps
git show vX.Y.Z:src/classes/cPerformanceManager.cls | grep "VERSION:"
```

The tag should resolve to your merge commit, and the tagged tree should carry
the version and case count you just recorded. If it does not, delete the tag and
the release and start again from step 5.

---

## 8. Close out

- [ ] Record the Static checks run URL alongside the provenance block
- [ ] Close the milestone
- [ ] Delete any merged release branch
- [ ] Confirm the CHANGELOG release link resolves
- [ ] Confirm the download link on the README reaches the attached asset

---

## What this does not cover

The regression suite is run and recorded by hand. Automating it needs a
self-hosted Windows runner with Excel installed, which is tracked separately.

Until then, the suite result in the pull request template and the certification
block in the release **are** the record that the code was run. They are only
worth anything if they describe the build actually being tagged, which is why
steps 3 and 4 come after the final commit and step 7 verifies the result.
