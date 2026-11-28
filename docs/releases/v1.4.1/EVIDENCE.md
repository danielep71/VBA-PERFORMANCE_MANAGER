# v1.4.1 release evidence

This is the certification and release record of v1.4.1, published on
28 November 2026 from `main` at `189668851754ec10e4c86d02f2e12b16739a54a5`,
the target of tag `v1.4.1`. It was completed after publication, in the first
pull request of the next cycle (RELEASING step 12). v1.4.0's evidence is
recorded in [docs/releases/v1.4.0](../v1.4.0/IMPLEMENTATION_AND_RELEASE_PLAN.md).

## Where the work is tracked

| What | Where |
|---|---|
| Open and closed work | The [v1.4.1 milestone](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/milestone/4) |
| Final certification and publication | [#101](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/101) |
| Evidence for each change | The closing comment of its issue and the pull request that delivered it |
| Branch and merge rules | [CONTRIBUTING.md](../../../CONTRIBUTING.md#maintainer-workflow) |
| Release procedure | [RELEASING.md](../../../RELEASING.md) |
| Preparation notes for individual changes | [#74 Pause method 3](PAUSE_METHOD3.md), [#75 worker seam cleanup](WORKER_SEAM_CLEANUP.md), [#76 strict error codes](STRICT_ERROR_ASSERTIONS.md), [#77 measurement examples](MEASUREMENT_EXAMPLE_EVIDENCE.md), [#47 demo workbook inventory](DEMO_WORKBOOK_INVENTORY.md) |

The implementation plan that tracked v1.4.1 until 3 November 2026 was retired
under [#90](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/90).
Its open work items all have issues. Its full text, Revisions 4–31, stays
readable at its
[last version](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/blob/52717f836226faea123d9405ee23b76d752fae88/docs/V1.4.1_IMPLEMENTATION_PLAN.md).

## Integration into `main`

Each increment reached `main` through a pull request: increment 1 as a merge
commit, later ones squash-merged from `release/v1.4.1`. None of them carried an
executable VBA change without an Excel run.

| Increment | Pull request | `main` commit | Contents |
|---|---|---|---|
| 1 | [#67](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/67) | `bd094704` | Assurance tooling, documentation reconciliation and immutable workflow pins (#36, #45, #46, #51, #63, #64); the Apache-2.0 relicensing merged alongside as [#68](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/68) |
| 2 | [#79](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/79) | `b32d7af7` | Pin-scanner and provenance fixes, the Release package, the comment audit (#69–#72) |
| 3 | [#100](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/100) | `cac3a4c6` | Template-parity tooling and governance (#80–#86, #89) |
| 4 | [#106](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/106) | `19a8a208` | #33's sticky Calculation exemption, the local Excel runner (#11 part 1), planning in issues (#90), the settings-script fix (#103) |
| 5 | [#114](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/114) | `7c38d5ad` | The local Excel runner releases every COM reference before Excel quits and reports its cleanup (#11); no VBA change |
| 6 | [#117](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/117) | `4aa9b366` | #33's lifecycle coverage (cases 82–85) and the double-click certification launcher (#115, #116). Excel at `2dd25f1`: 85 cases, 696 assertions, 0 failures |
| 7 | [#118](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/118) | `107454a2` | #31's timer-period rollback (case 86). Excel: 86 cases, 714 assertions, 0 failures, with a negative control |
| 8 | [#121](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/121) | `931aee4d` | A red check passes only when Excel cleanup completed, with runner verdict fixtures (#112); no VBA change |
| 9 | [#123](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/123) | `e0a95c15` | Demo sheets skip the printer-dependent pagination reset (#119). Excel: 86 cases, 714 assertions, 0 failures |
| 10 | [#127](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/127) | `1230b821` | Continued comments and inline `Rem` in the source rules (#111), issue planning line endings (#110), milestone cleanup for PRs in every state (#109); no VBA change |
| 11 | [#130](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/130) | `30ff335e` | Pinned source/wiki release-state checks (#50) and inline code in Markdown heading anchors (#108); no VBA change |
| 12 | [#133](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/133) | `f430f0ae` | The Office/VBA support policy (#34) and coverage of Pause method 3's coarse wait (#74). Excel at `ca66242`: 86 cases, 715 assertions, 0 failures |
| 13 | [#136](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/136) | `eb455d7d` | Worker fault seams disarmed on validation errors (#75) and strict tests requiring documented error codes (#76), case 87. Excel: 87 cases, 771 assertions, 0 failures |
| 14 | [#141](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/141) | `471b9021` | CodeQL `upload-sarif` 4.39.3 (#139), quality-tool pins (#140) and the Stars badge; no VBA change |
| 15 | [#148](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/148) | `beb7e69c` | #44, #30, #35, #73 and #32 (cases 88–93). Excel at `3d61981`, tree-identical: 93 cases, 861 assertions, 0 failures; each fix also passed its own red check |
| 16 | [#151](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/151) | `8e1f0125` | #77's measurement examples, #138's local workbook tools, #11's batch runner and wiki pin `66527bb` (#150). VBA identical to `307c703`: 93 cases, 861 assertions, 0 failures, four examples passed |
| 17 | [#153](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/153) | `ac139907` | Wiki pin `36accf1` (#152); no VBA change |
| 18 | [#155](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/155) | `43bc71ca` | This record's rows for increments 5–17; documentation only |
| 19 | [#158](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/158) | `78645cf8` | The automated Excel gate references point at #156 (#157); documentation only |

#33's runtime change was withdrawn for increments 2 and 3 and restored on
`release/v1.4.1`. It reaches `main` in increment 4. Before that, the
maintainer ran it in Microsoft 365 Excel 64-bit (16.0.20430.20180) at `c058b63`
and `4a896d0`, whose VBA is byte-identical to the increment's: 81 cases, 656
assertions and 0 failures. In a red check with the fix removed only case 81
failed. The evidence is on
[#33](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/33#issuecomment-6606918958).

## Release record

Completed under [#101](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/101) from the frozen candidate on `main`.
Each field is one that [RELEASING.md](../../../RELEASING.md#evidence-record) or
#101 requires. The demo workbook was built by `tools\Build-DemoWorkbook.cmd`,
whose JSON build record names the commit, the imported files, the smoke-tested
actions and the package hashes (#47, #72).

| Evidence | Result |
|---|---|
| Candidate commit SHA | `189668851754ec10e4c86d02f2e12b16739a54a5` on `main`, from [#207](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/207) |
| Version and tag | `1.4.1`; annotated tag `v1.4.1` (object `40b501bc6731dc3499139a6e98bcbcef83aa1cd2`), unsigned; previous tag `v1.4.0` |
| Static gate | 31 of 31 checks with the tag present and wiki pin `d0c2b7e4a646820261483bee7451c26fda90716e`; hosted Static checks [run 46167180526](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/actions/runs/46167180526) passed on the candidate after the tag was pushed |
| Excel version and bitness | Microsoft Excel 16.0 build 20527 (16.0.20527.20194), 64-bit |
| Regression-suite result | 93 cases, 861 assertions, 0 failures; cleanup completed |
| Local Excel runner JSON (#11) | `cpm-regression-1896688-20261128T163406Z.json` (suite) and `cpm-regression-1896688-20261128T163517Z.json` (suite plus the four measurement examples, all passed), both from a clean checkout of `main`, retained by the maintainer |
| Imported source SHA-256 | `M_cPM_TIMEWASTERS.bas` `8976013788537fbf06041430c6d200f179353dc891f245ef3b25fe5b21a56429`; `cPerformanceManager.cls` `962c356d16e7f1606510bd8f11829bc406d3b806f367714f881904a764b59665`; identical to the tagged files |
| Demo workbook build | `demo-build.json` at the candidate: compiled, eight demo actions smoke-tested with no failure, build metadata removed, reopened, no metadata problem |
| Release package and member SHA-256 | `VBA-Performance-Manager-v1.4.1.zip` `0411b6e167cb4e9556b8ddd11adce462f2120fb83399ecb3b5f39a2fed8f29ec`: `PERFORMANCE MANAGER.xlsm` `60f999e65f86f4dfbd976eab66326424f5f07b8d05deda07458ac1b3ec02778c`, `LICENSE` `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`, `NOTICE` `d616f18a99714ebc45dbb7033162a93bfdb0d3de27d1160f0cb9dcb5beae5cd8` |
| Artifact filenames and sizes | `VBA-Performance-Manager-v1.4.1.zip`, 3,146,521 bytes; `release-manifest.json`, 1,837 bytes, SHA-256 `02d893480d0f7d98c590ef05359533415d8f41abbf041eebd3cd6b00e941c9d9`. Package members: workbook 3,637,579 bytes, `LICENSE` 11,357, `NOTICE` 283 |
| `tools/release_provenance.py` output | Version 3.1.0, from a clean checkout of the tag; reproduced unchanged in the [release notes](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/releases/tag/v1.4.1) |
| Package smoke test | Extracted package held exactly the workbook, `LICENSE` and `NOTICE`; the extracted workbook opened and ran Run Benchmark, and was left unsaved (re-hashed unchanged) |
| Post-publication check | Both assets downloaded from the release and re-hashed: identical to the values above; the release body matches the prepared notes; the GitHub source archive holds `VERSION` `1.4.1` |
| Release-preparation pull request | [#206](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/206) into `release/v1.4.1`, aligned to `main` by [#207](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/207); workbook changes in [#195](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/195), [#204](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/204) and [#205](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/205) |
| Validation timestamps (UTC) | Suite 2026-11-28 16:33:27–16:34:06; suite with examples 16:34:39–16:35:17; workbook build 16:36:00–16:38:19; tag 16:46:11; release published 16:56:08 |
| Known deviations and approving reviewer | See below. Approved by the maintainer, Daniele Penza |

### Deviations

- **Office 32-bit** execution remains unverified under
  [#29](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/29) (v1.5.0, contributor-dependent).
- **The tag was recreated before publication.** The first `v1.4.1` tag object,
  `48200c13f24311a4b03aec9d17b6d0c3c6b51cb6`, targeted the same commit but
  recorded a tagger identity other than the project's noreply address. It was
  deleted and replaced by `40b501bc…` about three minutes later, before the
  GitHub Release existed and before any asset referred to it. RELEASING forbids
  recreating a public tag to hide a mistake; this record states it instead.
  The target commit, the manifest and the assets are unaffected.
- **Changelog date.** The tagged `[1.4.1]` section is dated 2026-11-30, the
  planned date; the release was published on 2026-11-28. The released section
  stays as tagged; README and INSTALLATION give the publication date.
- **Provenance wording.** The tool's generic statement that the modules are
  imported by hand predates `tools\Build-DemoWorkbook.cmd`; the release notes
  say so above the unchanged block. Building in CI is tracked in
  [#87](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/87).
