# v1.4.1 release evidence

v1.4.1 is in development. No v1.4.1 release certification is claimed. The
latest published release is v1.4.0, whose evidence is recorded in
[docs/releases/v1.4.0](../v1.4.0/IMPLEMENTATION_AND_RELEASE_PLAN.md).

## Where the work is tracked

| What | Where |
|---|---|
| Open and closed work | The [v1.4.1 milestone](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/milestone/4) |
| Final certification and publication | [#101](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/101) |
| Evidence for each change | The closing comment of its issue and the pull request that delivered it |
| Branch and merge rules | [CONTRIBUTING.md](../../../CONTRIBUTING.md#maintainer-workflow) |
| Release procedure | [RELEASING.md](../../../RELEASING.md) |

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
| 3 | [#100](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/100) | `cac3a4cc` | Template-parity tooling and governance (#80–#86, #89) |
| 4 | [#106](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/pull/106) | `19a8a208` | #33's sticky Calculation exemption, the local Excel runner (#11 part 1), planning in issues (#90), the settings-script fix (#103) |

#33's runtime change was withdrawn for increments 2 and 3 and restored on
`release/v1.4.1`. It reaches `main` in increment 4. Before that, the
maintainer ran it in Microsoft 365 Excel 64-bit (16.0.20430.20180) at `c058b63`
and `4a896d0`, whose VBA is byte-identical to the increment's: 81 cases, 656
assertions and 0 failures. In a red check with the fix removed only case 81
failed. The evidence is on
[#33](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/33#issuecomment-6606918958).

## Release record

This section is completed under #101 from the frozen candidate. Each field is
the one [RELEASING.md](../../../RELEASING.md) requires.

| Evidence | Result |
|---|---|
| Candidate commit SHA | Pending |
| Version and tag | Pending (`v1.4.1`); previous tag `v1.4.0` |
| Static gate | Pending (count and JSON result on the candidate) |
| Excel version and bitness | Pending |
| Regression-suite result | Pending (cases, assertions, failures) |
| Executable-gate artifact (#11) | Pending |
| Release package and member SHA-256 | Pending (`VBA-Performance-Manager-v1.4.1.zip`) |
| `tools/release_provenance.py` output | Pending |
| Release-preparation pull request | Pending |
| Validation timestamps | Pending |
| Known deviations and approving reviewer | Office 32-bit execution remains unverified under [#29](https://github.com/danielep71/VBA-PERFORMANCE_MANAGER/issues/29); others pending |
