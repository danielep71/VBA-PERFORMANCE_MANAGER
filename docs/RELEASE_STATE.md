# Release-state consistency

The static gate checks source and a pinned wiki revision without network access.
The source `tools/release-state.json` records the full wiki commit. The separate
preparation command fetches only that exact commit into a new isolated directory:

```sh
python tools/release_state.py --fetch-wiki --wiki-root /tmp/cpm-wiki-input
python tools/vba_lint.py --wiki-root /tmp/cpm-wiki-input --json /tmp/cpm-lint.json
```

Choose a directory that does not exist. Preparation refuses existing directories.
Alternatively supply a clean local wiki checkout at the pinned commit, or set
`PERFORMANCE_MANAGER_WIKI`. Missing, dirty, or mismatched input fails the gate.
CI uses the same commands with read-only access; checking never fetches live wiki
content. Text and JSON report the source commit, checked wiki commit and expected
wiki commit. JSON also records the latest tag and its target. Local uncommitted
source edits are checked as working-tree content; commit the source before using
its reported identity as final evidence.

## Enforced surfaces

The existing version-stamp rule now includes root `VERSION` (missing, malformed
and mismatched values fail), while retaining class, module, tests and README
comparisons. Six additional rules are individually reported:

1. Release changelog sections: one first `Unreleased` heading and dated sections
   for the source version and release tags, with normalized version labels.
2. Comparison links: `Unreleased` runs from the latest tag to `main`; every
   released version links its predecessor to itself. The first tag uses its
   release page. The historical 1.1.0, 1.2.0 and 1.3.0 chain is included.
3. README current-release claims: the latest-release note and heading match the
   highest semantic release tag (`vX.Y.Z`), independently of development labels.
4. Wiki input: exact pinned commit, clean checkout, required pages available.
5. Wiki badges: Home's Release and Version History's Current badge match the tag.
6. Marked current-status blocks: README and the three wiki pages agree on the
   latest published release, development gate count and execution boundary.

The marked blocks are bounded by `<!-- release-state -->` and
`<!-- /release-state -->`. Their fields follow `status_block` in
`tools/release_state.py`; a latest-release claim of unreleased, superseded, or
pending certification/tag/assets fails. Arbitrary prose is not searched globally.
Historical certification counts, earlier-release tables and archived plans remain
valid. The existing released-body freeze remains an independent, unchanged rule.
Do not rewrite tagged release bodies or move tags to repair a comparison link.

## Paired updates and the transition

Source and wiki are separate repositories, so the update cannot be atomic:

1. Prepare matching source and wiki changes. Keep dated evidence as history;
   distinguish the published release from development checks. For this baseline,
   #29 remains contributor-dependent v1.5.0 work and execution evidence is 64-bit
   only; tooling does not certify Excel or require a 32-bit host.
2. Publish the wiki changes first. During this short transition, the wiki can
   describe the prepared source baseline, while older source commits still
   validate their previous pinned wiki. State this explicitly on the wiki.
3. Fetch the resulting wiki commit, inspect its pages, and put its full SHA in
   `tools/release-state.json`. Never select an obsolete wiki merely to pass.
4. Run preparation, the full static gate, local test discovery, quality checks
   and `git diff --check`. Commit the source and rerun for final source/wiki
   identities. Publish the source PR and require passing hosted checks.
5. Merge according to repository policy. The immutable pair now represents the
   baseline. Later wiki edits cannot alter results for this source commit. Each
   further paired update repeats the sequence; rollback restores the previous
   source pin and corresponding source claims together.

Before a new release tag exists, keep the source/current release stamps and
published blocks on the existing release, and describe development in Unreleased.
During release preparation the normal version promotion can temporarily fail
until the new tag and paired claims exist; never bypass the gate or mutate an
existing tag. Run final release evidence against the complete final pair.

## Fixture coverage

`tools/test_release_state.py` uses only temporary local Git repositories and
strings: valid and missing/malformed source versions; valid/missing/duplicate or
undated sections; each historical comparison and both Unreleased endpoints;
current README notes/headings versus legitimate history; both wiki badges;
valid and contradictory/absent/duplicate status blocks; missing/dirty/mismatched
wiki input; exact local fetch and refusal to reuse a destination; paired-update
failure, success and rollback. Both text rule integration and JSON identities are
covered. Existing compile/inventory CLI fixtures deliberately lack Git history,
so they assert all six release-state input failures in addition to their specific
source errors. They never fetch a wiki or silently skip production rules.
