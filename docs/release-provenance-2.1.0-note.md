# Release provenance 2.1.0

Issue #63 hardened the strict release-provenance gate on `v1.4.1`.

The tool now:

- writes a requested JSON manifest successfully before emitting publishable Markdown;
- accepts documented SemVer pre-release identifiers while rejecting invalid leading-zero forms;
- requires annotated release tags;
- prevents the manifest output path from aliasing the workbook, `VERSION`, the provenance tool, or protected source inputs;
- requires `--version`, `--tag`, and the tagged repository `VERSION` value to agree.

Exact implementation commit: `459e2a394dc9928ad9c31e8d11cdacf9e775a163`.

GitHub Actions Static checks run #203 (`34268016615`) passed all 15 checks on that exact SHA, including the expanded release-provenance fixture gate.
