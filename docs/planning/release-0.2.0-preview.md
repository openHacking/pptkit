# 0.2.0 preview candidate

This local candidate collects the current provisional PPTKit capability set for evaluation. It is not a 1.0 stability declaration and does not require every future backlog task to be complete. The eight publishable packages share a version; the private SVG parser remains research code. No npm publication or dist-tag change is part of preparing this candidate.

The release scope includes Core authoring and IR v2, layout, PPTX fresh export, SVG and PDF rendering, bounded PPTX inspection and opt-in source-preserving edits, media inspection, and the CLI. Individual profiles and limits are defined in the [API documentation](../api/README.md) and [capability matrix](capability-matrix.md); the [changelog](../../CHANGELOG.md) records the detailed changes. Experimental profiles are not general Office compatibility guarantees.

Important boundaries for evaluation:

- Parser import is a scoped projection; arbitrary PPTX content, inherited presentation semantics and unrestricted structural source saves are not guaranteed.
- Existing text, chart values/categories/names, metadata and picture edits require their documented profiles and source guards. Chart series-order source IDs are reported experimentally, but order-save requests are blocked.
- SmartArt, animation, advanced typography, native table styles, broader notes/master behavior and application-specific fidelity retain the limitations recorded in the roadmap and linked evidence.
- Canonical IR remains version 2. Consumers implementing managed-slide interfaces or consuming newer optional IR fields should review [migration guidance](../migration/release-strategy.md) before upgrading.
- PowerPoint and LibreOffice observations are limited to named versions and fixtures; they do not imply blanket application compatibility. Interactive workbook editing and broader native operations remain open.

The local release audit records build, typecheck, lint, full tests, candidate-version installed-consumer checks, package archive hashes and any failed or untested checks. A candidate is ready for local review only after those checks are recorded; the separate v1 gates STB-01–STB-04 remain open.

## Local verification record

On 2026-11-12, using Node v22.19.0 and pnpm 10.13.1, the `0.2.0-preview.0` workspace passed `pnpm typecheck`, `pnpm lint`, `pnpm build`, and the full `pnpm test`. The full suite includes candidate-version packed-package installation and consumer declaration checks; its final import matrix reported 175 passing tests. The chart save evidence suite also passed before the version-only manifest changes (22 tests). The full test and build logs are `/tmp/pptkit-candidate-test.log` and `/tmp/pptkit-candidate-build.log` on this machine; temporary logs are not durable release evidence.

Eight local package archives are under `.tmp/release-0.2.0-preview.0/`, with `manifest.json` and `SHA256SUMS`. Each archive's package name, candidate version, and internal dependency versions were checked after packing; the private SVG parser is excluded. The installed-consumer suite independently packs and installs from the same candidate workspace, rather than installing these exact archive files. These local files may be removed by temporary-file cleanup and do not constitute an npm publication.

The build passed with a Vite warning that the example development application's minified JavaScript chunk exceeds 500 KB. Native-application observations remain limited to the versions, fixtures and operations in the linked evidence; the automated suite does not independently establish blanket PowerPoint or LibreOffice compatibility. Backlog items and the v1 stability gates remain open.
