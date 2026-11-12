# Multi-master authoring foundation verification

This delivery implements [ADR0057](../architecture/decisions/0057-master-theme-ownership.md). It advances AUT-07.1.1 as a native-output/preview foundation while keeping complete template, artwork, editing and native application gates open.

The [shared public-authoring fixture](../../tests/fixtures/multi-master.mjs) has red and blue masters with the same theme role and layout-local placeholder key, an unused third explicit master, interleaved red/blue/red/default pages, independent notes and distinct fonts/backgrounds. No XML patch creates its source.

[Core tests](../../packages/core/test/masters.test.mjs) check input capture, frozen getter state, detached effective themes, copies retaining layout identity, legacy custom documents/absent optional registry, invalid declarations and missing references, local none/zero backgrounds and rejected unsupported parent/artwork keys. [Layout tests](../../packages/layout/test/masters.test.mjs) verify intrinsic sizing, table sizing and ordinary overflow measurements under the actual owning theme, including layout objects; incomplete transported domains reject. The initial test exposed global-theme diagnostic measurements; the corrected resolver checks all measurement paths.

[Exporter tests](../../packages/pptx-exporter/test/masters.test.mjs) check independent master/theme/content-type entries, master-owned backgrounds, interleaved bidirectional relationships, globally unique native layout IDs, unused master preservation and independent notes theme. Empty/omitted master registries generate byte-equal legacy output.

[Cross-package/browser verification](../../scripts/multi-master-parity.test.mjs) compares full normalized documents, Layout results, SVG results and PPTX bytes across Node, browser ESM and global with no network requests. Independent DOM traverses actual slide→layout→master→theme edges and checks exact expected color domains and unique layout identities. It does not reuse the Parser for those facts. Parser structure reads all four native masters and an actual structure-session no-op returns exact bytes. Installed runtime/strict type consumers exercise the public initializer and native graph.

[Prepared source manifest](multi-master-authoring-evidence.json) records a current public-authoring PPTX and SHA-256, hash-bound normalized snapshot, four SVG pages and a Chromium screenshot under a new local `.tmp/multi-master/run-*` directory. It is a source artifact, not a native-app capture; native save/edit/reopen is explicitly `not performed`. Missing ignored artifacts cannot be restored by the manifest alone.

Adversarial review focuses on domain leakage, inherited background duplication, relation/ID collisions, mutable registry state and accidental legacy changes. Tests respectively check actual theme chains and all measurement paths, master-only background ownership, interleaved IDs/unused owners, frozen capture/detachment and legacy bytes/custom document compatibility. Full master artwork/placeholder inheritance, shared mutations/override provenance, template replacement, source semantic import/writes and actual native operations remain pending.


## Completed checks

Full `pnpm build` and `pnpm test` exited 0. The full suite includes Core150, Layout139, Parser696, Exporter56, Renderer19, CLI11, 34 visual checks and 118 import integration tests; isolated installed Node/browser/types consumers pass. Logs: `.tmp/masters-final-build.log`, `.tmp/masters-full-test.log`, `.tmp/masters-focused-final.log`, `.tmp/masters-typecheck.log`, `.tmp/masters-lint.log`. Production code was fixed before the full command started; the readable-background fixture update was additionally checked by all 10 focused groups and later import integration.

The existing native notes inventory runner was re-executed under the new Core: 10 inputs, 98 pages, 4 fields and 7550 source references with exact no-ops still pass (`.tmp/masters-notes-regression.log`). All seven prepared artifact hashes were checked; the current screenshot equals the inspected screenshot bytes. These remain browser/source checks. Native opening, edits and three save/reopen rounds are the next evidence task and are not complete.


## Native correction and no-edit admission

Subsequent [PowerPoint validation](multi-master-native.md) found that the original public source triggered repair because master IDs overlapped layout IDs. The original failed source is preserved. Production allocation now places new master IDs after the full layout range; only those presentation ID values change. A two-factor actual-open control separates the collision from unused master presence. The corrected unpatched public source has four first-open PDF pages and twelve pages over three actual save/close/reopen rounds, with independent full source domains/styles, native raster/trace/font-program stability and 12 actual corruption rejections. This updates the earlier native-pending state for no-edit admission only; actual layout/theme/placeholder edits and complete artwork/template gates remain pending.
