# Named Presentation Fixture Corpus

The shared [fixture definitions](../../tests/fixtures/presentations.mjs) are the source of truth for fixture IDs, capability labels, and expected warnings. They accept the public Core factory and are reused by export, preview, and regression tests. Fixture assets and content are repository-authored under the root MIT license; the 8×4 PNG contains four colored quadrants and no external resources. Visual tests use the existing [Roboto Mono fixture and license notice](../../packages/svg-renderer/test/FONT-NOTICE.md).

| ID | Contract exercised | Preview snapshot | Export assertions |
| --- | --- | --- | --- |
| `text` | Rich runs, XML characters, numbered/bulleted paragraphs, URL query string, internal slide link, notes | [text](../../packages/svg-renderer/test/snapshots/corpus-text.png) | Native numbering, escaping, speaker notes, valid link relationships |
| `theme-layout` | Theme colors/fonts, inherited background, placeholder, override, hidden slide, metadata | [theme/layout](../../packages/svg-renderer/test/snapshots/corpus-theme-layout.png) | Placeholder binding, inherited versus local background, hidden flag |
| `geometry` | Editable shape text, anchored connector, two nested groups, rotation and flip | [geometry](../../packages/svg-renderer/test/snapshots/corpus-geometry.png) | Resolved endpoints, group nesting, complete relationships |
| `images` | Shared real PNG, stretch/contain/cover/crop | [images](../../packages/svg-renderer/test/snapshots/corpus-images.png) | One packaged media part, crop output, browser image decoding |
| `table` | Three-column heading span and two-row cell span | [table](../../packages/svg-renderer/test/snapshots/corpus-table.png) | Native grid/row spans and vertical merge |
| `charts` | Bar/line/pie with embedded workbooks and an asynchronously loaded image | [charts](../../packages/svg-renderer/test/snapshots/corpus-charts.png) | Native chart families, nested workbook structure, concurrent export independence |
| `missing-asset` | Deliberately unsupported local path in browser-neutral export/preview | No baseline; warning contract | Explicit load/omission warning, valid remaining package |

## Run and reproduce

```sh
pnpm test:fixtures
pnpm test:visual
pnpm fixtures:export
```

`fixtures:export` writes PPTX, IR JSON, per-slide SVG and a manifest containing SHA-256 hashes and warnings to `.tmp/fixtures`. Re-running replaces only these generated artifacts. To choose another directory after building, run `node scripts/export-fixtures.mjs /absolute/output-directory`.

[Corpus tests](../../scripts/fixture-corpus.test.mjs) check deterministic IR/output, detached Layout, expected warnings, key XML semantics, generated ZIP structure, every XML part's parseability, internal relationship targets/references, content-type coverage, and embedded XLSX packages. The XML verifier runs in Chromium's DOM parser; mutation tests deliberately corrupt XML, relationships and content types to prove it rejects those failures. It is not a full OOXML schema validator or a parser for untrusted input.

The [package-local concurrent export regression](../../packages/pptx-exporter/test/concurrency.test.mjs) also runs without a browser. Eight overlapping exports must match an isolated export byte-for-byte. The chart fixture includes an asynchronous asset so the test exercises the actual previous race.

## Visual review rules

The corpus renders with a pinned font, viewport and device scale. New snapshots must be inspected before acceptance; do not regenerate baselines to conceal a regression. The corpus permits at most 1% differing pixels at a 0.2 pixel threshold to allow limited platform rasterization variation. Existing legacy snapshots retain their own settings.

Snapshots establish current preview behavior, including explicit approximations; they are not proof of Office rendering parity. Font metrics, shape text, merged tables and transformed groups require separate application-level inspection. Record any known divergence in the capability/application matrix rather than changing expected output silently.

## Adding a regression

Add a named fixture or a narrowly focused package test, declare expected warnings, and assert the behavior at the owning boundary. New export features add XML/package evidence and a preview baseline or documented preview fallback. Import/round-trip features will additionally record source-producer versions, supported edit operations, preserved structures and reopen results. Keep externally sourced fixtures small and record provenance/license before checking them in.

## Content import image fixture

`createImageImportFixture` in the [content fixture module](../../tests/fixtures/content-import.mjs) reuses the repository-authored 8×4 PNG above across two slides and three pictures. It covers a nonuniform crop, fixed opacity, flips/rotation and a transformed parent group. It is exercised by [cross-package import tests](../../scripts/import-content-parity.test.mjs), the Chromium authored/imported pixel comparison in [browser consumers](../../scripts/browser-bundles.test.mjs), and the optional [native image experiment](../../scripts/inspect-image-import.mjs). It does not change the seven general corpus IDs or their snapshots.

Independent [Parser fixtures](../../packages/pptx-parser/test/image-fixture.mjs) also reuse the 1×1 GIF literal from the existing asset tests and include a deliberately synthetic JPEG frame header. The synthetic JPEG tests header parsing only and is not a valid full-image decoding fixture. Negative cases cover malformed dimensions, conflicting MIME, external links, unsupported clipping/crops, relationships, byte budgets and stale plans. See the [image import evidence](image-import-evidence.json) for exact native application/environment/hash information and unverified scope.

The repository-created [EXIF orientation fixture](../../tests/fixtures/exif-orientation6.jpg) is a 3×2 asymmetric RGB image generated with Pillow 12.1.1 and an EXIF 0th-IFD orientation value of 6. It is synthetic test data; Pillow independently reads its dimensions and tag. It supports metadata-reader tests and does not establish application display or color fidelity.

[The normalized PNG control](../../tests/fixtures/exif-orientation6-normalized.png) was generated by applying Pillow 12.1.1 EXIF transpose to the synthetic JPEG; it is used to exercise the resolver conversion contract and does not prove a general decoding implementation.

The [progressive JPEG component fixture](../../tests/fixtures/jpeg-progressive-components.jpg) is an 8×8 synthetic RGB gradient generated with Pillow using progressive JPEG encoding at quality 85. It exercises repeated component identities across separate scans in picture replacement validation; it does not prove entropy decoding or full JPEG conformance.
