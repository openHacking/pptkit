# Native multi-master operations — 2026-10-27

This follows the [ID-collision correction and no-edit baseline](multi-master-native.md). In PowerPoint for Mac 16.104, three real operations were performed in exact predecessor copies, each followed by three Save As / close / exact-path reopen rounds and four-page slide-print captures. No original input or failure file was overwritten.

The [native manifest](multi-master-operations-evidence.json) binds the baseline manifest hash, all native PPTX/PDF hashes, operation screenshots, full app/core metadata differences, actual font programs and raster/trace observations. The [independent oracle](../../scripts/research/oracles/multi-master-operations.py) imports no production Parser/Core code. The captured directory is `.tmp/multi-master/native-b85qrfu0`.

## Operations actually performed

| Stage | Native action | Observed result |
| --- | --- | --- |
| Layout | On slide 1, choose Blue theme's `layout-1` from the layout gallery | Slide 1 binds layout ID 2147483651 and master 2147483655, joining slide 2; Georgia/red becomes TimesNewRomanPSMT/blue with the blue inherited background. Slides 2–4 remain unchanged. |
| Placeholder text | Enter slide 1 body text and append ` EDIT` at the end | Original `Owned 0` remains, new text inherits the blue theme/font. Native XML has original en-US run, a zh-CN space with en-US altLang, and en-US EDIT with zh-CN altLang. |
| Theme | Right-click gallery `Office 主题` and choose **apply to selected slides**, with only slide 1 selected | Native adds master 2147483657 and twelve layouts (17 total). Slide 1 uses copied layout 2147483669 under the new master; slide 2 keeps the original blue master. Actual first-page face is DengXian-Regular, accent is 4472C4 and background is white. |

There are 12 first-operation pages and 36 reopened pages over nine actual save/close/reopen rounds. Native source files are the three `*-round1/2/3.pptx` groups; `*-first.pdf` records the visible operation before its first save. Reopened AX text/screenshots and operation screenshots remain beside them. The generated three-stage native contact sheet was visually inspected: labels, changed theme/background and appended text are visible without clipping.

## Source and identity checks

The oracle requires each `*-input.pptx` to equal its preceding third-round file byte-for-byte. It verifies the original baseline's SHA against the prior native manifest. It resolves part/relationship renaming by native master/layout/slide identity rather than matching filenames.

For layout switching, the expected change is one slide's layout relationship and removal of that slide's local transform. Before allowing removal, the original transform must equal the actual blue layout's inherited transform exactly. The effective box remains x10/y20/300×60pt. All other master/layout/theme/slide/notes semantic trees and content types remain equal.

For text editing, only the first body's expected runs and explicit native language/dirty attributes may change. Original run style, body frame, relationships, shared definitions, other slides and notes remain compared in full. A pixel comparison of the original text region prevents hiding changes behind the appended text.

For selected-slide theme replacement, every old master, layout, theme, other slide and note remains semantically equal. Only presentation/content-type registrations and the first slide's layout edge may change. Exactly one new master/theme and twelve new layouts are admitted for this named capture; the first copied layout's body equals the original blue layout body. New native definitions are retained in the actual PPTX and remain stable across rounds, rather than being projected as modeled Core artwork. Ancillary payloads outside mapped parts are byte-equal except app/core metadata, whose complete before/after XML is recorded. Later-round payloads may differ only in core metadata.

The selected gallery theme is a localized native profile. Its palette and format scheme match the installed Office 2013–2022 asset after the name becomes Office. Its font scheme structurally matches that legacy asset with major/minor Latin names from its Hans entries and 17 supplemental entries from the modern Office asset. Both original assets are locally captured and hash-bound under `office-assets`, so later application updates do not silently alter the oracle inputs. This describes an independent structural match; it does not claim knowledge of the native internal assembly algorithm or that the English modern Aptos theme was applied.

## Visual and font distinctions

All operation and reopened pages match the intended text, actual font face, colors, inherited background and visible ink. The saved first/second/third rounds have exact full 144dpi raster, raw glyph trace and embedded-font-program equality. Full bystander rasters remain exact relative to the original baseline (and the text stage's verified predecessor).

Two source observations are retained explicitly:

- Across operations, shared TrueType subsets change when the target page adds/removes glyph use. Bystander GIDs and whole font hashes can change while the text is unchanged. The oracle requires the same Unicode, glyph origin/bbox, other trace properties, **simple TrueType glyph program bytes, units-per-em, advance and sidebearing**, plus exact full raster. It supports the named ASCII bystanders and fails on an unadmitted composite glyph; this is not a general font equivalence algorithm.
- The text operation's unsaved PDF has three text spans; the saved result has two. Character `0` at index 6 moves from x99.1875 to x99.18272399902344pt (−0.0047760009765625pt), and its span-reported spacewidth changes from 6.999999046325684 to 3.499999523162842. The full raster and font program stay exact. This one exact observation is isolated in the oracle and manifest rather than concealed by a numeric tolerance or applied to other glyphs. All three **saved** raw traces remain identical. It does not establish exact unsaved/saved vector-position equivalence.

## Actual negatives and public consumers

Eighteen actual corrupted PPTX files break target text, bystander text, notes, owner color, owner font and the target layout relationship across the three stages. Nine actual PDF files cover the label with white ink, swap page order, or alter a bystander font advance while preserving the entire raster. The font-metric controls have correct updated TrueType checksums and expose why pixel comparison alone is insufficient. All 27 files are rejected, with their hashes and reasons retained under `operation-negatives`.

The [captured-profile runner](../../scripts/inspect-native-multi-master-operations.mjs) verifies nine actual native input hashes with content projection, `script-families-v3`, `master-style-v1` notes and `text-runs-v3` save profile. It asserts the correct 10/20/300/60 geometry and independent effective font/color on all four pages, including Chinese source family 等线 versus actual PDF face DengXian-Regular.

Its [manifest](multi-master-operations-parser-evidence.json) covers Node/ESM/global full document/report/no-op/candidate equality and zero network requests: 27 exact no-ops/restorations, 27 mixed character/font-write rejections and 27 scalar candidate paths. Nine actual Node candidate files and full snapshots have SHA values. Independent ZIP checks require exact payloads outside the one actual source text leaf and exact replacement within it. Reimported candidate snapshots equal the approved Core snapshot. **These headless X candidates were not native reopened** and grant no source layout/theme/style writer.

## Adversarial review and remaining scope

1. Recoloring a slide cannot prove layout rebinding: actual native IDs/edges, inherited transform and font change are verified.
2. A shared-master theme change could leak into slide 2: the selected-slide command creates a separate owner, and all old definitions, other slide trees and full bystander rasters remain checked.
3. Extractable text could hide wrong/covered glyphs: actual face/program/ink/raster and real white-overlay negatives are checked.
4. Treating GIDs or span grouping as stable can produce false conclusions: bystander glyph programs and the precise unsaved-text exception are recorded; saved-round comparisons remain strict.
5. Native UI editing cannot authorize broader source writes: current permitted scalar paths and mixed-style rejection/restore are separately exercised, with candidates labeled as headless.

This completes the named layout/text/selected-theme operation evidence for the fixture. Complete Core shared-definition mutation/override provenance, master artwork, generalized template rebinding, full source master import and source layout/theme/style writes remain open. AUT-07.1.4 and its artwork/effective-value parent dependencies are not closed by these three native actions. The original complete backlog goal remains active.


## Completed current checks

The operation oracle completed all 48 native capture pages and 27 real corruption rejections. The nine-input captured runner completed all three consumers and source candidate boundaries. Both final manifests and their baseline hash chain are current; 63 actual positive/negative/snapshot/candidate file hashes were rechecked. The native contact image was visually inspected.

`pnpm typecheck`, lint/boundaries/docs examples, four Exporter/browser regression tests and diff checks pass. Logs: `.tmp/masters-operations-{oracle,parser-final,regression,typecheck,lint}.log`. The original default 4-master/5-layout baseline oracle and its four-input consumer runner were rerun after the internal oracle parameterization. This turn changes verification tools/evidence/docs, with no production TypeScript changes; the full workspace test command was not newly rerun. The previous complete run remains explicitly dated in the recovery record.
