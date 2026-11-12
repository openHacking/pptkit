# Shared master heading and named chart controls

The [advanced recipe](../../tests/fixtures/master-advanced.mjs) owns a merged table, three chart families, native straight connector targets and ordinary heading text under red/blue masters. Four pages include two visible red pages, one hidden-artwork red page and a blue control. This advances AUT-07.1.3 without claiming general table/chart/image editing, complete template reset/rebinding or source master writes.

## Public shared replacement

The [public consumer test](../../scripts/master-advanced-parity.test.mjs) calls `replaceSlideMaster` to change only `MASTER red` to `MASTER red 2`. Node, browser ESM and global produce equal complete IR/origins/Layout/SVG/warnings/PPTX with blocked network. Only `ppt/slideMasters/slideMaster2.xml` changes. Hidden/blue page SVGs remain exact; both visible red pages update. Independent Chromium DOM checks actual master ownership, table merge/continuation/cell property order, three chart families per master and valid local connector target IDs. Shared objects are absent from local slide parts. The case runs in complete `test:import`; materialized table-style warnings remain explicit.

[Prepared evidence](master-shared-evidence.json) binds fresh original source9643618a… and fresh edited sourceb4fb8dd173f387a6237b69098316584be8df45f52da987d05aeaacbdd00f94b3. The separate actual native editing input0f06d3c164b9501a7a58795c94cae485f85a80890cd36ff472cd6e790d3ba58c is a byte-exact copy of the preceding cell-repair third native save. These identities are intentionally distinct. [Preparation](../../scripts/prepare-master-shared.mjs) refuses to overwrite the native fresh-source control; retained native files are not regenerated during verification.

## Actual native shared edit

In a new directory `.tmp/master-advanced/shared-native-vwggt5tr`, actual PowerPoint16.104 (25121423) opens the native input, enters Slide Master view, selects the red master and edits its ordinary heading. AX/screenshots capture the master selection/text edit and return to ordinary view. A before/after ordinary-slide print is followed by three actual Save As/close/exact-path reopen/print rounds.

The [independent oracle](../../scripts/research/oracles/master-shared-native.py) requires every package part to equal the original native input, after exactly the red heading text edit, two `dirty=0` flags confined to its run/paragraph-end properties, revision1→2 and existing validated clock/TotalTime handling. No other body/resource/relationship/theme/notes/layout/slide part changes. The raw changed files are only the red master plus core/app properties; native connector XML, shared table/caches/workbooks and visibility flags remain unchanged. Later packages remain exact under the same time handling, with revision fixed at2.

[Native evidence](master-shared-native-evidence.json) verifies the visible red heading updates on pages0/2. All other trace records, original heading glyphs, vector drawings and image records remain exact. Pixels outside the old/new heading bounds are exact; hidden page1 and blue page3 retain full raster and text traces. The heading PDF font changes only its format-6 `cmap` mapping order, with the same glyph mapping set. Every font stream byte outside `cmap`, including the complete table directory and all glyph outline/advance/sidebearing/units tables, remains exact; the other font programs remain exact. This is a named captured font-subset change, not permission to ignore all embedded font differences.

Twelve reopened pages retain exact full144dpi raster, raw text trace, drawings, image records and embedded font bytes against the immediate post-edit print. Four actual corrupt PPTX files are rejected: changed blue master, copied local artwork, lost hide flag and unintended heading font size. Three actual bad PDFs are rejected: covered table, swapped pages and a raster-identical font advance mutation. The last mutation uses a real checksum-repaired font stream, showing that pixel equality alone does not pass the bystander check.

The fresh Core-edited PPTX was also actually opened without repair and printed. All four page records are exactly equal to the manual native-edit print, including embedded font programs. This is one independent fresh-output first-open comparison; it is not described as three additional fresh-output rounds. It connects the public Core replacement to the observed native shared behavior.

[Five actual-file consumer records](master-shared-native-parser-evidence.json) cover native input, three edited native saves and fresh Core edit. Node/ESM/global inspections/reports match, with fifteen exact source-only no-ops. Parser source artwork/style write authority is unchanged.

## Six-chart first-save review

The [separate chart oracle](../../scripts/research/oracles/master-chart-first-save.py) reviews six master-owned positive two-point bar/line/pie controls across the existing three cell-repair native rounds, eighteen full XML comparisons. The [bound result](master-chart-first-save-evidence.json) is tied to the exact original source/native proof, with no new native captures attributed to it.

All authored chart XML, categories/caches, colors, text, geometry, axes and data relations must match after these explicitly checked captured differences:

- `en-US` becomes `zh-CN` on chart language; style2 and `autoTitleDeleted=0` appear.
- Width/height `factor` modes and bar overlap0 disappear; manual layout numbers retain exactly the same binary64 values, including scientific notation. No epsilon tolerance is used.
- General non-linked category-axis number format and disabled data-label flags appear; pie leader lines are enabled while all data-label content remains disabled. See Microsoft's [data label model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.datalabels?view=openxml-3.0.1) and [layout mode explanation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/2c28b9f6-2128-40c8-8eeb-248651dcd7f3).
- The positive bar adds `invertIfNegative=1` and an exact white inverse-fill extension. Its unexercised negative branch is recorded, not certified. Pie points add `bubble3D=0`.
- Ten well-formed, distinct series/point GUIDs appear under one named extension; unknown extensions, malformed or duplicate identities reject. Later full-package checks retain their exact native identities.

These are observations for this explicit positive/A-B/no-label recipe. Locale-dependent categories, negative values, visible labels and other chart styles require new controls. Style2 is not declared universally equivalent to absent style. Workbooks remain exact payloads through native filename aliases, as independently checked by the cell-repair owner oracle. Original first-save package differences remain fully recorded; this review does not erase them.

Eleven actual corrupt ZIPs reject wrong color/font size, a one-ULP layout change, changed layout mode, visible labels, percent formatting, changed inverse fill, invalid/duplicate GUIDs, unknown extension and altered data cache.

## Adversarial review and remaining gates

Five likely failures are covered: flattening inheritance instead of editing one shared owner; updating hidden/other-master pages; losing glyph metrics behind unchanged pixels; accepting broad chart normalization; and confusing fresh Core output with a previously saved native file. Exact owner/package comparisons, physical bystanders, non-cmap font tables, eleven chart corruptions and distinct source bindings address them. General artwork/image editing, actual native table/chart data/style editing, full reset/rebinding, source import/writes and field/notes parents remain open.

## Check record

Final build/typecheck, lint/package boundaries/documentation/typed examples and diff whitespace checks pass: `.tmp/master-shared-final-build.log`, `.tmp/master-shared-typecheck.log`, `.tmp/master-shared-final-lint.log`. Complete `pnpm test` exits0 in `.tmp/master-shared-final-full-test.log`: Core208, Layout142, Parser696, Exporter66, Renderer20, CLI11, visual33, import126 and installed Node/browser/strict NodeNext consumers. The new public case and fixture helper were fixed before this run. Both independent oracles pass (`.tmp/master-shared-native-oracle.log`, `.tmp/master-chart-first-save-oracle.log`).

[Artifact verification](../../scripts/verify-master-shared-evidence.mjs) regenerates current fresh sources/IR/origins/Layout/warnings/four SVGs, checks90 artifact hashes/byte counts, thirteen producer/test dependencies, ten oracle dependency bindings across two proofs and fifteen actual-file consumer no-ops (`.tmp/master-shared-artifact-verification.log`). The contact sheet was viewed. Cell-repair, ordinary-artwork and format current-source compatibility verifiers also pass in `.tmp/master-shared-{cell,artwork,format}-compatibility.log`; their historical native sources remain untouched. No commit or publication was made.

The subsequent [rebind table paragraph-font repair](layout-rebind.md) has a separate [current producer bridge](master-paragraph-font-bridge-evidence.json). Fresh before0ecf999b…/edited8cd5b9e1… differ from historical9643618a…/b4fb8dd1… only by six explicit Arial paragraph-mark font additions in two master parts. Two new actual first-open prints match all eight corresponding historical page records and complete embedded fonts. The source edit still changes only the red master heading part. The verifier now records current producer bindings separately from historical sources/rounds; `.tmp/layout-rebind-shared-compatibility.log` passes90 historical artifact checks, fourteen current dependencies, original chart/Parser checks and the new bridge. These prints do not relabel historical reopen files as new-source rounds.
