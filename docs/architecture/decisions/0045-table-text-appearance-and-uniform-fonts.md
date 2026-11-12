# ADR0045: Table reference text colors and uniform explicit script-font collections

Status: accepted for experimental source projection, 2026-10-26. Complete default appearance, script selection/shaping and native fidelity remain open.

The single-family Core authoring limitation described below is superseded by [ADR0046](0046-script-font-authoring-and-selection.md). This ADR's existing source projection still uses its bounded single-family path; adding authoring slots does not relax source admission automatically.

## Decision

Apply a table region's `fontRef` color as a default below the region's direct color, then retain existing later cell list/paragraph/run declarations. Reuse source paint/color-transform resolution and its target theme/color-map scope. Color and font eligibility are independent: an unsupported font index does not erase a supported explicit reference color. A higher unsupported color blocks an older color and alpha; a later supported color resets both.

The [FontReference description](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontreference?view=openxml-3.0.1) defines a theme font plus color. The [table text-style example](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablecelltextstyle?view=openxml-3.0.1) supplies an additional direct color; the projection gives that direct declaration priority. This is source interpretation, not a claim that every producer follows the same behavior.

Resolve explicit major/minor lt/ea/cs theme references independently inside the selected font component. An absent/empty selected slot blocks; do not guess language-specific supplemental fallback or borrow an older component. Parsing remains under existing XML/path/work budgets and cancellation checkpoints.

Core currently has one family per run. A table's explicit font collection may project to that field only when the explicit Latin, ea/cs and supplemental declarations can agree on one concrete family. Use the shared run font reader for Latin and ea/cs; table collection handling owns its supplemental script entries. Preserve dependency paths for every agreement. Missing/empty/unresolved/different alternatives block the older font rather than silently retaining Latin. Duplicate interpreted script identities reject. Unknown attributes/extensions retain located audits.

Equal explicit theme slots also work for local shape/text runs. This does not add automatic script segmentation, language fallback, font matching, availability checks, glyph coverage, bidi or shaping. `fontRef` without an explicit collection retains its existing Latin-family projection boundary; it does not promise that the entire theme's supplemental font set has collapsed to Latin.

## Boundaries and evidence

No new package, dependency, public report schema or Core font model. Source choices remain in Parser, with the reference index as font source and actual color/font nodes as dependencies. Existing scalar V3 saves retain source style/theme bytes. Font/style mutations remain blocked atomically; fresh materialization is a separate operation.

[Twelve Parser groups](../../../packages/pptx-parser/test/table-text-appearance.test.mjs) cover two dialects, source precedence, alpha reset/zero, unsupported blockers, direct/theme/script collections, missing and duplicate entries, independent font/color components, local shape runs, two masters, three source saves, exact budgets, capture/cancellation and retry. [Two browser/cross-package groups](../../../scripts/table-text-appearance-parity.test.mjs) add ESM/global/Node agreement, independent source DOM and three fresh output rounds. Installed consumers use public package exports.

[Independent native evidence](../../testing/table-text-appearance.md) retains a real discrepancy: reference-only source PDF uses white text and its resaved PDF uses black instead of the projected accent2. Explicit region color, local run color and uniform-collection color agree. The producer substitutes LiberationSans for Arial; font fidelity is not certified. These remain native admission gates under IMP-06.1.4.1/.4.3, not successful default-appearance acceptance.
