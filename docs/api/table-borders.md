# Direct table borders and shared physical edges

Core accepts sparse `style.borders: { left?, right?, top?, bottom? }` on table cells. Each value uses `StrokeStyleInput`; missing properties inherit the cell's normalized uniform `stroke`. Missing edges keep uniform fallback. An explicit `borders: {}` opts the entire table into shared edge resolution without adding an override. Tables with no `borders` retain their legacy rendering path.

See the [typed example](../examples/table-borders.ts), [ADR0041](../architecture/decisions/0041-shared-table-borders.md) and [native observations](../testing/table-borders-evidence.json). This slice covers direct authoring, direct import and fresh output. Full table-style/theme inheritance and source style writes remain open.

## Conflict rules

Layout resolves each physical grid edge once. Between adjacent owners, precedence is: visible over invisible; explicit per-edge override over uniform fallback; greater width; dash rank `solid > dash > dashDot > dot`; exact tie uses the upper cell for a horizontal edge or left cell for a vertical edge. Visible means solid paint, positive width and positive paint opacity. Table opacity does not change conflict priority. These are PPTKit's deterministic rules, not a claim to implement the Office table-style cascade. An explicit noFill does not erase an adjacent visible line.

Two slots belonging to the same merged cell have no interior line. Empty unowned slots supply no candidate; an adjacent owned side can still contribute. Outer sides use the single owner. Colors retain direct RGB or theme-role references until rendering. Arrowheads, unsupported dashes, negative/nonfinite widths, invalid paints and unknown edge fields reject at validation; no diagonal edges are introduced.

`resolveTableBorders(normalizedTable, limits?)` returns detached `{ grid, cells }`. It is synchronous and pure. Automatic Layout resolution runs when any cell contains `borders` or a non-solid uniform stroke, after sizing/pagination, including nested tables. Thus dashed, dotted and dash-dot uniform strokes use the same shared-edge precedence and resource limits as explicit per-edge borders; ordinary solid-only tables retain the prior path. `cells` contains compact `rowIndex/cellIndex`, physical `columnIndex` and local boxes computed from prefix sums. Layout stores these as `cellLayout` and stores `borderGrid` with schema `pptkit-table-border-grid-v1`:

- `horizontal`: `(rowCount + 1) * columnCount`, horizontal line first, then column.
- `vertical`: `rowCount * (columnCount + 1)`, row first, then vertical line.
- Invisible merged interiors are explicit noFill strokes with width zero.

Frame/grid differences over 0.01 pt produce located `table-border-frame-grid-mismatch` warnings; resolution never stretches authored dimensions to hide the discrepancy. This is distinct from Parser's source frame mismatch finding. SVG uses exact cell boxes and emits one visible SVG line per grid segment over the HTML cells, applying root transform and opacity once. Cell fills retain paint opacity. The HTML fallback paths also retain the opacity of the uniform cell stroke; their CSS border placement and conflict resolution can differ from the resolved SVG border grid. Older externally supplied Layout data without a border grid still uses solid CSS borders even for non-solid strokes. It rejects unknown grid versions, inconsistent counts/array lengths and missing cell geometry; JSON Layout transport preserves the metadata.

## Bounded resolution

| Limit | Default | Maximum |
| --- | ---: | ---: |
| `maxGridSlots` | 100,000 | 2,000,000 |
| `maxEdges` | 250,000 | 6,000,002 |
| `maxWorkUnits` | 2,000,000 | 100,000,000 |

`LayoutOptions.tableBorderLimits` accumulates across all opted-in tables in one Layout request; standalone calls own independent budgets. Slot/edge counts are charged before allocation. Grid ownership and prefix geometry avoid pairwise cell-overlap scanning. Work is deterministic accounting, not serialized bytes or a hard peak-memory guarantee. Unknown limits/accessors and invalid positive integer bounds reject. `TableBorderError` has `code`, `path`, optional `elementId`: `table-border-limit-exceeded`, `table-border-grid-invalid`, or `table-border-style-invalid`. Failure exposes no partial Layout and does not mutate inputs; retry is independent.

## Native representation and import

Fresh PPTX retains editable native table cells. It writes all four physical border hints on each slot, plus logical outer sides on merged anchors when each side is uniform. Horizontal continuation cells duplicate rowSpan and vertical continuation cells duplicate gridSpan as required by Office's [merge interpretation](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/de4562ba-7901-48df-9803-414ded131f2b). This fixes missing right/bottom borders observed for a wholly merged table whose anchor previously held only its first physical slot's borders.

A merged logical side with unequal physical segments cannot fit one native anchor stroke. The anchor side is noFill, physical perimeter and neighboring owner hints remain, and located `table-border-segmented-native` warns that native applications may render or rewrite it differently. SVG retains every resolved segment; heterogeneous native fidelity is not fully admitted.

Office [ignores rotation/flips on graphic frames](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/0967224b-efc3-42dd-907a-3f8c1a240172). Nonneutral table transforms therefore use a native group with the authored drawing ID/action/accessibility and a private, locally positioned editable table child. `table-transform-grouped` makes the hierarchy change explicit; identity transforms retain the flat path. Root opacity is applied once. This encoding does not prove interactive editing or all-producer transform fidelity.

Parser imports unequal/incomplete supported direct sRGB/none lines as sparse borders; equal four-sided strokes keep the legacy uniform representation. Uniform outer segments on merged continuations aggregate into a logical side. Unequal or partially absent segments stay source-only with located findings. Legacy merges lacking explicit continuation lines keep their anchor-only projection. Theme/table-style inheritance and unknown modifiers remain separately audited. Scalar source text saves retain original border XML; style/structure mutations remain blocked.

## Evidence and limits

Focused Core/Layout/Parser tests cover normalization, priority, invalid strokes, merge ownership, quotas, frame mismatch, two dialects and three source text-save rounds. [Cross-package tests](../../scripts/table-borders-parity.test.mjs) compare Node/ESM/global, exact work limits, fresh RGB rounds, independent Chromium DOM physical edges/merge spans, SVG paint/opacity, sizing/pagination, JSON transport and source-write rejection. Uniform dashed strokes are additionally checked through three fresh export/import rounds against the resolved physical border graph. Exporter tests cover spans 2/3/7, logical merged-anchor sides, hidden interiors and group identity/action/opacity.

The [browser import fixture](../../scripts/browser-bundles.test.mjs) checks the merged table's cell geometry and compares rendered pixels across fresh export/import, allowing at most ten pixels to differ by one channel value. One independent capture observed five such pixels after the native transformed table was represented as a group; this tolerance is not a native-application fidelity claim.

The optional [independent runner](../../scripts/inspect-table-borders.mjs) uses Python ZIP/ElementTree and PyMuPDF, with deliberate XML corruption checks and source-control PDFs. In the recorded LibreOfficeDev 26.8 alpha experiment, ordinary/segmented/wholly merged RGB scenes cover all expected geometry/color segments (12/12, 20/20, 12/12). Strict line checks remain partial: native PDFs alter widths and ignore stroke opacity; native resave changes alpha/invisible line widths and collapses a wholly merged grid. The transformed scene also has geometry/flip discrepancies in both source-control and resaved PDFs; PNGs confirm these. No producer-specific calibration is applied to canonical dimensions or alpha. Dash pixel fidelity, interactive PowerPoint editing and source style-write gates remain unverified. AUT-05.1 and IMP-06.1.4 stay open at their full scopes.
