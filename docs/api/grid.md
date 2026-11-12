# Grid composition

`GridComposition` is an explicit `LayoutOptions.composition` operation. It arranges existing fixed-size element frames in bounded tracks. Text/image/group intrinsic sizing can resolve those frames first. This grid does not stretch objects, reflow text to a cell width, create implicit tracks or infer automatic placement. Explicit [frame constraints](constraints.md) can run before or after the grid; this API is not CSS Grid.

## Tracks and placement

Supply nonempty `columns` and `rows`, and a `placements` array (instead of `elementIds`). Track types are:

- `{ type: "fixed", size }`: exact logical-point size.
- `{ type: "content", min?, max? }`: grows from its minimum to satisfy frame/span demands, subject to its maximum.
- `{ type: "fraction", weight?, min?, max? }`: satisfies content demands first, then shares remaining space in proportion to positive weights (default 1), redistributing growth after a maximum is reached.

Minimum defaults to zero; omitted maximum is unbounded. Sizes, bounds and gaps must be finite/nonnegative, and min must not exceed max. Fraction weights must be finite and positive. Fractions allocate *additional free space*, not proportional total track sizes after content demands.

Each placement specifies `elementId`, zero-based `row`/`column`, optional positive `rowSpan`/`columnSpan` (default 1), and `horizontal`/`vertical` alignment (`start` default, `center`, `end`). Every cell must lie inside the declared tracks. IDs must be unique direct children. Multiple elements may overlap cells deliberately; drawing order is preserved. There is no occupancy-based auto placement.

`columnGap`/`rowGap` default to zero. Container, area, padding, hidden-item handling, visual/frame bounds and free/anchored connector rules are shared with [composition](composition.md). Explicit empty tracks still occupy their configured minimum/fixed space and fraction allocation; hidden items do not impose content demands unless included.

## Deterministic allocation and precedence

1. Table measurement and explicit intrinsic sizing finish before composition. Their resulting frames provide grid content demands.
2. Fixed tracks and track minima establish the baseline. A spanning element requires the sum of its tracks plus interior gaps to fit its frame.
3. Span demands sort by ending track, then by descending starting track and ascending demand. Each demand grows the rightmost available capacity first. This deterministic policy lets overlapping spans share capacity without the false overflow caused by independent equal splits. Single spans can therefore leave earlier content tracks at zero.
4. If minimum feasible tracks plus gaps exceed the available area, the whole grid rule fails. Otherwise fraction tracks share extra space with caps; unused space stays at the trailing edge when no eligible fraction tracks remain.
5. Frames align inside their resolved cells, preserving their dimensions. Later composition rules may move those frames again.

No element or track references another rule: execution order is explicit and this grid introduces no dependency graph. Intrinsic parent/child cycles are handled by the preceding sizing stage. A separate constraints rule resolves its own explicit affine property dependencies and rejects cycles. Across rules, declared composition order determines precedence.

## Failures and evidence

Invalid contracts throw before composition. Track caps, fixed-size conflicts or insufficient area produce `composition-grid-unsatisfiable` at `options.composition.<index>` and retain all positions for that rule; earlier successful rules remain and later rules continue. Numeric placement overflow uses `composition-space-unavailable`. Input documents stay detached.

[Grid tests](../../packages/layout/test/grid.test.mjs) cover mixed tracks, caps/weights, gaps/spans, ordering, visual/group geometry, hidden objects, invalid inputs and atomic failures. A bounded integer oracle enumerates candidate assignments independently to verify feasibility rather than only checking examples derived from the implementation. [Cross-format tests](../../scripts/composition-parity.test.mjs) compare complete PPTX bytes and SVG against independently positioned elements and connectors.

## Checked example

<!-- doc-test: docs/examples/grid.ts -->
```ts
import { createPresentation } from "@pptkit/core";
import { resolveLayout, type GridComposition } from "@pptkit/layout";

export function createGrid() {
  const document = createPresentation({ size: { width: 600, height: 300 } });
  const slide = document.addSlide({ id: "slide" });
  for (const id of ["label", "content"]) slide.addElement({ type: "shape", id, shape: "rect", box: { x: 0, y: 0, width: 100, height: 60 } });
  const grid: GridComposition = {
    type: "grid", container: { kind: "slide", id: slide.id }, padding: 20,
    columns: [{ type: "content", max: 140 }, { type: "fraction" }],
    rows: [{ type: "fraction" }], columnGap: 20,
    placements: [
      { elementId: "label", row: 0, column: 0, vertical: "center" },
      { elementId: "content", row: 0, column: 1, horizontal: "center", vertical: "center" },
    ],
  };
  return resolveLayout(document, { composition: [grid] });
}
```
