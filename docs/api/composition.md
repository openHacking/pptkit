# Composition and frame constraints

`LayoutOptions.composition` accepts an ordered array of `LayoutComposition` rules. Both Layout entry points and the document-facing export/preview APIs accept these options. Rules operate on a private normalized copy before connector anchors, image fitting and chart geometry resolve. Authored documents and supplied IR remain unchanged.

## Selection and coordinate system

Every rule supplies `container: { kind: "slide" | "layout" | "group", id }` and `elementIds` (explicit `placements` for grids, or `elements` for constraints). IDs must name direct children of that container; selection does not implicitly cross group boundaries or include inherited layout elements. Use a separate layout rule to arrange template elements. The order of IDs determines placement order, while the original drawing/z-order remains unchanged.

The default area is the slide's full size, or the group's local `coordinateSize`, starting at `{ x: 0, y: 0 }`. An optional `area: Box` replaces that area in the same local coordinate system. `padding` is a nonnegative scalar or partial top/right/bottom/left insets, with omitted sides equal to zero. Group scale, rotation and flips are applied downstream; moving a group never rewrites its child coordinates.

`bounds` defaults to `frame`, using unrotated normalized frames. `visual` uses each frame's rotated/flipped axis-aligned bounds in the container coordinate system and translates the original frame accordingly. It does not expand a group to include outlying children, or measure text ink, shadows or strokes. Image fitting occurs after frame allocation, so a `contain` image may occupy a smaller visible area than its allocation.

Hidden elements, hidden slides and hidden ancestor groups are skipped by default. Set `includeHidden: true` on a rule to arrange them. An explicitly missing ID is diagnosed; a present hidden ID is skipped. Empty selections are valid no-ops.

## Operations

| Type | Required/optional fields | Behavior |
| --- | --- | --- |
| `align` | At least one of `horizontal`, `vertical`: `start`, `center`, `end` | Aligns every selected frame to that edge/center of the padded area. Unspecified axes remain unchanged. Overlap between selected objects is permitted because alignment is independent, not packing. |
| `distribute` | `direction`: `horizontal` or `vertical` | Places the first/last frame at the padded area edges and equalizes gaps between frames in supplied ID order. The other axis is unchanged. A single element goes to the start edge. |
| `stack` | `direction`; optional `gap`, `justify`, `align` | Packs fixed-size frames with the requested minimum gap (default zero). `justify` defaults to start and accepts center/end/space-between. `align` controls the other axis and defaults to start. |

For a stack, `space-between` distributes leftover space in addition to the minimum gap; with one item it behaves as start. Alignment, distribution, stacks and grids never stretch or shrink frames. Constraint rules may explicitly change selected dimensions. Negative gaps are rejected. [Intrinsic sizing](intrinsic-sizing.md) runs before composition. Explicit [frame constraints](constraints.md) can run anywhere in the composition sequence; pagination remains separate work.

## Ordering, diagnostics and failure

Rules execute in declared order. A later rule can intentionally reposition objects changed by an earlier rule. Across rules this is a sequential transformation. A `constraints` rule resolves its own explicit property dependency graph atomically; it does not infer dependencies between rules.

Malformed configuration (duplicate IDs, invalid axes/modes, nonfinite numbers, negative gaps/padding/area size) throws before any rule executes. A missing container/child, independently selected anchored connector, numeric overflow or unsatisfiable available space emits a warning and applies **none** of that rule's placements. Earlier successful rules remain effective and later rules still run. No partial placement or silent resizing is used to make an impossible stack fit. Constraint-specific failures and supported resizes are documented in [frame constraints](constraints.md).

Diagnostics live in `LayoutResult.analysis.diagnostics` and are forwarded by export/preview. Codes are `composition-container-missing`, `composition-element-missing`, `composition-anchored-connector`, `composition-space-unavailable`, and `composition-overflow`, and `composition-grid-unsatisfiable`. Paths identify `options.composition.<index>` and include the known slide/layout/group identity. Without composition or analysis options, the previous result shape and behavior remain unchanged. Enabling composition creates an analysis collection even when it is empty.

Alignment checks space only on requested axes. Distribution checks the distributed axis. Stacking checks both axes. These checks do not replace slide-boundary diagnostics: an explicitly off-slide area can be composed successfully and then reported by `diagnostics.slideBounds`. Only a machine-rounding tolerance is allowed for exact floating-point fits; meaningful negative available space fails.

Anchored connectors follow their referenced shapes after composition. Do not independently select them in a rule; that rule will be rejected with a diagnostic. Explicit bend/route points on an anchored connector remain authored points, without automatic routing. Free point-based connectors can be selected: their actual point bounds determine placement, and every endpoint/route point translates together.

## Checked example

<!-- doc-test: docs/examples/composition.ts -->
```ts
import { createPresentation } from "@pptkit/core";
import { resolveLayout, type LayoutOptions } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

export async function createStackedCards() {
  const presentation = createPresentation({ size: { width: 600, height: 300 } });
  const slide = presentation.addSlide({ id: "cards" });
  for (const id of ["first", "second", "third"]) {
    slide.addElement({ type: "shape", id, shape: "roundRect", box: { x: 0, y: 0, width: 140, height: 100 } });
  }
  const layout: LayoutOptions = {
    composition: [{
      type: "stack", container: { kind: "slide", id: slide.id },
      elementIds: ["first", "second", "third"], direction: "horizontal",
      padding: 30, gap: 20, justify: "center", align: "center",
    }],
    diagnostics: { slideBounds: true },
  };
  const resolved = resolveLayout(presentation, layout);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  const [pptx, preview] = await Promise.all([
    generatePptx(presentation, { layout }),
    renderPresentationToSvg(presentation, { layout }),
  ]);
  return { pptx, preview };
}
```

## Verification

[Layout composition tests](../../packages/layout/test/composition.test.mjs) cover direction/alignment/packing combinations, arbitrary supplied order, heterogeneous sizes, exact fractional fits, padding/gaps, failed-rule atomicity, missing identities, hidden content, group coordinates, rotation and connectors. [Cross-format parity tests](../../scripts/composition-parity.test.mjs) compare an automatically composed mixed deck against an independently positioned authoring document: the resolved geometry, complete PPTX bytes and SVG output agree. The deck includes shapes, image fitting, a table, a native chart, a group and anchored connectors.

## Explicit grids

See [grid composition](grid.md) for fixed/content/fraction tracks, spans, bounded allocation and explicit placement.

## Frame constraints

See [bounded affine size and position constraints](constraints.md) for equalities, bounds, peer/container references and cycle rejection.
