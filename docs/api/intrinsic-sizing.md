# Intrinsic sizing

`LayoutOptions.intrinsicSizing` explicitly selects text, shape, image and group dimensions to resolve. An omitted axis retains its normalized size. Normalized IR has no reliable provenance for omitted versus authored dimensions, so zero dimensions never implicitly mean auto. Rules are opt-in; default calls retain existing behavior and public Core/IR contracts.

## Rules and order

Each `IntrinsicSizingRule` has a unique `elementId`, at least one of `width`/`height`, and optional `includeHidden`. Each selected axis is one of:

- `{ type: "content", min?, max? }`: derive a content size with optional bounds.
- `{ type: "parent", fraction?, min?, max? }`: use a fraction (default 1) of slide dimensions or the enclosing group's local coordinate size, bounded by min/max.

All values are logical points; bounds and fractions must be finite and nonnegative, with min ≤ max. Hidden slides, elements and ancestor subtrees are skipped unless the rule opts in. Group fitting excludes hidden direct children unless its own rule opts in.

The pipeline runs table sizing → intrinsic sizing → composition → connectors/images/charts → diagnostics. Fitting a group includes already measured table frames. Later composition can move children beyond that fitted group; rules are not iterated to a fixed point. Use explicit layout stages if arranging children before fitting is required. Diagnostics can detect resulting slide/text overflow but do not guarantee all children remain inside their group.

## Text and shapes

Content sizing requires a synchronous `textMeasurement` provider; use the explicit [browser adapter](browser-measurement.md) with already loaded fonts or supply another provider. Shape content sizing requires shape text. Parent sizing can resize a plain shape.

Width measures unwrapped, left-aligned content, including margins; max width may constrain it. Height then measures at the resulting content width using the original wrap setting and paragraph alignment, with top alignment for measurement. Minimum height adds space; content exceeding maximum height rejects the intrinsic batch rather than silently clipping. Width-only rules preserve the original height; request text overflow diagnostics when appropriate. The caller's text styles remain unchanged.

For parallelograms, content width inverts the preset text rectangle at the resolved height and includes margins. For `wrap: false` with a fixed or parent-sized width, content height solves the preset geometry and remeasures at the resulting width; it handles the nonmonotonic text height at high slants, and rejects an impossible bound or a result that does not fit. When both axes use content and wrapping is off, Layout selects a deterministic fitting aspect ratio and remeasures the resulting frame. If a height maximum is below the initial choice, Layout retries at that maximum with a wider frame. The chosen frame is not a general global optimum; a maximum can still reject it even when another aspect ratio might fit. With a fixed or parent-sized width, wrapped content height starts from the widest possible text rectangle, then repeatedly solves and remeasures the resulting rectangle, up to 32 iterations. The result is verified but is not guaranteed minimal; nonconvergence or a failed bound reports `intrinsic-shape-height-unavailable` rather than claiming no other solution exists. For wrapped content with both axes automatic, Layout chooses a candidate width from the unwrapped content, authored height and width bounds, solves the height, then remeasures both axes. If the remeasured text extends beyond that frame horizontally, or its measured height exceeds the height maximum, Layout may widen and retry within the width maximum, up to 16 candidates; measurement-provider errors are not retried. It does not search every possible width; an unverified candidate rejects the batch even if another width could work. [Pinned-font Chromium tests](../../packages/svg-renderer/test/measurement.spec.mjs) cover one fitting wrapped case across Layout, SVG frame and PPTX geometry, plus one measured horizontal-overflow rollback; these do not establish native PowerPoint wrapping parity. A result that fails remeasurement likewise reports `intrinsic-shape-height-unavailable`; the intrinsic batch is atomic. Parent-sized height remains available.

This computes frames, not shared line layouts. Preview and native Office may reflow text differently; provider accuracy is not a claim of native font or line-break parity. Glyph overhang can still require overflow diagnostics.

## Images

Layout reads positive finite asset width/height metadata and never loads the image. One content axis follows the opposite resolved axis and the natural aspect ratio. Bounds that conflict with that ratio reject the batch. With both axes set to content, metadata dimensions seed a 1:1 logical size, then a common scale satisfies both bounds. No physical DPI conversion is inferred. Set a fixed opposite axis or explicit bounds when metadata is in pixels and a particular display size is required.

Explicit parent sizing on both axes allocates the requested frame; it does not enforce an aspect ratio. The normal contain/cover/crop stage follows sizing.

## Groups

Content fitting includes the local origin and transformed child frame bounds, including direct connector routes and same-group anchors. It excludes painted stroke/shadow extents. Negative child positions are rebased; group frame position and coordinate size change together to retain child world geometry under rotation and flips. Minimum bounds add trailing space; maxima cannot shrink children and fail when too small. Nested fitted groups are resolved before enclosing bounds.

An originally zero frame axis uses a seed scale of 1 because its previous collapsed scale cannot be preserved while deriving a positive content extent. Empty groups and zero resulting coordinate sizes fail with a diagnostic. Unselected axes retain their scale; explicit parent-sized axes intentionally rescale children.

Parent-relative child dimensions and content-relative parent dimensions form dependency edges. Cycles return `intrinsic-sizing-cycle` with the dependency chain, regardless of rule ordering. Connector-to-connector or out-of-group targets during fitting return `intrinsic-connector-dependency`.

## Failure behavior

Malformed rules throw before sizing. Missing targets, unavailable content/providers/fonts, failed measurements, unsatisfiable bounds, cycles or nonfinite geometry return warnings with object/source paths. The entire intrinsic stage is atomic: on failure all intrinsic changes are discarded, while successful preceding table sizing is retained. Subsequent composition still runs. Callers inspect `analysis.diagnostics` or forwarded exporter/renderer warnings to decide whether to accept output.

Tables use [table sizing](table-sizing.md); direct table, chart and connector intrinsic rules are unsupported. Neither input documents nor provider requests are mutated. A deterministic provider yields repeatable output for the same source and options.

## Checked example

The caller supplies a document containing `title` and `photo` IDs and loaded-font measurement services.

<!-- doc-test: docs/examples/intrinsic-sizing.ts -->
```ts
import type { PresentationDocument } from "@pptkit/core";
import { resolveLayout, type LayoutOptions, type TextMeasurementOptions } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function fitAndExport(document: PresentationDocument, textMeasurement: TextMeasurementOptions) {
  const options: LayoutOptions = {
    textMeasurement,
    intrinsicSizing: [
      { elementId: "title", width: { type: "content", max: 400 }, height: { type: "content" } },
      { elementId: "photo", height: { type: "content" } },
    ],
    diagnostics: { textOverflow: true },
  };
  const resolved = resolveLayout(document, options);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  return { generated: await generatePptx(document, { layout: options }) };
}
```

Verification: [Layout tests](../../packages/layout/test/intrinsic.test.mjs), [PPTX/SVG parity](../../scripts/intrinsic-parity.test.mjs), and [pinned-font browser measurement](../../packages/svg-renderer/test/measurement.spec.mjs). These tests do not establish native Office reflow equivalence.
