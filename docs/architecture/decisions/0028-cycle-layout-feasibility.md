# ADR 0028: Deterministic equal-circle Cycle layout feasibility

Status: private executable research, started 2026-10-25; verified 2026-10-26. Advances ADV-02.1.3.2 and informs ADV-03.6; no public family/cache/native admission.

## Problem and decision

The content, source-binding and native-identity experiments establish topology and identity, not a PPTKit layout. Native observations also disagree: PowerPoint's retained Basic Cycle caches have circular nodes, while the named LibreOffice non-square-frame observations have ellipses. Implement an explicit **owned equal-circle packing contract** before attempting full DiagramML evaluation or cache generation. Compare its outputs with native evidence and retain mismatches.

The pure [solver](../../../scripts/research/cycle-layout.mjs) accepts the exact cycle2 research profile, a flat ordered content forest with layout-generated closure, a point frame, optional size/spacing/direction/start parameters, a frame transform and an explicit synchronous text measurement callback. It emits detached circle boxes, connector polygons/arcs, text placement and located findings. It is not imported by production packages or exposed as a public Layout API. It does not consume or copy native cached geometry into its result.

## Geometry contract

All nodes share diameter D. For n >= 2 and sibling gap ratio g, adjacent center distance is `(1 + g) D`, giving unit orbit radius `(1 + g) / (2 sin(pi / n))`. Semantic root order determines angular order; node-array storage order does not. Positive angles are clockwise in the page's downward y axis. The default start is the top (-90 degrees), default g is 0.5. Managed IDs and original paragraphs are retained.

Compute all unit-node bounds and connector bounds first. Center their joint bounding box in the requested frame. An explicit D must fit; otherwise choose the largest diameter bounded by the frame and maxDiameter. Reject below minDiameter (default 8 points) rather than silently shrinking below that policy. This bounding-box centering can shift the orbit center relative to the frame center for odd node counts. Arbitrary start angle changes packing extents; it is distinct from rotating an already fitted frame.

| Node count | Policy |
| --- | --- |
| 0 | No nodes/edges; diameter zero; no vacuous text-fit success |
| 1 | Centered circle; no invented self-edge |
| 2 | Two distinct clipped directed circular arcs and arrowheads; an explicit proposed policy without native comparison |
| >= 3 | Directed block-arrow polygons centered in each adjacent chord gap, including final-to-first closure |

The two-node policy avoids coincident reversed straight arrows. Circle/track intersections are analytic; cardinal extrema and arrowhead vertices participate in packing. Polygons have explicit vertices. Paint strokes, shadows and other effects are outside the centerline/fill geometry bounds; callers need an explicit paint inset. Rotation/mirroring transforms circles/connectors about the frame center; output coordinates are page points inside the **transformed** frame, not a promise to fit the unrotated slide rectangle. Text stays upright inside its circle.

## Text and failures

Each circle reserves an upright square equal to 80% of its inscribed-square side. Without a provider, every nonempty node is text-unmeasured. With a provider, retain its raw x/y/width/height, compute an explicit text origin that centers the measured bounds (including negative bearings), and record placedTextBounds. Fits means only that the supplied extent fits this content rectangle. Overflows are located findings, or throw under the strict overflow policy. No implicit font, wrap, glyph coverage, font shrink or Office text fit is inferred.

The existing public browser measurement provider can be wrapped explicitly; unavailable/disposed/unsupported-font diagnostics must become a failed measurement, not zero-size success. Real pinned-font tests exposed negative top bearings, so placement now accounts for them instead of falsely flagging short text or discarding the measured origin. Long labels continue to overflow; unsupported CJK and disposed providers reject in the integration test. Callback inputs are detached; callback side effects are caller-owned and are not rolled back.

Defaults bound 10,000 nodes, 1,000,000 UTF-8 paragraph bytes and 300,000 shared work checkpoints across validation, geometry, transforms, measurements and output. Frame coordinates/sizes and scalar ratios have explicit finite ranges. Invalid graphs, unknown options, asynchronous/invalid metrics, over-budget input and insufficient geometry fail without a partial result. Production runtime/dependencies/public contracts remain unchanged.

## Evidence and native discrepancies

Six pure groups include hand-derived square/triangle coordinates and arrow direction, zero/one/two-node policies, root-order/storage independence, start/direction/rotation/mirrors, negative bearings/overflow/provider failures, exact budgets/retry and 120 seeded diagrams with independent pairwise separation/segment checks and inverse-transform containment across sizes/angles/gap ratios. Two browser groups connect the public loaded-font measurer and verify actual SVG circle/path geometry against hand-calculated expectations. The [browser preview](../../testing/smartart-cycle-layout.md) is manually reviewed; two-node arrows are visibly separate.

The [optional comparison](../../../scripts/inspect-smartart-cycle-layout.mjs) reads six retained PowerPoint cycle2 stages with 3/4/5 nodes. Independent Python ZIP/XML associates cached circles through presentation-to-content model references, verifies exact frame/cache assumptions and converts EMU with Decimal. A separate Python packing calculation independently checks every proposed box and all numeric deltas. The [evidence](../../testing/smartart-cycle-layout-evidence.json) records maximum coordinate differences of about **0.137–0.235 points**. All six exceed the explicit **0.1-point research threshold**. These are mismatches, not admission passes. Only circle boxes are compared; native arrows, text fitting, themes, other parameters and layout policies are not verified by these numbers.

Retain the deterministic geometry candidate and failure controls. Continue native spacing/rounding/paint attribution and parameter variations, non-square/direction cases, full text/style/overflow policy, two-node native behavior and per-family cache decisions. Process/Hierarchy layout, heterogeneous node sizes, public API, native cache generation and full family admission remain open. ADV-02.1.3.2 and all parent tasks remain incomplete.


## Spacing attribution update — 2026-10-26

[ADR0029](0029-diagram-cache-consistency-research.md) adds descriptive ring fitting, EMU rounding and stroke-change controls. Inferred native gap ratios vary by retained node count; nearest-EMU rounding still fails the unchanged 0.1-point comparison. The five-node theme pair changes line width but not circle boxes. These findings do not justify automatically substituting observed constants or a universal paint inset into the solver.
