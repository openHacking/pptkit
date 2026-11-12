# Cycle deterministic layout research protocol

Contract: [ADR0028](../architecture/decisions/0028-cycle-layout-feasibility.md). This is an owned equal-circle geometry experiment for the exact cycle2 input profile, not full DiagramML/native layout support.

```js
import { resolveCycleLayout } from "../../scripts/research/cycle-layout.mjs";

const result = resolveCycleLayout(managedCycleCandidate, {
  frame: { x: 100, y: 100, width: 520, height: 300 },
  direction: "clockwise", startAngle: -90, gapRatio: 0.5,
  minDiameter: 8,
  // nodeDiameter: 80, maxDiameter: 120,
  // transform: { rotation: 30, flipH: true, flipV: false },
  // measureText: ({ id, paragraphs, availableWidth, availableHeight }) => bounds,
  // overflow: "error",
});
```

The provider returns plain finite bounds relative to its text origin. The solver preserves negative bearings and emits textOrigin/placedTextBounds after centering. Check textFit/findings; absent measurement is unmeasured, never fits. Provider errors propagate. A wrapper around resolveLayout must check its analysis diagnostics when no measurement exists; see the [browser integration test](../../scripts/cycle-layout-browser.test.mjs). Glyph coverage and provider disposal stay explicit. Transforming a frame does not constrain its output to unrotated slide bounds; text remains upright. Stroke/effect extents are outside this geometric fill/centerline contract.

## Automated checks

After `pnpm build`:

```sh
rtk proxy node --test scripts/cycle-layout-research.test.mjs scripts/cycle-layout-browser.test.mjs
rtk proxy pnpm test:diagram-research
rtk proxy pnpm test:import
```

Six pure groups cover hand-calculated geometry, empty/single/two-node closure, transforms/order/direction, measured fitting/negative bearings, invalid and budget controls, and 120 seeded layouts. Pairwise circle and connector-segment checks provide an independent collision oracle; inverse transforms check conceptual-frame containment. Two browser groups use pinned Roboto Mono with explicit ASCII coverage, exercise long labels/CJK rejection/disposal, and independently observe SVG bounding boxes and separate two-node paths. Browser subpixel coordinates use a 0.001-point observation tolerance.

The browser run retains `.tmp/cycle-layout/browser-preview.png`; it shows a four-node block-arrow layout and the proposed two-node curved layout. Manual review confirms node labels and separate directed paths; it does not certify native geometry or text fitting.

## Optional PowerPoint cache comparison

Retain the prior generated PowerPoint evidence manifest and its immutable artifacts, then run:

```sh
rtk proxy node scripts/inspect-smartart-cycle-layout.mjs /absolute/smartart-generated-powerpoint-evidence.json
```

This creates a new artifact directory with full comparisons and paired-circle SVGs. The checked-in [evidence JSON](smartart-cycle-layout-evidence.json) hashes the input evidence, each PPTX and each SVG. Files are independently hash/ZIP/XML checked. Source/managed/cache identities link through explicit model references; labels and array order do not establish ownership. Direct unrotated cached circles and a single untransformed cache tree are required; unsupported geometry does not get guessed into the oracle.

The six retained 3/4/5-node PowerPoint stages exceed the declared 0.1-point threshold: maximum coordinate deltas are roughly 0.164 points (5 nodes), 0.235 points (3 nodes) and 0.137 points (4 nodes). Keep those failures visible. Independent Python computes proposed packing and deltas separately from the JavaScript solver. Native arrows, fonts, clipping, full parameters and cache freshness are not part of this numeric comparison. No new native UI operation is performed; two-node native behavior remains untested.

[Manual preview review](smartart-cycle-layout-review.json) records the inspected PNG hash and the limited visible observations separately from automated checks.


[Cache/spacing protocol](smartart-cache-consistency.md) documents the new ring-fit, nearest-EMU and raw stroke observations. The existing mismatches and their threshold remain unchanged; observations are not fed back as hidden layout calibration.
