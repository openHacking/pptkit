# Diagram cache consistency research protocol

Contract: [ADR0029](../architecture/decisions/0029-diagram-cache-consistency-research.md). This private session manages exact dependency witnesses and byte-plan publication; it does not certify native SmartArt caches or authorize source-part cleanup.

```js
import { createDiagramCacheSession } from "../../scripts/research/diagram-cache-session.mjs";

const session = createDiagramCacheSession(explicitDependencySnapshot, {
  sourceCache: optionalSourceBytes,
});
// Imported bytes remain missing/empty/unknown until explicitly rebuilt/validated.
const plan = await session.prepare({
  policy: "rebuild", generate: ownedGenerator, validate: independentValidator,
  signal: cancellationSignal,
});
if (plan.status === "ready") {
  const candidateCacheBytes = session.read(plan);
  // Assemble/validate/publish the actual package in a caller-owned serialized
  // transaction. This session alone provides no filesystem atomicity.
  session.commit(plan);
}
session.update(nextDependencySnapshot); // present old payload becomes invalidated
session.dispose();
```

A snapshot includes ownerId, exact layoutIdentifier, complete/partial coverage, generatorVersion, environmentKey, width/height/rotation/flips and dependencies. Each dependency has role, unique id and lowercase SHA-256. Data/layout/colors/style are singular; complete also needs theme evidence. Additional resources and environment/version changes remain explicit. Coverage is the caller's assertion; the actual format adapter must derive and validate its footprint. Equivalent-looking XML/transforms are not automatically equal. Managed-current refers to this declared generator/validator scope only.

Block returns a reason for unavailable/unverified/invalidated data. Invalidate stages omission, with no OPC relationship/part deletion. Rebuild stages positive-length validated bytes. Authentic plans are single-use, revision/epoch/signal-bound and session-owned. Release unused/stale plans; default prepared capacity is four. Generation/validation must honor cancellation; disposal detaches the external signal relay immediately, but arbitrary callbacks are not forcibly stopped. Failed preparation never replaces the selected cache. The original source package stays outside this session.

## Automated checks

```sh
rtk proxy node --test scripts/diagram-cache-research.test.mjs scripts/cycle-spacing-research.test.mjs scripts/diagram-cache-browser.test.mjs
rtk proxy pnpm test:diagram-research
rtk proxy pnpm test:import
```

Seven pure cache groups and four pure spacing groups run without third-party source files. One browser integration independently validates actual geometry-only SVG cache metadata and circle/edge structure against hand-derived expectations; its scope explicitly excludes text and native cache payloads. Mutation, stale generation/validation, cross-session plans, input ABA, publication changes, quota/exact size limits, escaped snapshot amplification, a 120-round state oracle, callback errors and dispose-time abort listener cleanup have focused controls.

## Native spacing/rounding attribution

Retain the earlier PowerPoint artifacts and run:

```sh
rtk proxy node scripts/inspect-smartart-cycle-layout.mjs /absolute/smartart-generated-powerpoint-evidence.json
```

The [live comparison evidence](smartart-cycle-layout-evidence.json) retains all original 0.1-point threshold failures and adds per-case spacing/nearest-EMU observations. Independent Python checks inferred circle ring parameters and rounding deltas. Raw line widths/alignments are retained; centered stroke extents are labelled hypotheses if alignment is unspecified. The theme-change control independently extracted widths of 1.5 and 1 point with exact unchanged boxes. Do not infer that all stroke edits leave native layout unchanged.

The 3/4/5-node inferred gap ratios differ and are observations, not universal correction constants. One EMU is 1/12700 point; applying nearest-EMU rounding does not remove the existing mismatch. No tolerance is widened and no inferred parameter is automatically fed back into layout. Native stale-cache fixtures, full dependency closure, native rebuild/playback and source-preservation updates remain future gates.
