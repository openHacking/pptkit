# Layout Engine

`@pptkit/layout` owns presentation placement and sizing transformations that should not live in authoring state or file serialization.

## Current contract

The package accepts a `PresentationDocument` through `resolveLayout()` or Canonical IR through `resolveNormalizedLayout()`. It returns detached metadata, theme, assets, layouts, slides, and resolved elements.

Implemented resolution covers connector anchors/bounds, image contain/cover geometry, and recursive group processing. Layout neither mutates its input nor writes files. Core normalization may already have resolved deterministic text-box heights before Layout receives the Canonical IR.

## Expansion direction

Layout is the owner for future:

- additional runtime measurement adapters and shared resolved text lines
- automatic layout beyond explicit alignment/distribution/stacks and bounded grids
- simultaneous/cyclic constraint solving beyond explicit acyclic affine frame constraints
- automatic table column sizing
- general text/story continuation beyond explicit table pagination

These capabilities are not current behavior and should be added as explicit, testable transformations.

## Controlled, not magical

Every layout capability must define its inputs, deterministic output, supported constraints, failure/overflow behavior, and measurement dependencies. PPTKit can learn from web layout without promising DOM or CSS parity.

## Non-goals

- full browser CSS compatibility
- arbitrary DOM fidelity
- implicit exporter-only auto-layout
- environment-dependent measurement hidden inside Core

See [Layout API](../api/layout.md) and [Rendering Pipeline](rendering-pipeline.md).


## Injected analysis boundary

Layout accepts optional synchronous text/font providers and diagnostic options through its public entry points. Frozen detached requests cross this boundary; provider results produce detached measurements and warnings, without changing authored geometry or Core's deterministic height estimate. Layout owns coordinate transforms, margins/alignment and overflow comparisons; the provider owns font metrics and text shaping/wrapping. Core receives no font, filesystem, network or DOM dependency. The separate `@pptkit/layout/dom` adapter owns explicit browser measurement side effects; resolved line composition remains future work.

Exporter and Renderer accept Layout options on document-facing entry points and forward diagnostic identities/paths as warnings. Rendering an existing Layout result reuses its analysis. Default calls retain previous output; analysis does not claim native application layout parity. See the [Layout API](../api/layout.md#opt-in-measurement-and-diagnostics).

The [DOM measurement adapter](../api/browser-measurement.md) is an explicit runtime entry, not imported by the pure Layout entry. It consumes caller-owned loaded fonts and a document, creates a disposable isolated host, and returns synchronous providers. It has no implicit filesystem/network/font loading, and introduces no new package or dependency.


## Explicit composition order

Composition plans are runtime Layout options, not new authoring/IR state. The orchestrator validates all rule configuration, clones the normalized document once when composition is present, then applies ordered atomic placement rules. Connector anchors, image fitting and chart geometry resolve from those placements; analysis runs last. Template rules address template elements, and group rules use local coordinateSize. No renderer/exporter invents its own placement defaults. See [composition contract](../api/composition.md).

## Measured table ordering

Explicit tableSizing rules run first on the detached normalized clone, then intrinsic sizing, composition, table pagination and connector/image/chart resolution. Successful tables carry table-local cellLayout in LayoutResult; the renderer enforces these bounds and the exporter uses the same native row dimensions. Table sizing and overflow analysis share a grid resolver. Fixed rows remain fixed, and unresolved tables retain their full geometry. See [table sizing](../api/table-sizing.md).

## Intrinsic sizing stage

Explicit `intrinsicSizing` rules operate after table measurement and before composition. Layout owns dimension dependency resolution and group rebasing; text/font providers remain injected. The stage works on a detached clone and discards all of its changes on a failed dependency. Group fitting uses the same internal connector geometry helper as final resolution. No OOXML, asset loading or authoring-provenance fields enter Core. See the [sizing contract](../api/intrinsic-sizing.md) for coordinate, failure and stage-order limits.

## Explicit grid allocation

Grid rules share composition ownership, container selection, atomic placement and downstream connector resolution. A pure internal track solver computes minimum feasible bounded interval demands, then allocates weighted free space. It consumes frames after intrinsic/table sizing and never changes Core state or element dimensions. Track references are explicit indices, so there is no hidden dependency graph or implicit grid growth. Bounded affine size/position relationships live in the constraints composition operation. See [grid semantics](../api/grid.md).

## Frame constraint solving

A constraints composition rule builds one node per selected frame property, expands derived metrics into their component dependencies, and evaluates a directed acyclic graph with an iterative topological queue. Equalities are hard requirements; bounds-only properties clamp their incoming values. Missing references, cycles, unsatisfiable bounds and nonfinite arithmetic retain the whole rule's input. Solver candidates are detached boxes; mutation occurs only after type-specific resize restrictions and all geometric checks pass. Subsequent composition determines precedence, and final connector/image/chart resolution consumes the new geometry. No environment services or Core/IR changes enter this solver. See [the public contract](../api/constraints.md).

## Table pagination stage

Pagination owns detached structural changes after composition and before final element resolution. A pure row-block planner respects merged-cell boundaries, repeated headers and page budgets. The per-table commit step allocates deterministic collision-checked IDs, retains source identity on the first fragment, inserts continuation slides and records source rows in LayoutResult.pagination. The exporter consumes the resulting slide order without implementing its own pagination logic. Only each fragment's rows and slide/table metadata are copied; source siblings/notes are not redundantly cloned for every page. The public [pagination contract](../api/table-pagination.md) defines its top-level-table scope, continuation layout policy and atomic failures. No Core/IR schema, filesystem behavior or implicit generic document copying is added.


Connector anchors use the target's local frame sites, applying target flips then rotation about the frame center. Resolution stays in the immediate group's coordinates; ancestor transforms apply downstream once. Explicit free endpoints/routes remain unchanged. Intrinsic group fitting uses the same transformed connector geometry as final resolution. Cross-group reference resolution remains outside this contract.

[ADR0041](decisions/0041-shared-table-borders.md) adds bounded direct table edge resolution after sizing/pagination. Layout owns priority and exact cells; format adapters consume the graph, with no format dependencies in Core.


`planVisibleImageReplacement` explicitly resolves the current layout for a top-level slide image, optionally applying captured composition rules before planning against that resolved frame. Core owns the replacement snapshot and fit policy; Layout owns image fitting. It does not mutate the document, persist composition rules, or provide source-package preservation writing. See [asset lifecycle](../api/asset-lifecycle.md#per-image-replacement-plan).


`planNestedVisibleImageReplacement` uses the same image policy in immediate-parent coordinates and returns whole top-level group snapshots for Core's atomic `replaceElement`. It rejects composed changes to ancestor group frames because those transforms are not included in the image-only change. No source-package write permission is implied.


`planSharedVisibleImageReplacement` resolves master/layout artwork in its owning domain and produces full shared-definition snapshots for Core's atomic replacement APIs. It uses the same local-coordinate and ancestor-composition guard as slide-group planning; visibility and source-package preservation remain separate concerns.
