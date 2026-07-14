# Example Workbench

The examples workbench is the internal development surface for quickly inspecting PPTKit capabilities in the browser.

The workbench is intentionally allowed to use application-layer UI tooling such as React and shadcn/ui.
That exception is limited to `examples/dev-app`; `packages/*` should remain framework-agnostic.

## Goals

- keep preview-driven development visible while render and export implementations mature
- browse examples by capability rather than only by business scenario
- retain scenario tags so cross-feature flows can grow into richer regression fixtures later

## Adding a New Example

1. Add a small entry to `examples/dev-app/src/example-registry.ts`; place large, cohesive fixtures in a dedicated module under `src/examples/` and register it from the registry.
2. Choose one primary `feature` that matches the current implementation focus in the repo.
3. Add one or more `scenarioTags` when the example is useful outside a single feature lane.
4. Keep the source payload small and reviewable.
5. Update or add tests if the new example introduces a new fixture shape or report behavior.

Presentation-config examples may set a `size` in points and a solid `background` per slide. The `Export SaaS Hunt Swiss Style` fixture demonstrates a 960 × 540 pt, seven-slide deck made entirely from editable text boxes, rectangles, and lines.

## UI Implementation Notes

- prefer shadcn/ui components over hand-rolled primitives when extending the workbench
- keep the default component styles rather than creating a custom design layer
- reserve framework-specific code for `examples/dev-app`, not shared packages

The feature list is intentionally open-ended.
Seed examples can be reorganized as the project matures, so the workbench contract should stay stable even when the initial categories change.

## Output Expectations

The workbench is intentionally honest about implementation status.

- normalized document output should be inspectable
- preview may be structural rather than visually final
- browser export generates and downloads bytes locally, while server export exercises the Node.js file adapter
- **Preview SVG** renders the applied example through `@pptkit/svg-renderer` and displays every slide locally in the browser

The SVG gallery is a QA demonstration, not an editor. It uses the same authored
presentation as PPTX export and reports renderer warning counts alongside the slides.
- both export paths use the currently applied source; visual preview may still be structural

That honesty makes `pnpm dev` useful now without pretending unfinished subsystems are complete.
