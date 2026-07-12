# Example Workbench

The examples workbench is the internal development surface for quickly inspecting PPTKit capabilities in the browser.

## Goals

- keep preview-driven development visible before full render and export implementations exist
- browse examples by capability rather than only by business scenario
- retain scenario tags so cross-feature flows can grow into richer regression fixtures later

## Adding a New Example

1. Add a new entry to `examples/dev-app/src/example-registry.ts`.
2. Choose one primary `feature` that matches the current implementation focus in the repo.
3. Add one or more `scenarioTags` when the example is useful outside a single feature lane.
4. Keep the source payload small and reviewable.
5. Update or add tests if the new example introduces a new fixture shape or report behavior.

The feature list is intentionally open-ended.
Seed examples can be reorganized as the project matures, so the workbench contract should stay stable even when the initial categories change.

## Output Expectations

The workbench is intentionally honest about implementation status.

- normalized document output should be inspectable
- preview may be structural rather than visually final
- export may remain placeholder until file generation is implemented

That honesty makes `pnpm dev` useful now without pretending unfinished subsystems are complete.
