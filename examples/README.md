# Examples

This directory contains runnable development examples for PPTKit.

## Developer Workbench

Run the internal workbench with:

```bash
pnpm dev
```

The workbench starts at `http://localhost:3210` and is designed for fast UI-level verification of whichever capabilities are most relevant to the current project phase.

The main navigation is feature-first so contributors can quickly isolate one capability at a time.
Each example may also include `scenarioTags` so broader end-to-end flows like `basic-deck` or `pitch-deck` remain reusable as the library grows.

## Fixture Contract

Examples in the workbench should register a single `ExampleDefinition` with:

- `id`
- `title`
- `feature`
- `description`
- `inputKind`
- `source`
- `scenarioTags`
- `expectedCapabilities`
- `status`

Use feature grouping for primary navigation and scenario tags for cross-feature regression coverage.
Do not treat any specific feature list as fixed: the first batch of cases should follow the actual implementation priorities in the repo.

## Current Layout

- `dev-app/`: internal preview server, registry, and workbench UI
- `browser-esm/`: zero-build CDN example for SVG preview and browser PPTX download
