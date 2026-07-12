# Testing Strategy

This document defines the testing approach for the workspace bootstrap and the first implementation phases.

## Goals

- Keep the initial feedback loop simple.
- Validate package boundaries early.
- Add richer regression coverage only when the output becomes meaningful.

## Test Layers

### 1. Package Smoke Tests

Every initial package should include a smoke test that proves:

- the package builds in isolation
- the public entry point can be imported
- the provisional runtime contract behaves as expected

Bootstrap examples:

- `@pptkit/core` can create a presentation shell
- `@pptkit/layout` can consume a core document
- `@pptkit/pptx-exporter` returns a structured placeholder result
- `@pptkit/cli` responds to `--help`

### 2. Workspace Validation

Root-level commands should remain stable:

- `pnpm lint`
- `pnpm typecheck`
- `pnpm test`

CI should run the same commands contributors run locally.

### 3. Future Integration Tests

Once export behavior becomes real, add integration coverage for:

- authoring to export flow
- file creation and output paths
- unsupported feature diagnostics
- error handling around invalid document input

### 4. Future Golden and Fixture Tests

Golden-file coverage is appropriate only after the exporter produces stable output.

At that stage:

- fixtures should live near the owning package
- snapshot artifacts should be small and reviewable
- every golden file should correspond to a named behavior or regression

## Directory Guidance

- package smoke tests live under `packages/<name>/test/`
- future package fixtures should live under `packages/<name>/test/fixtures/`
- cross-package examples belong in `examples/`

## Review Rules

When real implementation work begins:

- add tests for every new public API or boundary behavior
- prefer small focused fixtures over giant all-in-one decks
- treat parse/export regressions as high-priority coverage candidates

