# Developer Workflow

This guide describes the expected contributor workflow for the workspace bootstrap phase.

## Prerequisites

- Node.js 20 or newer
- `pnpm`

## Local Setup

```bash
pnpm install
```

## Common Commands

```bash
pnpm lint
pnpm typecheck
pnpm test
pnpm build
```

## Adding a New Package

When a new package is justified by architecture and roadmap scope:

1. Create `packages/<name>/`.
2. Add `package.json`, `tsconfig.json`, `README.md`, `src/index.ts`, and `test/smoke.test.mjs`.
3. Update root `tsconfig.json` project references.
4. Add or update any workspace-level documentation that mentions package inventory.
5. Add the package to repository lint checks if it becomes part of the required initial set.

## Dependency Rules

- `@pptkit/core` should stay free of PPTX package concerns.
- `@pptkit/layout` may depend on `@pptkit/core`.
- `@pptkit/pptx-exporter` may depend on `@pptkit/core` and `@pptkit/layout`.
- avoid reverse dependencies that collapse documented boundaries

## Naming Rules

- package names use the `@pptkit/*` scope
- source entry points live under `src/`
- smoke tests live under `test/`

## Documentation Expectations

Update docs alongside structural changes when:

- package inventory changes
- public API drafts change
- contributor commands change
- architecture, API, or workflow documentation changes materially affect implementation direction
