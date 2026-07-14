# Developer Workflow

This guide is for PPTKit contributors and maintainers working in the TypeScript
monorepo. It is not required for application developers using the published packages;
start with [Install PPTKit](../getting-started/installation.md) instead.

## Prerequisites

- Node.js 20 or newer
- pnpm 10.13 or compatible pnpm 10 release

## Local Setup

```bash
git clone <your-pptkit-repository-url>
cd pptkit
pnpm install
```

Build the workspace before running checked examples or the workbench:

```bash
pnpm build
```

## Common Commands

```bash
pnpm lint
pnpm typecheck
pnpm test
pnpm dev
pnpm build
```

`pnpm dev` starts the local examples workbench at `http://localhost:3210`.

Run the repository's checked Quick Start example with the workbench TypeScript runner:

```bash
pnpm --filter @pptkit/dev-app exec tsx \
  --tsconfig ../../docs/examples/tsconfig.json \
  ../../docs/examples/quick-start.ts
```

This writes `hello-pptkit.pptx` in `examples/dev-app/`.

## Documentation checks

Documentation links, checked snippets, stale API references, and documentation
example types are validated with:

```bash
pnpm docs:check
```

Run this when changing Markdown or files referenced by a `doc-test` marker.

## Adding a New Package

When a new package is justified by architecture and roadmap scope:

1. Create `packages/<name>/`.
2. Add `package.json`, `tsconfig.json`, `README.md`, `src/index.ts`, and `test/smoke.test.mjs`.
3. Update root `tsconfig.json` project references.
4. Add or update any workspace-level documentation that mentions package inventory.
5. Add the package to repository lint, boundary, documentation, and root TypeScript checks.

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
- public API contracts change
- contributor commands change
- architecture, API, or workflow documentation changes materially affect implementation direction
