# Installation

PPTKit packages are not published yet.

This page documents the intended installation model so the project can keep the developer experience stable as implementation begins.

## Planned Package Entry Points

Depending on the use case, developers will install one or more of these packages:

- `@pptkit/core`
- `@pptkit/layout`
- `@pptkit/pptx-exporter`
- `@pptkit/pptx-parser`
- `@pptkit/svg-parser`
- `@pptkit/svg-renderer`
- `@pptkit/cli`

## Planned Package Manager Support

The repository is being prepared for a JavaScript and TypeScript workflow centered on `pnpm`.

Expected installation examples:

```bash
pnpm add @pptkit/core @pptkit/pptx-exporter
```

```bash
npm install @pptkit/core @pptkit/pptx-exporter
```

## Local Development Setup

The workspace bootstrap now supports:

```bash
pnpm install
pnpm lint
pnpm typecheck
pnpm test
```

To compile all current packages:

```bash
pnpm build
```

## Current Status

At this stage, the workspace exists for local development and contributor validation. Public package publication is still deferred.
