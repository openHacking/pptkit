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

## Planned Local Development Setup

Once the workspace is initialized, contributors should be able to start with:

```bash
pnpm install
pnpm dev
pnpm test
```

## Current Status

At this stage, the installation contract is being documented before the packages are published. This helps keep naming, package boundaries, and onboarding materials consistent from the beginning.
