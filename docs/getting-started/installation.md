# Installation

PPTKit is in pre-release development and its packages are not published to npm. Use the monorepo workspace for current development and evaluation.

## Requirements

- Node.js 20 or newer
- pnpm 10.13 or compatible pnpm 10 release

## Workspace setup

```bash
git clone <your-pptkit-repository-url>
cd pptkit
pnpm install
pnpm build
pnpm typecheck
pnpm lint
pnpm test
```

Start the local examples workbench with:

```bash
pnpm dev
```

The workbench is served at `http://localhost:3210`.

## Current package entry points

| Import | Runtime |
| --- | --- |
| `@pptkit/core` | Browser and Node.js |
| `@pptkit/layout` | Browser and Node.js |
| `@pptkit/pptx-exporter` | Browser-neutral generation; URL assets |
| `@pptkit/pptx-exporter/node` | Node.js generation/file output; URL and path assets |
| `@pptkit/cli` | Node.js command line |

## Intended published installation

After the first preview release, the standard installation will be:

```bash
pnpm add @pptkit/core @pptkit/pptx-exporter
```

or:

```bash
npm install @pptkit/core @pptkit/pptx-exporter
```

These commands describe the intended package split; they are not expected to succeed before publication.

Continue with the [Quick Start](quick-start.md).
