# Install PPTKit

This guide is for application developers installing PPTKit to build presentations.
For contributing to PPTKit itself, use the [Developer Workflow](../guides/developer-workflow.md).

PPTKit is in pre-release development. The package API may change before the stable release.

## Requirements

- Node.js 20 or newer
- pnpm 10.13 or compatible pnpm 10 release, or npm

## Install the packages

```bash
pnpm add @pptkit/core @pptkit/pptx-exporter
```

or:

```bash
npm install @pptkit/core @pptkit/pptx-exporter
```

## Package entry points

| Import | Runtime |
| --- | --- |
| `@pptkit/core` | Browser and Node.js |
| `@pptkit/layout` | Browser and Node.js |
| `@pptkit/pptx-exporter` | Browser-neutral generation; returns PPTX bytes and loads URL assets |
| `@pptkit/pptx-exporter/node` | Node.js generation and file output; loads URL and local path assets |
| `@pptkit/cli` | Node.js command line |

For a typical application, import `@pptkit/core` to create and validate a
presentation. Use `@pptkit/pptx-exporter` when the application needs bytes for a
browser download, or `@pptkit/pptx-exporter/node` when a Node.js process should write
an editable `.pptx` file.

## Next step

Continue with the [Quick Start](quick-start.md) to create and export your first deck.
