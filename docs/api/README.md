# API Reference

This section documents the public package contracts currently implemented by PPTKit. It describes callable APIs, inputs, results, defaults, validation, and runtime boundaries.

## Packages

- [`@pptkit/core`](core.md) — authoring state, elements, assets, themes, layouts, validation, and Canonical IR.
- [`@pptkit/layout`](layout.md) — connector and image geometry resolution over normalized documents.
- [`@pptkit/pptx-exporter`](pptx-exporter.md) — browser-neutral PPTX generation and Node.js file output.
- [Package overview](package-overview.md) — package responsibilities and entry points at a glance.

`@pptkit/cli` remains a minimal workflow shell and does not yet expose a stable presentation-authoring command surface. Planned parser and renderer packages are documented as architecture direction, not current APIs.

## Reading order

1. Start with the [Quick Start](../getting-started/quick-start.md).
2. Use the [Core overview](core.md) to find the relevant authoring topic.
3. Read [Layout](layout.md) when working with resolved geometry.
4. Read the [PPTX exporter](pptx-exporter.md) for runtime, asset-loading, and output behavior.

All examples use the current pre-release contract. PPTKit does not maintain documentation for superseded unpublished APIs.
