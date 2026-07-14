# Renderer

PPTKit provides browser-oriented slide preview through `@pptkit/svg-renderer`.

The implemented preview path is Core authoring → Canonical IR → Layout Result → one
standalone hybrid SVG per slide. Geometry uses SVG primitives; rich text and tables use
XHTML `foreignObject`. Preview output remains derived rather than authoritative state.

The renderer exposes asset resolution and fidelity diagnostics and has Chromium visual
regression fixtures. It does not promise PowerPoint pixel parity, general SVG portability,
PPTX parsing, or browser editing. See the [API](../api/svg-renderer.md),
[Rendering Pipeline](../architecture/rendering-pipeline.md), and [Roadmap](../../ROADMAP.md).
