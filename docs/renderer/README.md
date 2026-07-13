# Renderer

PPTKit currently provides editable PPTX export, not a general preview renderer.

The implemented forward path is Core authoring → Canonical IR → Layout Result → PPTX package. A future browser/SVG renderer should consume the same normalized/layout contracts and remain a derived view rather than authoritative editor state.

Renderer work must define fidelity strategy, font and asset resolution, unsupported-feature diagnostics, and visual regression fixtures. See [Rendering Pipeline](../architecture/rendering-pipeline.md), [Preview and Editing](../architecture/preview-and-editing.md), and the [Roadmap](../../ROADMAP.md).
