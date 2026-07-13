# Best Practices

Use these current patterns when building decks with PPTKit:

- Define a presentation theme and use semantic color/font roles instead of repeating direct values.
- Use layouts and placeholders for repeated structure; keep slide-local elements for actual slide content.
- Supply image source dimensions when using `contain` or `cover`.
- Use stable slide/element IDs for actions and connectors; never derive references from array positions.
- Treat element order as drawing order and use method operations to change it.
- Run `validatePresentation()` before environment-dependent asset loading when diagnostics should be shown to users.
- Inspect exporter warnings even when a file was generated.
- Keep large binary assets outside Core; Core stores identity and source metadata only.
- Test representative output in both PowerPoint and LibreOffice when fidelity matters.

Text measurement, automatic layout, pagination, and cross-application font portability remain active roadmap areas. Do not assume the exporter silently solves them.

See [Core API](../api/core.md), [PPTX exporter](../api/pptx-exporter.md), and [Testing Strategy](../architecture/testing-strategy.md).
