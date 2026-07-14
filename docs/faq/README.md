# FAQ

## Are the packages published?

PPTKit is in preview development. Packages are published manually when a release
gate is satisfied; the local workflow is `pnpm release:npm`. See [Installation](../getting-started/installation.md)
and [Release Strategy](../migration/release-strategy.md).

## Why can I not push into `presentation.slides` or `slide.elements`?

Collections are readonly snapshots. Use add/insert/move/remove/duplicate methods so Core can preserve document-wide IDs and reference validity.

## Why did my image not appear even though a PPTX was generated?

Inspect `result.warnings`. Asset loading is recoverable: an unreadable URL/path is omitted while valid content is packaged. The default exporter does not read local paths; use the `/node` entry.

## Why does output differ between PowerPoint and LibreOffice?

Both applications interpret parts of OOXML, fonts, and text metrics differently. Use installed fonts, explicit styles, representative render fixtures, and cross-application regression checks.

## Does PPTKit parse or round-trip arbitrary PPTX files?

Not yet. Parser preservation and degradation tiers are architecture/roadmap work and are not current guarantees.

## Does Layout automatically measure or paginate content?

Not yet. Current Layout resolves connector anchors and image fitting. Text measurement, overflow, constraints, and pagination are explicit roadmap items.
