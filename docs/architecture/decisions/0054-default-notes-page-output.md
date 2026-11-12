# ADR0054: Default fresh notes-page output

Status: accepted for bounded fresh Exporter output, 2026-10-26. General notes-page authoring/import remains open; the initial empty insertion-style defect has a separate correction below.

## Decision

Keep the existing 540×720pt notesSize and public `slide.notes` contract. Exporter owns a fixed default notes page: a slide-image master placeholder fits the actual final Layout slide aspect into the upper 468×288pt area at (36,36), centered on both axes; a body placeholder occupies (36,360,468,324). Both master frames have explicit rect geometry. This corrects native blank-image output without introducing notes geometry or OOXML in Core/IR/Layout.

Every local notes page contains a matching slide-image placeholder. Its geometry inherits from the master and it has no artificial text body. Existing local notes-body drawing ID2 and first-shape order remain stable; the new local image uses ID3 after the body. Absent/plain default notes omit the local body, while authored empty formatting/actions retain it. Master prompt text is empty and is never projected into speaker notes.

The existing per-slide notes relationships remain authoritative: a unique notes part points to its own slide and the single shared notes master. Same-document copies receive separate slide/notes parts. Slide actions remain notes-local relationships and no image relationship or embedded screenshot is introduced: the native application renders the owning slide.

Default notes output changes intentionally; public types, ordinary slide geometry, imported source no-op/scalar writers and browser slide preview do not change. Research source assemblers explicitly replace the default placeholders before constructing their independent controls, preventing duplicate identities or accidental coupling to fresh-output defaults.

## Evidence and remaining work

[Cross-package regression](../../../scripts/notes-page-parity.test.mjs) checks source DOM/OPC geometry, no overlap, aspect, copied ownership, absent/empty distinctions, browser global bytes and three fresh projection/no-op rounds over four slide sizes. [Native evidence](../../testing/authored-notes-pages.md) covers three unpatched public-author inputs, twelve first-open pages and 36 pages across nine Save As/close/exact-path reopen/Notes-print cycles.

The empty 12pt bold run survives fresh output/projection but PowerPoint deletes it on first native save. The evidence records that failure instead of treating an empty raster as styling proof; insertion-property serialization needs a separate correction. Custom notes layouts, arbitrary imported artwork/fields, complete inheritance, notes preview/overflow/pagination, source style writes and other applications remain separate tasks.


[ADR0055](0055-empty-paragraph-insertion-properties.md) subsequently corrects the all-empty insertion-style defect, with three empty and three actual-typing native rounds. The preceding loss observation remains true for the original captured inputs; those files and their manifest were not changed.
