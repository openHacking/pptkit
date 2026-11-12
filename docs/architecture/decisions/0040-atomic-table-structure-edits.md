# ADR0040: Pure atomic table structure transactions

Status: accepted for the existing direct Core table model, 2026-10-26. Full style, source-write and native-interaction gates remain open.

## Problem and decision

Compact rows and merged rectangles cannot be safely edited as independent array splices. Insertion through a merge changes its extent; deleting its anchor must preserve remaining content. Merge/split also need explicit content and style ownership. A failed operation must not leave half-updated dimensions or source mappings.

Expose `editTable(input, operations, { frame, limits? })` from Core. Capture plain bounded data, expand the compact representation into a proven physical partition, apply the whole batch privately, then return a frozen candidate and complete original-anchor mappings. The caller commits via existing `replaceElement`, preserving document root identity and transaction guards. Add no package, format dependency, native cell IDs or hidden document mutation.

Insertion expands only merges crossing the insertion point. Deletion shrinks surviving rectangles and promotes anchor content when necessary. Default merge rejects conflicting content; explicit paragraph append retains rich text/actions and requires compatible presets. Split keeps content at the original anchor and independently copies style into new blanks. Geometry policy must explicitly preserve the frame or fit grid sums. Omitted original heights are resolved once against the original frame.

## Alternatives and consequences

Directly mutating row arrays would bypass geometry/provenance/ownership checks. A stateful table model with new cell IDs would introduce lifecycle and cross-document obligations before source/native target contracts exist. Adding element-specific methods to PresentationSlide would mix table topology with registry ownership. The pure transaction owns grid semantics and composes with the existing commit API.

The result maps original explicit anchors, not implicit/new cells or OOXML identities. Merge then split cannot reconstruct prior cell content/styles; applications can retain the untouched input snapshot for undo. Source edits remain blocked until Parser implements independent structure/reference preservation. Broader border/theme precedence is not fabricated by this API.

## Verification and a discovered defect

Core hand cases, exact quotas, generated inverses and an independent physical-array splice oracle cover the transformation. Browser DOM verifies every emitted physical PPTX continuation and static HTML anchor. Cross-package source imports, actual candidate reopens, Layout sizing/pagination and source-save rejection/retry cover integration.

The integration exposed an existing exporter defect for rowspans wider than two columns: the changing output cursor was incorrectly reused to place later vertical continuation entries. Retain a stable anchor column; exporter regression checks every slot for spans 2/3/4/7 and leading-column variants.

See the [contract/tests](../../api/table-editing.md) and [independent/native evidence](../../testing/table-edit-evidence.json). Three fresh edited-table outputs pass Python physical-grid checks and named LibreOffice headless structure/content observations within a 0.1-point research dimension tolerance. Saved-output PDF confirms remaining/control text, not native interactive editability or source preservation. AUT-05.1/05.2 and IMP-09.3.4 remain open at their full scopes.
