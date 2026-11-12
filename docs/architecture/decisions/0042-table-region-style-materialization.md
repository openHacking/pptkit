# ADR0042: Materialize format-independent table region styles before Layout

Status: accepted for authoring/effective direct output, 2026-10-26. Native reusable rules, source inheritance/writes and interaction remain open.

## Problem and decision

Per-cell styles cannot express a coherent table design that survives row/column edits. Independently computing header/band/corner defaults in SVG and PPTX would cause disagreement and make empty/merged cells depend on renderer behavior.

Add table-local `TableStyleInput` with whole-table, physical row/column bands, first/last rows/columns, four corners, outer/interior border roles and text defaults. Core resolves the cascade during normalization into effective cell inputs; source authoring rules remain detached and reusable. Use existing theme color/font roles and named local text presets. No Office GUIDs, OOXML types, package dependency or mutable style registry enter Core.

Order: wholeTable, table border roles, column bands, row bands, first/last columns, first/last rows, corners in topLeft/topRight/bottomLeft/bottomRight order, local text preset and cell/content overrides. Strokes/insets merge by property, paint and discriminated unions replace, and later uniform strokes patch all inherited side candidates before per-edge patches. Shared physical-edge conflicts remain Layout's responsibility.

## Grid, edits and bounds

A shared pure bounded compact-grid proof validates rectangles and materializes implicit blank slots. Bands use physical anchor starts, while edge/corner regions test the full merge rectangle's boundary contact. Content is never duplicated; merged interiors stay hidden. A styled table has at most 100,000 physical slots, checked before allocation. Standard Core exhaustive validation reports invalid region grammar and metrics through located diagnostics. This adds no asynchronous resources or hostile-Proxy sandbox claim.

Core `editTable` retains the table-local rules, so normalization recomputes them against changed topology. Layout receives effective styles before measuring/paginating; repeated header/source body row formatting is retained across fragments rather than recomputed against each page. Existing unstyled tables retain their established path.

## Native boundary and alternatives

`styleSource: "materialized-regions"` marks canonical effective formatting. Exporter emits located `table-style-materialized` and writes direct native cells/theme references. It cannot reconstruct a custom native style definition with identical cascade semantics, nor can Parser infer original authoring rules from direct formatting. Copying effective styles after each authoring edit would freeze old band assignments; implementing Office-specific rules in Core would violate format independence. A generic document-wide style registry is deferred until its identity/inheritance/source contracts are defined.

Source saves continue to reject rule/style mutations. Fresh reimport's effective values are tested separately from authoring-rule retention. The native runner records fill/text resaves only; prior border width/alpha/segment/transform differences and native interactive editing gates stay explicit. See the [API and independent evidence](../../api/table-style.md).

## Verification

Eight Core groups include 160 geometric-partition oracle cases and exact grid-cap boundaries. Four cross-package groups verify browser forms, independent DOM/native output fields, three RGB rounds, theme role reuse, edit/pagination behavior and source-save guards. Python independently reads native source XML and deliberately rejects corrupted fills; three named headless resaves and saved PDFs provide scoped fill/text evidence. No full Office cascade or native reusable rule/interaction completion is claimed.


## Authoring format-origin follow-up

[Table format origins](../../api/table-format-origins.md) now carry a declaration sidecar through the actual raw region/preset/cell merges, before default normalization. A dedicated raw-input trace records selected declarations, atomic union replacement and font-map null barriers; the existing normalized selector consumes those declarations and emits a separate immutable report. Neither metadata layer enters IR or Parser source provenance. The bounded grid retains each anchor's authored occurrence index, rather than inferring it from object identity or inventing an index for implicit cells.

Paint/stroke/text normalization records supported cell leaves without changing its value path. Uniform-stroke copies and sparse edges preserve the actual selected source. These are logical cell candidates; Layout's independent shared-edge winner remains authoritative for visible borders. [Verification](../../testing/table-format-origins.md) includes direct source reads, independent geometric selection, repeated-object/implicit controls and browser XML/SVG. Native reusable rules, source style writes and full parent gates remain open.
