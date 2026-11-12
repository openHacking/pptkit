# ADR0043: Bounded source table-style projection with separate provenance

Status: accepted for experimental direct-region import, 2026-10-26. Complete native/style-matrix/default appearance and source-write gates remain open.

## Decision

Parser owns source style selection, theme scope and native region semantics. Add `tableStyleInheritance` to content reports and three session-wide property/dependency/work limits. Resolve inline, explicit registry and registry-default selections through verified source relationships/content types/MCE. Canonicalize GUID lookup without changing source bytes, reject ambiguous effective identities, and never guess a built-in style definition.

Supported direct fill/line/text declarations become ordinary Core cell and text properties. Unknown explicit paint/font/line declarations block older values, with selected-source diagnostics. Source identities and dependencies remain in Parser. Core's custom `TableStyleInput` is not used to smuggle source GUIDs or recreate a native definition with different rules.

## Shared and separate responsibilities

Reuse the verified target slide/layout/master theme/color-map resolver and ordered color transforms, with an optional owner path map for definitions outside slide XML. Cache region evaluation per theme scope, not globally by style ID. Reuse complete table grid/merge proofs, physical border aggregation, local text reading and existing atomic scalar text-save bindings. Declaration paths identify effective source choices; they are not writer authority.

Source region order is whole, horizontal bands, vertical bands, first/last columns, first/last rows and corners, then local declarations. Bands use physical anchor starts and borders use region/merge bounds. Conditional overlaps remain preview fallback; merged conditional regions have an additional documented approximation. Inconsistent direct perimeter segments explicitly block an inherited logical border instead of uncovering an older line.

## Independent findings

A valid generated native source exposed a real content-type mistake in the initial reader/fixture: the registry type is `presentationml.tableStyles+xml`. The reader now verifies that actual type rather than relaxing validation. Independent native fill observations also exposed four dual-band errors when source order copied the custom authoring horizontal-over-vertical policy. Source now applies vertical bands later; all three recorded controls agree on saved direct fill values.

These observations are bounded, named-producer evidence. PDF text presence is not text appearance or interactive editability. Default native text appearance, style-matrix references, script fonts, complete overlap/merge semantics and PowerPoint validation remain open. See the [contract and evidence](../../api/table-style-import.md), [Microsoft region description](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tablestyle?view=openxml-3.0.1) and [DrawingML primer](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf).

## Save boundary

No-op and admitted scalar text saves preserve original style/theme/relationship bytes. Source styles are not rebuilt from the projection. Unsupported style/geometry/structure mutations still block the full save; restoring the baseline permits retry. Fresh output materializes direct formatting and does not establish native source-style retention.

Ten Parser groups, two browser/cross-package groups, installed consumers, exact limits/cancellation/retry and independent Python/DOM checks cover the profile. Native fill observations remain separate from full cascade/interaction/source-write acceptance.
