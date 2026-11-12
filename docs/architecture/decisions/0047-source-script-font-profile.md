# ADR0047: Explicit source script-font profile with independent cascade and saved-candidate reprojection

Status: accepted for experimental `script-families-v1` content import, 2026-10-26. Language/supplemental selection, complete native appearance and source font writes remain open.

## Decision and compatibility

Add explicit `textFontProfile: "script-families-v1"`, captured at byte/file invocation and recorded on the content report. Keep the legacy single-family projection unchanged when absent. Reuse Core's existing optional family fields and Parser's property records; extend the property union with East Asian/complex-script keys and a notes-direct source level. No new package, registry, dependency or OOXML/Core identity coupling.

Source run/list/paragraph declarations select three fields independently. Theme references resolve explicit lt/ea/cs slots in the target's verified nearest font component. Empty/missing/unsupported higher declarations block only that field; local Latin does not discard inherited alternate families. An explicit table collection replaces its collection choice, with missing slots blocked, and a reference resolves the three named theme slots. Unrepresentable supplemental selection stays unsupported rather than being silently discarded.

Inherited reader consumption/diagnostic state is isolated per target theme scope in the new profile. Shared source declarations can have unsupported results under one slide and valid results under another without contaminating audit findings. Selected missing/empty theme slots can retain the actual component/slot dependency. Existing source paths, MCE processing, budgets and cancellation remain authoritative.

## Notes and save boundary

Expose direct notes font provenance through the same bounded text-property collector. Do not invent notes master/theme inheritance or substitute the slide's theme. Notes identities/source binding remain in their existing ownership model.

Propagate the captured profile through session edit state and actual-candidate reprojection. This matters even for character-only changes: reprojecting with the legacy profile would erase independently mapped font fields and either reject a correct candidate or weaken snapshot validation. Keep the complete state comparison and exact source-package validation; do not mask font fields to make a save pass.

The projection profile does not grant write authority. Only existing admitted scalar characters can change; font/style/structure changes still block atomically and can retry after restoring baseline. Fresh materialized font output is a separate operation from source preservation.

## Evidence and remaining work

The [API](../../api/script-font-import.md) links fourteen Parser groups, three browser/cross-package groups and installed consumers. Tests exercise both dialects, independent source DOM values, per-field overrides/blockers, themes/notes/tables, scope isolation, exact budgets, three source/fresh rounds, capture/cancellation and rejected font mutations. Candidate validation keeps its full snapshot comparison and the selected profile.

This admits explicit slots only. Language/Han variants, supplemental script maps, full notes inheritance, actual shaping/font coverage and PowerPoint editing require separate evidence. Existing native synthetic-font observations expose substitutions and invisible CJK despite extractable text; a successful source projection or scalar preservation is not a native typography pass.

[Independent source/native evidence](../../testing/script-font-import.md) verifies four packages' source declarations, scalar text and exact bytes outside admitted character spans, and rejects a deliberately wrong ea value. Native/PDF run splitting and missing visual CJK are recorded as observations rather than passes. Theme collection and supplemental distinctions follow Microsoft's [MinorFont](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.minorfont?view=openxml-3.0.1) and [SupplementalFont](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.supplementalfont?view=openxml-3.0.1) model; the explicit-slot profile does not infer language selection from those extra entries.
