# ADR0056 — Read-only notes page source inventory

- Status: accepted for the source-only inspection contract
- Date: 2026-10-27

## Problem

Direct notes body import and independent notes text inheritance already exist. They cannot describe a complete notes page: image/header/footer/date/number placeholders, groups, artwork, geometry and cached fields also belong to the source. Treating these as ordinary authored slide objects would imply unsupported geometry, visibility and field evaluation.

## Decision

Add explicit `inspectPptxNotes` in the Parser, with small public plain-data types. Reuse its OPC/structure/MCE validation and source identity infrastructure. Do not change the default inspector, Core/IR, Layout/Renderer, existing text projection or preservation-save permissions.

Separate ordered page relationship domains from distinct source parts. Scan each actual/listed master once. Record effective and raw-only source observations with one-based raw preorder identities. Source classification and ownership follow selected MCE, including ProcessContent, without granting semantics to unknown drawing/extension islands. Match placeholders only by exact explicit type in the actual master; retain absent, ambiguous and grouped candidates. Do not infer schema defaults or use idx/name fallback. Never merge master prompt text into page body.

Record raw attributes, geometry components, nonvisual identity, header/footer flags, text and stored field cache segments. No inherited geometry, current date, displayed page number, crop/visibility policy, artwork rendering or custom path interpretation is computed. Stable output markers explicitly state unresolved source geometry and no field evaluation/authored mapping.

Share bounded limits across output records, strings, candidate work and both notes/structural findings. Reuse infrastructure limits and cancellation. Capture invocation data, reject unknown/accessor limits and return deeply frozen detached JSON-compatible data. Fail atomically on exhaustion or malformed admitted structure.

## Alternatives and consequences

Expanding Core notes authoring now would require a renderer/evaluator contract not yet proven. Extending the default inspector would impose this additional cost and failure surface on existing consumers. A standalone inventory makes the missing semantics visible while preserving compatibility and package direction; it adds no dependency.

[API contract](../../api/notes-source-inventory.md), [source/native-file evidence](../../testing/notes-source-inventory.md) and focused Node/browser/installed-consumer tests define this slice. Full IMP-03.3 inheritance/output and native field/artwork authoring remain open. This is not a general OOXML validator.

The source model follows the official [notes slide structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides) and retains flags such as [showMasterShapes](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.notesslide.showmastershapes?view=openxml-3.0.1) as observations rather than inventing rendering behavior.
