# ADR0048: Script font maps and explicit language selection

Status: accepted for experimental authoring, measurement, SVG and fresh PPTX materialization, 2026-10-26. Supplemental source import and native typography admission remain open.

## Decision

Three named font slots cannot express Arabic and Hebrew using different fonts inside one author run. Add optional `scriptFontFamilies` to Core run styles as a bounded dictionary of 37 admitted ISO 15924 selection keys. Values reuse existing concrete families and heading/body roles. The public frozen `TEXT_FONT_SCRIPT_CODES` and `TextFontScriptCode` define the admitted set. The map contains no source paths, Office theme collections, font files or environment services.

Undefined inherits; a key null removes that inherited key; whole-map null removes all inherited keys. Normalized maps omit resets and empty maps. Table regional layering retains a whole-map reset barrier even when a later layer supplies only one key. Copy, JSON and immutable author snapshots retain detached maps. This is additive IR v2; documents without maps retain the historical selector/output path.

Core owns pure classification and explicit-language interpretation in `text/font-scripts.ts`; the existing shared selector applies map entries before named slots. Layout resolves concrete fragments for providers without altering author runs. SVG reuses that selector. PPTX materializes the current string into concrete families in all three native slots, coalescing adjacent equal families. `script-font-map-materialized` warns once per text owner, including maps currently unused by text, notes and reusable placeholder defaults. Native run topology and reusable theme/language rules are not preserved by this fresh output.

## Selection policy

The [API](../../api/script-fonts.md#script-font-map-extension) defines the exact aliases, language cases and neutral behavior. Character classification uses ECMAScript Unicode properties, informed by [UAX24](https://www.unicode.org/reports/tr24/). Language parsing uses explicit components from [Intl.Locale](https://tc39.es/ecma402/#sec-intl-locale-constructor), never `maximize()`, host locale or a font lookup. Chinese region-to-variant selection is an explicit PPTKit policy, not a normative Office equivalence claim. Bare zh/yue remain generic Han. Combining marks and neutral context remain deterministic; this is not a shaping or bidi algorithm.

Native [supplemental font records](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.supplementalfont?view=openxml-3.0.1) motivate a transport map but do not establish source precedence. Parser `script-families-v1` deliberately retains its three-slot semantics and existing unsupported findings. A future explicit profile must transport source supplemental rules independently of the current text, with theme scope, blocking, provenance and captured-candidate verification. Resolving only the current text at import would couple a scalar text edit to font/style changes and source-save eligibility.

## Evidence and limits

Eight new Core groups cover all admitted keys, language variants/ambiguity, inheritance/reset barriers, owners, invalid maps, Unicode offsets and budgets. A Layout group checks map-field missing/fallback diagnostics and frozen concrete provider styles. Four cross-package groups check Node/ESM/global parity without network, independent XML/DOM slots/actions/alpha and SVG paths, three fresh generations, source no-op and map-edit rejection, empty insertion and unused maps. Installed runtime/type consumers exercise public map contracts and invalid-key/reset negative types.

No new dependency/package, source writer permission or synthetic native typography pass is introduced. Full script coverage, shared line layout, installed-font coverage, native shaping/bidi and PowerPoint theme/edit behavior remain open. Existing synthetic-font observations remain observed; this ADR records automated source/output semantics only.

Follow-up [ADR0049](0049-source-supplemental-font-projection.md) now implements an explicit source-v2 transport/provenance and captured scalar-save slice. The original default/v1 and native typography gates above remain separate.
