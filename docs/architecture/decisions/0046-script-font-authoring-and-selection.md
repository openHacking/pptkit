# ADR0046: Explicit script-font authoring with shared deterministic selection

Status: accepted for experimental authoring/measurement/preview/fresh output, 2026-10-26. Source multi-family projection and native shaping/appearance admission remain open.

Follow-up [ADR0047](0047-source-script-font-profile.md) now admits an explicit source-slot profile. The legacy default and its blockers remain compatible; complete language/supplemental and native behavior are still open.

## Problem and decision

A single Core font family cannot retain independently authored Latin, East Asian and complex-script families. The prior source reader correctly blocks conflicting families rather than pretending they are all Latin, but applications also need a format-independent authoring contract that can express the difference.

Add optional flat `eastAsianFontFamily` and `complexScriptFontFamily` fields beside `fontFamily`. Reuse existing semantic theme roles; no OOXML part, source identity, charset byte or font registry enters Core. Flat fields use existing fieldwise preset/region rules. Undefined inherits and authoring null clears one slot; normalized absent means use default family. Legacy normalized runs remain unchanged. This is an additive IR v2 extension.

Put pure shared Unicode/font selection in Core's cohesive `text/font-segments.ts`, exported as `resolveTextFontSegments`. Layout uses concrete fragments only for measurement. SVG uses the same fragments for native and XHTML paths. PPTX retains original run identities and emits native font slots rather than segmenting editable text into new runs. No new package/dependency or environment font lookup is introduced.

## Policy and failure boundaries

The [API](../../api/script-fonts.md) specifies the exact named script sets, neutral/combining/joiner behavior, UTF-16 offsets, empty insertion result and per-call segment limit. The policy uses Unicode Script/Script_Extensions properties described by [UAX24](https://www.unicode.org/reports/tr24/), but is not a full shaping/bidi algorithm or Office language-selection clone. Newly assigned characters follow the runtime's Unicode version. Language-based Han variants and supplemental script maps remain separate.

Validate new family names/roles at every owning text surface, including closed regional style grammar. Preserve immutable snapshots, copy/JSON values and per-field reset semantics. Measurement providers receive only selected concrete families and fragment coverage text; missing fonts or selection limits skip whole-body measurement and retain authored geometry with located diagnostics. SVG action wrappers remain owned by original runs. CSS font names escape control characters as well as quotes/backslashes.

Fresh native font slots follow the [EastAsianFont](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.eastasianfont?view=openxml-3.0.1) and [ComplexScriptFont](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.complexscriptfont?view=openxml-3.0.1) element model. Existing theme-role values still mean the actual Core heading/body family and therefore retain their Latin theme-reference suffix across slots; they do not silently opt into a different native theme collection.

## Scope and evidence

Eight Core, three Layout and four browser/cross-package groups plus installed runtime/types cover authoring, normalization, validation, Unicode selection, budgets, measure-provider failures, exact native slots, SVG spans, actions and source no-op/edit guards. Test expectations include named script characters and independent browser DOM, not only calls to the new selector. Existing single-family paths remain unchanged in normalized shape and rendered output.

The Parser's current unequal-source-font blocker remains intentional. A future source profile must represent each field and cascade independently, preserve raw provenance and prove source-save behavior before relaxing it. This ADR does not claim source round-trip semantic support, native interactive editing, font availability or complete multilingual layout. Existing TODO parents stay open.

[Independent native evidence](../../testing/script-fonts.md) verifies authored source slots separately from resaved/PDF observations. The synthetic-font control exposes substitution and visually absent CJK despite extractable text, and the producer splits source runs. None of those outcomes is treated as a typography or identity-preservation pass.

[ADR0048](0048-script-font-maps-and-language-selection.md) subsequently adds opt-in map/language selection with warned fresh-output materialization. This ADR's original native one-run slot behavior remains the path for runs without maps.
