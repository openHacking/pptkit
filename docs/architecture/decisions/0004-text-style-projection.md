# ADR 0004: Bounded text style projection during PPTX import

Status: accepted for the bounded experimental Parser profile; full text/theme and template semantics remain pending.

Date: 2026-10-25.

## Problem

Text in a slide can omit properties that come from paragraph list styles, matching placeholders or master named styles. Copying an ancestor text body would duplicate prompt text; flattening an entire style object would erase which explicit local values clear inherited properties. Unsupported high-priority declarations must also stop older values from appearing effective.

## Decision

Parser resolves only supported text properties and materializes their effective values into ordinary Core frame, paragraph and run styles. Per-property provenance stays in the Parser `ImportReport`; Canonical IR remains format independent. An additive optional `colorOpacity` run field carries alpha; absence means opaque, and normalization omits opaque values to preserve existing IR v2 snapshots.

The bounded order is: master named text style; presentation default for plain shapes; matched master placeholder list; matched layout placeholder list; local list defaults and the exact selected level; direct paragraph/run declarations. Frame values resolve separately from matching master placeholder to layout placeholder to local body properties. Paragraph level is selected from the source `lvl` integer (default zero) and maps only to the corresponding `lvl1pPr` through `lvl9pPr`. Prompt paragraphs/runs never supply content or property values.

Each explicit declaration maps independently. Supported zero/false/no-style declarations clear lower values. An unsupported declaration becomes a blocker for that property; lower-priority values do not leak through. Theme color references resolve through the verified master theme, source theme override and applicable color-map override. The supported profile evaluates RGB/system fallback palette entries and the [named text-color transforms](../../api/text-color-transforms.md), recording palette/map/transform dependencies. A fill replaces both RGB and alpha, including opaque reset; an unsupported fill blocks both properties. Major/minor Latin theme font references resolve to concrete family names; East Asian/complex script font choices that cannot fit Core's single font-family property block that property. The source report records the selected mapped value or blocker, including source/target locations, level and `resolvedFrom` dependencies. Unselected properties continue to use Core projection defaults and are not claimed as source appearance.

Color transforms outside the named profile, non-Latin theme font mapping, complex effects, full Office layout, note/master cascade, dynamic template binding, and source edit-save are outside this profile. Unknown or unsupported declarations receive source-located findings. Separate `maxTextStyleProperties` and `maxTextStyleDependencies` limits bound selected records and their source dependencies.

Theme resolution uses the target slide's verified slide/layout/master chain rather than the location of the style declaration. Color and font schemes inherit independently through partial theme overrides; a present component replaces the older component, including unsupported/missing-slot blockers. Component selection is cached within each target scope, while parsed theme parts are shared within the import. Typed theme/override roots and content types are checked before use. Parser retains all source paths and does not replace Core's document theme with one arbitrarily selected source theme.

## Rationale and limits

The profile encodes explicit contracts visible in the source and avoids inventing values for unsupported encodings. The exact-level rule and plain-shape presentation default precedence are intentionally narrower than a claim of complete Office property resolution. They require continued fixture review and native producer validation before expanding support.

Ancestor prompt text is not copied. Notes/table bodies share local text parsing but do not receive this shape ancestry cascade. This keeps the implementation in Parser and avoids moving OOXML-specific source identity into Core.

## Verification and remaining gates

See the [public contract](../../api/text-style-inheritance.md), independent parser fixtures, three-round fresh export checks, and installed Node/browser consumer checks. These establish deterministic projection and public API behavior; they do not establish PowerPoint/LibreOffice visual equivalence, editing fidelity or template rebinding. IMP-05.1, AUT-07 and IMP-09 remain open.
