# ADR0050: Explicit notes-master text inheritance

Status: accepted for experimental `notesTextProfile: "master-style-v1"`, 2026-10-26. Full notes-page appearance and native editing/typography remain open.

## Decision

Notes are a separate text/style domain. A slide theme must not resolve a notes font reference merely because both objects belong to the same presentation. Add a separately captured content-only notes profile; default import and font v1/v2 alone retain direct-notes behavior. The new profile reuses current text/font/color transformations, with only source-owned inheritance and provenance changes. No Core model, package or dependency is added.

The effective cascade is explicit notes-master `notesStyle`, type-matched notes-master body placeholder's `lstStyle`, local notes `lstStyle`, paragraph defaults and run/end properties. Parent prompt paragraphs are not copied or used as scalar contents. Only an existing unique direct local body placeholder remains the admitted speaker-note owner; grouped/local ambiguous or missing local bodies retain existing admission boundaries.

Microsoft's [PowerPoint placeholder inheritance note](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/3ec954b2-37a6-41da-8973-04a592c91fb2) specifies matching notes placeholders by **type**, unlike slide-to-layout index matching. A bounded private notes-master scan follows recognized drawing containers, verifies drawing identities/placeholder indices and retains only a candidate/count. Multiple body types, grouped candidates or no candidate produce findings and do not choose arbitrary placeholder defaults. The independent global notesStyle can still apply. Source ids/paths remain Parser-owned.

## Independent theme domain

Extend the existing theme resolver with an explicit notes scope. The actual notes-part→notesMaster→theme relationships select its components; the presentation's listed notes master is not substituted when the actual link is absent. Notes/master overrides use component replacement; notes clrMapOvr and notes-master clrMap resolve theme color roles. No slide/layout/master palette or font is borrowed. A missing/unsupported nearer component or slot blocks its own value instead of exposing an older scope.

Reuse the selected font profile: default single-family behavior, independent v1 slots or [v2 supplemental maps](0049-source-supplemental-font-projection.md). New `notes-master-text-style` and `notes-master-placeholder` provenance levels retain actual declarations and dependencies. Placeholder-specific properties additionally retain local/master type-match dependencies. Private reader/audit state is scoped even when the font profile is absent; two notes parts evaluating a shared master under different overrides cannot taint each other.

## Save, resources and evidence

The profile marker is opt-in. Bytes and Node files capture it before asynchronous work, and actual source-save candidate projection carries it into full Core snapshot comparison. Existing scalar save permissions remain separate; style/font/map/language changes still block atomically. No-op and admitted scalar changes preserve original master/theme and source artifacts. Fresh output materializes effective note styles; it cannot reconstruct original notes-master associations, supplemental dictionaries or notes-page artwork.

Existing cumulative text property/dependency, XML/MCE/report, cancellation and snapshot limits apply. `maxNotesStyleWork` separately bounds opt-in master/style/theme/audit evaluation across target scopes; this prevents shared-master amplification that a per-source XML bound alone would not prevent. Direct-only modes do not consume it. Matching scans are bounded by source XML and checkpointed, cached once per actual notes master; the type match does not construct one candidate list per child. No new unbounded registry or host font lookup is introduced.

[Thirteen Parser groups](../../../packages/pptx-parser/test/notes-style-import.test.mjs) and [three cross-package/browser groups](../../../scripts/notes-style-import-parity.test.mjs) cover dialects, old modes, matching, overrides, missing/ambiguous/grouped cases, component/scope isolation, MCE, identities, levels/empty insertion, captured bytes/file/candidate options, budgets, source preservation and warned fresh output. Packed consumers exercise the options and provenance levels. [Independent Python evidence](../../testing/notes-style-import.md) reads actual relationships, type attributes, font/color/style values and immutable payloads in four saved stages, with wrong-font and wrong-placeholder-type controls.

These checks establish the declared bounded text projection. They do not certify notes-page geometry, artwork, headers/footers/fields, parent prompt paragraph formatting, complete font precedence/shaping, native application editing or broad round-trip. Those remain separate parent gates.

## Source-conformance clarification

[Microsoft's PresentationML structure specification](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document#notes-master-part) permits at most one Notes Master part per package. The existing second-unlisted-master tests are non-conforming diagnostic source controls. They verify actual-edge lookup and preservation with an `unlisted-master` finding; they do not admit multi-notes-master authoring or standard/native round-trip. Two listed IDs still reject before a usable session is returned. A single master with per-notes themeOverride is the separate cardinality-correct source case; [interleaved A/B/A evidence](../../testing/notes-master-isolation.md) verifies its fontScheme scope and distinguishes absent override fallback from present-component replacement. This clarification changes no runtime behavior, profile defaults or writer permissions.

## Named native admission evidence

The [Office-theme experiment](../../testing/native-notes-font-themes.md) verifies separate A/B notes themes and a single master with distinct per-notes fontScheme override parts through three native rounds. Native renumbering is followed by actual relationships. The source control sharing one override part across fifteen notes triggers a repair prompt in the named PowerPoint build, while the distinct-part variant opens without repair. No repair was performed; this observation neither proves a general schema violation nor changes the declared reader/source-save contract. Header kerning and ordinary-slide Y shifts remain explicit, and complete notes-page reconstruction remains open.
