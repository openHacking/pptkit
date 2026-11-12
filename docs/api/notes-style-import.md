# Notes-master text inheritance

Content import can opt into `notesTextProfile: "master-style-v1"`. It resolves supported inherited text styles for the existing admitted notes body, under its independent notes theme. Default import and `textFontProfile` alone retain direct-notes behavior. Source scalar-save permissions, native notes-page layout and style writes remain separate.

See [ADR0050](../architecture/decisions/0050-notes-master-text-inheritance.md), the [typed example](../examples/notes-style-import.ts), [direct notes import](notes-import.md), [font slots](script-font-import.md) and [supplemental rules](supplemental-font-import.md). `ImportReport.notesTextProfile` records the opt-in mode. Invalid values and structure projection reject before file/package work. Both byte and Node file imports capture the option before their first asynchronous work.

## Owner and inheritance selection

The existing owner rule remains: one direct local notes shape with an explicit body placeholder and a local text body. Arbitrary names/order, nested bodies, multiple local bodies or absent local text do not synthesize speaker notes from the master. Placeholder/text/frame/page identities stay outside Core.

| Order | Supported source layer |
| --- | --- |
| 1 | Actual notes master's `p:notesStyle` default paragraph and selected list level |
| 2 | Unique type-matched notes-master body placeholder's `a:lstStyle` defaults/level |
| 3 | Local notes body's list defaults/level |
| 4 | Local paragraph properties/default run |
| 5 | Local run properties; end properties only for an empty insertion run |

Notes placeholders match the master by `type="body"`, **not** by idx. Different local/master indices can match. The [PowerPoint implementation note](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/3ec954b2-37a6-41da-8973-04a592c91fb2) documents this distinction. Ambiguous types, grouped/non-shape candidates or a missing body match produce located findings and skip placeholder-specific defaults; independent global notesStyle remains eligible. Interpreted notes-master drawing identities and placeholder indices must be valid/unique. Extension islands are not searched for candidates. Effective MCE choices are selected before matching/style evaluation.

Parent prompt paragraphs and notes-page frame/artwork are not copied into Core notes or treated as style parents. Unsupported attributes/properties remain findings. This profile projects the shared reader's supported paragraph/run fields; it does not add a full notes-page document model.

## Notes themes, fonts and provenance

The actual notes-part `notesMaster` relationship supplies the master; its own `theme` supplies the font/palette component. A listed presentation notes master or slide/layout theme is never substituted for a missing actual relationship. A different actual notes master can be inspected even when unlisted, retaining the existing structure finding. Missing/wrong relationships retain existing structural/error semantics.

PresentationML permits at most one Notes Master part per package; see [Microsoft's structure specification](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document#notes-master-part). The reader rejects two listed notes-master IDs. Its diagnosed inspection of an unlisted actual master is tolerance of non-conforming source, not standard multi-notes-master authoring or native admission. [Cardinality and isolation evidence](../testing/notes-master-isolation.md) separates that adversarial source from a single shared notes master with supported per-notes fontScheme overrides. Exact source preservation does not repair or certify the original package.

Notes/master theme overrides replace individual components. A present fontScheme with missing slots/rules blocks an older scheme instead of mixing it with another theme. Notes `clrMapOvr` and notes-master `clrMap` resolve theme role colors. Shared-master source/audit caches are scoped per target notes domain, so unsupported values in one override do not pollute another notes part.

`textFontProfile` remains independent: omit it for legacy single-family projection, choose v1 for separate Latin/EA/CS, or choose v2 for admitted supplemental maps. The new notes profile enables actual notes-theme resolution for each choice. Empty/invalid slots keep unsupported provenance and normal projection fallback; they do not borrow slide fonts. Core's existing map/language selector evaluates retained v2 rules without querying host fonts.

Selected properties appear in `textStyleInheritance` with notes-part target identities. New levels `notes-master-text-style` and `notes-master-placeholder` identify master declarations; local levels retain their existing names. Actual theme/color-map/script dependencies remain in `resolvedFrom`. Placeholder-specific fields additionally record local/master type attributes; those prove the match independently of ids or indices. No raw source master/theme identity enters Core.

## Saving and fresh output

No-op retains the exact original bytes. Existing admitted scalar text profiles can save bound notes characters while retaining source master/theme/collection/relationship payloads. Actual-candidate reimport carries both captured notes and font profiles and compares the full approved Core snapshot. It does not remove inherited font/map/color fields to make validation pass.

Changing style/font/map/language, owner shape or paragraph/run topology remains atomically blocked. Restore original style for retry; enabling notes inheritance grants no new writer permission. Empty insertion style resolution does not create an editable source text leaf.

Fresh PPTX output serializes effective local notes styles. Script maps use warned materialization and can change run topology; original master/theme associations and notes-page artwork are not rebuilt. Slide SVG does not render speaker notes. Source preservation, fresh text-style output and complete native notes-page appearance remain distinct.

## Limits and evidence

Notes and drawing provenance share cumulative `maxTextStyleProperties` and `maxTextStyleDependencies`. The opt-in profile also has cumulative `maxNotesStyleWork` (default 1000000) across master matching, inherited list/audit and notes theme/font/color evaluation. Reusing a master under distinct target overrides consumes work for each actual evaluation; cached work is not repeated. Exceeding it rejects the import without a usable partial session. Direct-only notes imports do not consume this new profile budget. Notes-master matching scans are checkpointed and cached per actual part, bounded by existing XML/package limits; they do not expand master text into every local note. General MCE/report/snapshot/save/cancellation limits remain active. Failed imports produce no partial usable session and independent retry is supported.

[Thirteen Parser groups](../../packages/pptx-parser/test/notes-style-import.test.mjs), [three browser/cross-package groups](../../scripts/notes-style-import-parity.test.mjs) and packed runtime/types cover both dialects, old modes, source levels/choices, type-vs-index, overrides/blocks/scopes, exact budgets, failure/retry and three source saves. Independent DOM checks actual paths, note themes and emitted current font fragments. The [Python evidence](../testing/notes-style-import.md) verifies four actual source packages and rejects wrong-font and wrong-type mutations with two checks each.

The [independent native notes baseline](../testing/native-notes-fonts.md) now verifies two real-family source themes, type-vs-index matching, local/empty/bad slots and three PowerPoint save/reopen rounds with actual Notes-print PDFs. It records 24 matches and six declared/native differences; source-normalized theme relocation and redundant declarations remain explicit. The bounded [native append/local-style](../testing/native-notes-font-edits.md) and [Office-theme/independent-override](../testing/native-notes-font-themes.md) controls now have three actual native rounds. Notes font sources remain independent of the changed slide theme; named header kerning and ordinary-slide Y shifts remain recorded differences. The shared override-part source triggers a repair prompt in the named PowerPoint version; distinct per-notes parts open without repair. This does not classify all such shared sources as schema-invalid or change source import/preservation permissions. Complete shaping/bidi, page geometry/fields/artwork and source style writers remain unverified or unsupported. The specimen retains blank slide-image placeholders; this baseline does not certify full notes-page output. Complete notes and font parent tasks remain open.
