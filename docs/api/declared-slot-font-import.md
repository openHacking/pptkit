# Declared-slot supplemental font import

Use `textFontProfile: "script-families-v3"` with `projection: "content"` for conservative supplemental rule admission. See [ADR0051](../architecture/decisions/0051-declared-slot-supplemental-projection.md), the [example](../examples/declared-slot-font-import.ts) and [source evidence](../testing/declared-slot-font-import.md). Existing default, `script-families-v1` and `script-families-v2` behavior remains unchanged.

V3 resolves supported raw Latin/EA/CS declarations and source supplemental dictionaries like v2, then selects each script through its own **effective native bucket** after inheritance. It does not borrow the Latin dictionary for an absent EA/CS declaration. The admission is independent of current text, language and installed fonts, so admitted scalar edits do not change source-style interpretation merely by changing scripts.

| Effective source choice | V3 result |
| --- | --- |
| Matching `+mj/+mn-lt/ea/cs`, direct or inherited | Keep that bucket's admitted supplemental entries. |
| Matching theme reference with empty generic theme slot | Generic family stays unsupported; valid specific script entries can still be mapped. |
| Concrete literal or deliberate cross-slot reference | Use the concrete slot value; no supplemental substitution on the destination bucket. |
| EA/CS absent after the whole cascade, Latin dictionary contains corresponding rules | Record those rules as `unsupported`, without a `value`; do not put them in Core's map. |
| Explicit table collection/fontRef | The selected collection owns its bucket choices; its own supplemental rules remain eligible, with required generic-slot blockers retained independently. |
| Unknown/bad/duplicate/MCE/near-theme boundaries | Same conservative source parsing and per-key/component blocking as v2. |

For example, native `latin=+mn-lt` alone with Hant and Grek entries projects only the Latin-bucket Grek rule. `run.scriptFontFamilies.Hant` is a located unsupported property; its source points to the actual Hant font's `typeface`, dependencies include the actual Latin selector, and its target identifies the affected run. The report does not invent an EA attribute or substitute an installed font. A located import finding explains the missing effective selection. The property/dependency/finding limits apply to both mapped and unsupported records.

The profile covers the existing drawing/table readers and, when separately selected, `notesTextProfile: "master-style-v1"` with its independent notes master/theme. Direct-only notes retain their existing boundary without that notes option. Empty insertion text follows the same selection rules. Source font rules remain outside Core; only admitted family strings enter `scriptFontFamilies`.

Byte and Node file imports capture the profile before asynchronous work. `ImportReport.textFontProfile` records v3. Actual source-save candidates reimport with the captured profile and full canonical snapshot equality. The existing text save profile remains required; v3 alone grants no write permission. Font/map/language/style changes remain atomically blocked, including edits mixed into an otherwise permitted scalar text change.

No-op and scalar source saves retain original slot/theme XML. Fresh export materializes supported effective maps with the existing `script-font-map-materialized` warning; it does not preserve source rules or reconstruct unresolved selection. Core's generic fallback is still available for a preview, and callers must consult unresolved source findings when assessing native fidelity.

The [native observations](../testing/real-fonts.md) motivate this correction but do not establish universal Office behavior. V3 does not guess implicit EA/CS fonts, rewrite language tags, infer a Han region from bare `zh`, certify glyph coverage/shaping/bidi, or provide native table/notes/theme editing. Source font writers and the full font/native TODOs remain open.

[Native table observations](../testing/native-table-fonts.md) distinguish declared-rule transport from actual PowerPoint selection: 36 controls retain their declarations while 16 use different native faces. Explicit collection supplemental rules and short Korean tags therefore do not carry a native appearance guarantee. This evidence does not change v3 or earlier-profile compatibility or admit source font writes.

## Native appearance after editing

[ADR0053](../architecture/decisions/0053-declared-fonts-and-native-appearance.md) keeps declared source transport separate from application font selection. [15 native table append controls](../testing/native-table-font-edits.md) retain both run maps without local font overrides through three native save/reopen rounds; the existing native differences remain. This evidence grants no font/style write permission and does not cover bold/italic, theme changes, independent notes or general native typography parity.

The subsequent [native table style slice](../testing/native-table-font-styles.md) covers named bold/pure-italic/regular operations, with source maps retained and regular raster restoration. Independent notes and parent gates remain open; the theme slice is recorded below. Native styling does not authorize PPTKit source style writes.

The [Office table theme slice](../testing/native-table-font-themes.md) verifies actual major/minor relationship response, unchanged explicit collections and source-patched local overrides across three native reopen rounds. Declared maps follow actual theme sources while PDF faces, subset GIDs, ToUnicode and baseline shifts are recorded separately. No source font/style writer or independent notes/native typography guarantee is added.

The [independent native notes baseline](../testing/native-notes-fonts.md) covers two source themes with actual notes-domain relationships and three native reopen rounds. Six of 30 declarations still differ from native selection; empty/bad EA slots and short Korean tags remain explicit. Native notes editing/theme changes and full page output remain open; the baseline grants no source style writes.
