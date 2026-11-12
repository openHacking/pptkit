# Explicit script font families

`TextRunStyleInput` and normalized runs can carry independent `eastAsianFontFamily` and `complexScriptFontFamily` values beside `fontFamily`. Each is a concrete name or the existing `{ theme: "heading" | "body" }` role. This is an authoring, Layout, SVG and fresh PPTX capability. Source reading requires the separate explicit [script-font import profile](script-font-import.md); source font edits remain blocked.

See the [typed example](../examples/script-fonts.ts) and [ADR0046](../architecture/decisions/0046-script-font-authoring-and-selection.md).

## Defaults and inheritance

The original `fontFamily` is the default/Latin family. Omitted script slots use that family. Authoring `undefined` inherits an existing slot from a preset/placeholder/region; `null` explicitly clears that one inherited slot back to `fontFamily`. Changing the default family leaves explicit script slots intact. Normalized IR omits cleared/absent slots and never stores null. Existing documents without the fields retain their previous normalized shape.

Named slots inherit independently through text presets, shape text, table cell/region styles, notes and placeholder normalization. Copy/JSON/immutable author snapshots retain them. A new slot requires a nonblank family name without ASCII control characters, or one valid heading/body role; malformed run values produce located `invalid-font-family` diagnostics. Table regional grammar admits the same fields and retains its existing `invalid-table-style` diagnostics.

## Shared selection policy

`resolveTextFontSegments(text, style, theme, options?)` is a pure Core utility shared by Layout measurement and both SVG text paths. It returns detached `{ text, start, end, family, script }` records. Offsets are UTF-16 positions in the original text; strings are not rewritten, reordered or split inside surrogate pairs. Empty text produces one insertion segment. The result does not mutate the author run or assign new run identities.

With no script slots or nonempty map, the utility retains one historical default-font segment. Otherwise, the current deterministic selection policy is:

| Characters | Selected slot |
| --- | --- |
| Han, Hiragana, Katakana, Hangul, Bopomofo and Yi Script_Extensions | East Asian |
| Arabic, Hebrew, Syriac, Thaana, Nko, Samaritan, Mandaic, Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil, Telugu, Kannada, Malayalam, Sinhala, Thai, Lao, Tibetan, Myanmar, Khmer and Mongolian Script | Complex script |
| Combining marks and joiners following another character | Keep the preceding selected slot |
| Remaining punctuation/separators | Keep the preceding slot; leading neutral text starts with default |
| Other characters, including Latin/Greek/Cyrillic, ASCII digits and emoji | Default |

Newlines use the default slot and reset neutral context. Script-specific punctuation recognized by the first two rows retains that classification. Undeclared slots fall back to `fontFamily`. This is an explicit PPTKit policy based on ECMAScript Unicode properties, not a complete Unicode shaping/bidi implementation or a reproduction of every Office language/font-selection rule. Unicode assignments for newly added characters depend on the runtime's Unicode data. The opt-in map extension below adds explicit language/Han selection; native supplemental theme import and unlisted scripts remain further work.

The default `maxSegments` is 100000, charged before adding each segment. It must be a positive safe integer. Exceeding it throws `RangeError` without returning a partial result. Traversal is linear in code points, output allocation is bounded by the number of segments, and repeated same-script text does not build a per-character output array. This is a per-call segment bound, not a total document/CPU or text-byte budget.

## Measurement, preview and native output

Layout retains original run/style/identity fields. Its optional measurement path sends concrete script fragments through the caller's font resolver and frozen measurement request. The resolver receives each fragment's actual text for coverage checks. A missing script font prevents the entire body measurement; substitutions retain `font-fallback` warnings at the original run's script-field path. Selection failure also retains authored geometry and prevents a partial body measurement request. Measurement fragments contain one resolved family and do not ask the provider to interpret script slots again.

SVG native tspans and XHTML spans select the same families. XHTML actions continue to wrap one author run, even when it contains several font spans. CSS family names escape quotes, backslashes and control characters. Native SVG wrapping remains heuristic; browser shaping and actual glyph/font availability remain host-dependent. Core's default height estimate does not become actual font measurement merely because slots are present.

Without a map, fresh PPTX writes distinct `a:latin`, `a:ea`, `a:cs` typefaces in one original run, for text/shape/table/notes output. Theme font roles keep `+mj-lt`/`+mn-lt` because those identify the modeled Core heading/body families; changing to `-ea/-cs` would select unmodeled theme script fonts. Native producers may select or substitute fonts differently from PPTKit's preview policy. Actual PowerPoint editing/shaping/appearance requires separate named evidence.

The default Parser still blocks unequal explicit source script families under its existing single-family projection; an explicit [source profile](script-font-import.md) now maps the three fields independently. No-op source saves retain the original slots byte-for-byte. Adding/changing these author fields to an imported session does not grant source font-write permission; unsupported source edits remain blocked atomically.

## Evidence and remaining scope

[Eight Core groups](../../packages/core/test/script-fonts.test.mjs) cover all text owners, independent inheritance/null reset, validation, old IR compatibility, named script controls, combining/joiners, supplementary Han/UTF-16, theme roles and exact segment budgets. [Three Layout groups](../../packages/layout/test/script-fonts.test.mjs) cover provider fragments, original identities, script-specific fallback/missing paths and budget failure. [Four cross-package/browser groups](../../scripts/script-fonts-parity.test.mjs) compare Node/ESM/global authoring/output, independent native XML slots, SVG/XHTML spans, source no-op/edit guards and CSS declaration injection controls. Installed runtime/type consumers exercise public exports.

Complete source font semantics, source supplemental-script selection, complete shaping/bidi/vertical text, bullet font policy, font packaging/availability and full native compatibility remain open under IMP-05.1.3.3.2 and AUT-01.3. This capability does not close those parent tasks.

An [independent Python/native experiment](../testing/script-fonts.md) verifies six authored source runs and rejects a wrong ea family. The named producer splits five slide runs into fifteen, substitutes fonts, and leaves CJK visually absent despite extractable text in the synthetic-font control. These are observations and remaining coverage gates, not typography passes.

## Script font map extension

`scriptFontFamilies?: ScriptFontFamiliesInput | null` adds explicit per-script families. [ADR0048](../architecture/decisions/0048-script-font-maps-and-language-selection.md) records the transport/output tradeoff; the [typed example](../examples/script-font-map.ts) demonstrates mixed Han/Arabic/Hebrew text. This extension is independent of source `script-families-v1`, which still maps only the three named slots.

The frozen `TEXT_FONT_SCRIPT_CODES` and `TextFontScriptCode` admit exactly:

```text
Zyyy Latn Grek Cyrl Hani Hans Hant Hira Kana Jpan Hang Kore Bopo Yiii
Arab Hebr Syrc Thaa Nkoo Samr Mand Deva Beng Guru Gujr Orya Taml Telu
Knda Mlym Sinh Thai Laoo Tibt Mymr Khmr Mong
```

Map values use concrete families or heading/body roles. Undefined inherits per key; key null clears that key; whole-map null clears all inherited keys. Empty input maps inherit. Normalized maps contain neither nulls nor undefined/empty entries; absent maps keep the old IR shape. Unknown keys, non-dictionary roots and invalid families produce located `invalid-script-font-map` diagnostics; regional table grammar uses `invalid-table-style`. This is an admitted subset, not arbitrary ISO 15924 support.

The selector checks an explicit map entry before the corresponding named slot/default. Its optional `scriptFont` result field identifies the actual map key used; it is omitted when no map entry supplied the family. Layout diagnostic paths then end in `style.scriptFontFamilies.<key>`. Providers receive concrete families with all map/slot rules removed.

| Character/key | Candidate order |
| --- | --- |
| Ordinary admitted script | Exact script key |
| Hiragana / Katakana | Hira / Kana, then Jpan |
| Hangul | Hang, then Kore |
| Bopomofo with Traditional language hint | Bopo, then Hant; otherwise Bopo only |
| Han with explicit Hans/Hant script | That variant, then Hani |
| Han with explicit Jpan or language ja | Jpan, then Hani |
| Han with explicit Kore/Hang or language ko | Kore, Hang, then Hani |
| Han with zh/yue and CN/SG region | Hans, then Hani |
| Han with zh/yue and TW/HK/MO region | Hant, then Hani |
| Other/absent/invalid language, bare zh/yue, incompatible explicit script | Hani |

Explicit variant script precedes language/region; incompatible explicit script prevents region inference. Explicit Hani still allows the listed language/region policy. Locale private-use/Unicode extensions do not supply script hints. No likely-subtag maximization or host locale guessing occurs. These are PPTKit rules, not universal Office defaults. Missing candidates use existing slots/default. The map alone never installs or verifies a font.

Map mode inherits the preceding script code for punctuation, separators, combining marks and joiners. Leading neutral text uses Zyyy; unclassified characters including emoji use Zyyy. A newline is default/Zyyy and resets context. This intentionally makes CJK punctuation follow preceding text in map mode; legacy slot-only punctuation behavior remains unchanged. UTF-16 offsets, original author run identity, empty insertion and segment bounds follow the existing utility contract. New runtime Unicode assignments can affect classification.

SVG keeps one action wrapper per author run. Fresh PPTX emits concrete text fragments with the selected family in all three slots, retaining language, style, alpha and action. Adjacent equal-family fragments coalesce. It resolves theme roles to concrete families and does not persist reusable map rules. `script-font-map-materialized` warns once per element text owner, notes body or layout placeholder, even for empty/currently-unused maps that could matter after editing; an actually empty normalized map produces no warning. Placeholder defaults cannot carry the reusable map natively. Reimport may retain current concrete fonts through the explicit three-slot source profile, but cannot restore maps, theme roles or original run topology. This is a documented materialization boundary, not semantic map round-trip.

[Core map tests](../../packages/core/test/script-font-map.test.mjs), [Layout tests](../../packages/layout/test/script-fonts.test.mjs) and [browser/cross-package checks](../../scripts/script-font-map-parity.test.mjs) cover the extension separately. Full native supplemental precedence, notes themes, shaping/font coverage and PowerPoint native edit/theme acceptance remain open.

The separate [source v2 profile](supplemental-font-import.md) now transports admitted supplemental rules with actual per-key provenance and captured-profile scalar-save validation. Default/v1 keep their existing source boundaries; fresh map materialization still loses reusable rules.
