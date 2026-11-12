# Supplemental source font projection

Content import can opt into `textFontProfile: "script-families-v2"` to retain admitted theme/table supplemental font rules as Core `scriptFontFamilies`. It extends the independent named-slot projection while leaving default import and [v1](script-font-import.md) unchanged. This profile is experimental; its declared projection rules do not certify universal Office font selection.

See [ADR0049](../architecture/decisions/0049-source-supplemental-font-projection.md), the [typed example](../examples/supplemental-font-import.ts) and Core's [map/language selection contract](script-fonts.md#script-font-map-extension). `ImportReport.textFontProfile` records v2. Both bytes and Node file options are captured before asynchronous work; invalid profiles and structure mode reject before source I/O. Native fonts are neither installed nor queried.

## Source rule selection

The 37 Core key names are the admitted set. Parser retains complete supported collection rules independently of current text. Final native slot declarations win the existing fieldwise source cascade first; only then are script entries projected. This allows changing characters through admitted scalar saves without recomputing source font styles from those characters.

| Admitted script keys | Effective native slot |
| --- | --- |
| Zyyy, Latn, Grek, Cyrl | Latin / `fontFamily` |
| Hani, Hans, Hant, Hira, Kana, Jpan, Hang, Kore, Bopo, Yiii | EA / `eastAsianFontFamily` |
| All other admitted keys | CS / `complexScriptFontFamily` |

A matching theme reference such as Latin `+mn-lt`, EA `+mn-ea` or CS `+mn-cs` retains the selected major/minor collection's supplemental dictionary. If EA/CS is absent, the effective Latin declaration supplies that default bucket, including its collection rules. A present literal or unsupported EA/CS declaration replaces that bucket's previous association and prevents Latin/older-map fallback. Thus a local literal `cs` clears inherited Arabic/Hebrew rules while leaving EA rules intact. A major EA reference uses major EA rules even when Latin/CS select minor.

A deliberate cross-slot reference, such as EA `+mn-lt`, projects that named concrete theme slot; it does not opt the destination into script substitution. This preserves the explicit distinction between referencing a generic native slot and choosing script rules. These are the v2 profile's declared rules; named native comparison remains required.

Empty/missing/invalid generic theme slots remain `unsupported` named properties. Usable supplemental entries can still project individually. Their availability never turns an empty ea/cs slot into a successfully resolved Latin font. Missing schemes/groups never borrow an older component. A nearer fontScheme replaces the font component as a unit, including absent supplemental entries.

Table `fontRef` selects each native field and its collection association. A direct table `font` collection reads its own supplemental dictionary. A nonempty direct dictionary replaces referenced supplemental dictionaries for that collection; required generic slots still receive blockers when absent. Later conditional collection choices replace older choices, and local list/paragraph/run declarations remain later. A direct collection with no supplemental entries can retain rules explicitly selected by its named theme references. No opaque source theme identity enters Core.

## Findings, provenance and bounds

Each selected admitted entry records `run.scriptFontFamilies.<code>` in drawing/notes or table text provenance. Its value is a concrete string; its source is the actual supplemental typeface attribute. `resolvedFrom` includes the effective selection declaration, generic slot dependencies, collection and script identity. Missing/invalid admitted typefaces produce unsupported key records and are omitted from the Core map. Known entries require nonblank literal typefaces without controls or recursive `+...` theme references.

An entry without a script identity blocks that collection's entire admitted map. Duplicate selected script identities reject, including unknown duplicate identities. Unknown scripts, unknown attributes and children remain located source-only findings. They are not silently normalized to a supported script or installed-font fallback. MCE effective branches are selected first; effective paths are retained.

Existing cumulative `maxTextStyleProperties` / `maxTextStyleDependencies`, table property/dependency/work, XML/MCE/report and cancellation budgets still apply. Expanded map provenance is charged per effective target record; shared source metadata is bounded by existing XML limits and cached per selected theme component. Import failure returns no partial usable session; independent retry is supported. Direct notes retain the v1 scope boundary: no notes master/theme borrowing from slide themes.

## Save and fresh output

V2 does not grant edit permission. No-op retains exact input bytes. With an existing matching text-run save profile, bound scalar character changes can retain native theme/collection/map XML, including source-only rules. The actual candidate is reimported with the captured v2 profile and its complete Core snapshot compared with the approved edit. No font/map fields are masked to pass validation.

Font/map/language/style changes, including adding, changing or removing map keys mixed with characters, remain atomically blocked. Restoring baseline style permits retry. Source bindings and native target identities retain their existing scope; native font writes require separate admission.

Fresh `generatePptx(session.document)` is a separate operation. It emits the current chosen fonts through [warned map materialization](script-fonts.md#script-font-map-extension), may split/coalesce runs and loses reusable map/theme linkage. It does not rebuild the original supplemental collection. Source preservation and fresh output are tested independently.

## Evidence and remaining scope

[Eleven Parser groups](../../packages/pptx-parser/test/supplemental-font-import.test.mjs), [three browser/cross-package groups](../../scripts/supplemental-font-import-parity.test.mjs) and packed runtime/types cover both dialects, scope/cascade/choice boundaries, MCE, invalid rules, budgets/failures/capture, three scalar saves and mixed owners. Browser XPath checks actual value/script/dependency/target nodes and rejects a deliberate wrong Hant value. [Independent Python evidence](../testing/supplemental-font-import.md) checks four source packages with exact bytes outside allowed scalar contents.

Full native theme/font precedence, real-font coverage/substitution, shaping/bidi/vertical text, alt-language semantics, bullet fonts, complete notes master/theme inheritance and PowerPoint edit/theme acceptance remain open. Unsupported script keys such as Ethi stay source-only. This profile does not close the parent font or round-trip tasks.

The separate [notes text profile](notes-style-import.md) can now opt into independent notes-master/theme inheritance alongside v2. Without it, v2 notes remain direct-only as described above. Complete notes-page/native appearance still has separate gates.

[Named native observations](../testing/real-fonts.md) support matching EA/CS supplemental references for the tested PowerPoint cases, while retaining five coverage/selection differences and LibreOffice-specific differences. In particular, v2's supplemental fallback through a Latin collection when EA is absent is an experimental projection rule, not a guarantee of PowerPoint-selected faces. Bare `zh` remains region-ambiguous in Core. A resolved family is independent of installed glyph coverage; the source profile does not query host fonts or write native substitutions back into source rules.

The later 32 paired controls confirm that the absent-EA/CS Latin-collection heuristic differs from native selection even with covering installed fonts. Short `ko` also differs from `ko-KR` in this PowerPoint environment; the actual Korean face covers both tested Han and Hangul. The separate [declared-slot v3 profile](declared-slot-font-import.md) now excludes that missing-bucket heuristic and reports unresolved source rules; it does not rewrite imported fonts/languages or infer implicit Office fonts. Existing v2 remains explicitly experimental and compatible. Full native table/notes/editing gates remain open.
