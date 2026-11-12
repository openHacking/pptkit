# Declared-slot source font evidence

The [runner](../../scripts/inspect-declared-slot-font-import.mjs) builds a named [mixed-owner source fixture](../../tests/fixtures/declared-slot-font-import.mjs) with six scalar owners: two text boxes, shape text, two table cells and independent notes. Each local run references the native Latin theme slot; EA/CS are absent from all local and inherited style scopes. Slide and notes themes have distinct supplemental values. This is native source XML construction, separate from author-map fresh export.

The explicit `script-families-v3` profile retains **six mapped Latin-bucket Grek rules** and **36 unsupported EA/CS rules**, without mapped values on the latter. Original plus three admitted scalar text saves preserve those 42 source observations. The [recorded manifest](declared-slot-font-import-evidence.json) lists all PPTX hashes and local artifacts under `.tmp/declared-slot-font-import/run-LawUAt`.

The independent [Python stdlib oracle](../../scripts/research/oracles/declared-slot-font-import.py) reads actual ZIP/XML only. It verifies each source script/typeface against named expected values, actual selector/dependency/target paths, actual slide/layout/master/theme and notes/master/theme relationships, and absence of EA/CS declarations throughout the named local/inherited source scopes. It checks all six actual scalar texts and exact part-name sets, unchanged payloads and unchanged non-character bytes in the two allowed text parts.

At every stage, four actual ZIP controls mutate a mapped source value, an unsupported source value, the claimed absence by inserting EA, or the actual theme relationship. Both semantic and exact-preservation checks must reject every control. Unsupported report values are not compared to `undefined` as a substitute for validating their actual source font values. The fixture-specific scalar mask is not a general XML writer.

```sh
node scripts/inspect-declared-slot-font-import.mjs
```

[Browser parity tests](../../scripts/declared-slot-font-import-parity.test.mjs) separately verify both dialects, Node/ESM/global snapshots/report/bytes and independent XPath source values for mapped and blocked rules, with network blocked and a changed blocked source negative control. Mixed-owner source preservation and warned fresh materialization are distinct checks. [Parser tests](../../packages/pptx-parser/test/declared-slot-font-import.test.mjs) add inheritance/empty/literal/cross-slot/major choices, notes/table/empty insertion contexts, scope/MCE/unknown rules, actual candidate saves, captured file options and resource/permission failure behavior. Installed public runtime/types consumers also import v3 and prepare/save a scalar edit.

This evidence proves the named source contract. It does not certify native implicit fonts, glyph coverage/shaping, table/notes appearance, source style writes or PowerPoint edits. The earlier native language/fallback differences remain recorded; v3 does not silently normalize them away.
