# ADR0051: Supplemental projection through declared native slots

Status: accepted for experimental `script-families-v3`, 2026-10-26. Full native typography, implicit font selection and source style writes remain open.

## Problem and decision

[ADR0049](0049-source-supplemental-font-projection.md) retains source supplemental rules before selecting a font for current text. However, its v2 finalization also takes a Latin collection's script dictionary when the effective EA/CS declaration is absent. [Paired PowerPoint controls](../../testing/real-fonts.md#paired-native-precedence-controls) contradict native selection parity for that heuristic even with covering fonts. Silent mutation of v2 would change existing canonical snapshots and source-save contracts.

Add explicit `textFontProfile: "script-families-v3"`. V1, v2 and the default remain compatible. V3 uses the same bounded source readers, inheritance, collection resolution and source script dictionary; only final supplemental selection changes. Each script consults its effective native bucket after the complete field cascade. An inherited declaration is a declaration. A selected explicit table font collection owns its bucket choices. A matching native theme reference may carry its collection map even when the collection's generic family is unsupported. A concrete literal or cross-slot reference does not gain that destination's supplemental substitutions.

When an EA/CS bucket is absent, v3 does not map rules from the Latin declaration into it. Those entries remain located `unsupported` script properties with no `value`, an explicit reason, the actual source supplemental node, Latin selection dependencies and actual target. A located unmapped finding explains the unresolved selection. No missing EA/CS XML node is fabricated, and no current character/language/installed-font query decides this admission. Latin-bucket rules still use the effective Latin declaration.

Finalization uses the existing common text reader for drawings, tables, notes and empty insertion styles. Drawing/notes and table owners supply their existing bounded finding sinks explicitly. New unsupported records consume the same property/dependency/finding quotas; notes theme traversal retains the separate work budget. Private collection metadata is removed before report serialization. No Core model, new package, dependency or global font inventory is added.

## Preservation and compatibility

The byte/Node option capture and actual candidate reprojection retain v3. Source scalar edits can change characters/scripts while keeping the same admitted map and unresolved source observations. Complete canonical equality remains required; fields are not masked. Existing map/font/language/style edits remain blocked atomically, and v3 grants no save permission by itself.

Source no-op/scalar saves preserve the original theme and slot declarations. Fresh export is separate: admitted maps materialize with the existing rule/topology-loss warning. Unsupported source selection is not reconstructed by fresh output. Core's normal generic fallback may still render a preview; it is not a claim that Office selects that family.

V3 does not infer Office's implicit fonts for absent slots, normalize `ko` into `ko-KR`, choose a Han region for bare `zh`, emulate host substitutions, provide shaping/bidi, or certify native table/notes/theme editing. The public choice is conservative source-rule admission, not a universal native font resolver. Explicit table collections may retain their own supplemental entries while omitted required generic fields remain unsupported; this differs from borrowing an unrelated Latin declaration for an absent run bucket.

## Evidence

[Parser tests](../../../packages/pptx-parser/test/declared-slot-font-import.test.mjs) cover both dialects, unchanged old profiles, absent/empty/inherited/literal/cross-slot and major choices, table collection/fontRef/local scopes, independent notes themes, empty insertion styles, multiple masters, effective MCE, unknown/duplicate script boundaries, captured byte/file options, three scalar saves, atomic style guards, budgets/cancel/retry and no writer permission.

[Cross-package tests](../../../scripts/declared-slot-font-import-parity.test.mjs) compare Node/ESM/global full snapshots, provenance and bytes across three saves with network blocked. Independent XPath checks actual mapped **and blocked** source values, selector dependencies and target nodes, and rejects a changed blocked source value. Mixed text/shape/table/notes source saves retain all 42 rule records and every non-character source byte; fresh materialization is separately warned. Packed runtime/types consumers exercise the installed public v3 option and an actual scalar save.

[Independent source evidence](../../testing/declared-slot-font-import.md) checks six scalar owners across four actual packages, each with six mapped Latin-bucket rules and 36 unresolved EA/CS rules. Actual theme chains and absence of local/inherited native buckets are verified from source XML, not inferred from the Parser report. Four actual ZIP mutations per stage must fail both semantic and exact-preservation checks. Native editing remains a separate gate.
