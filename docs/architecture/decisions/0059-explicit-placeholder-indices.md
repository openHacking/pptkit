# ADR0059 — Explicit layout-local placeholder indices

- Date: 2026-10-27
- Status: implemented and verified for authored numeric matching/allocation; native template import remains separate

## Problem and contract

Logical keys and unique-role fallback cannot distinguish renamed body placeholders when several target bodies exist. The exporter previously assigned native numbers only from traversal order. Promoting those generated numbers into authoring identity would make an unrelated reorder change matching decisions.

`PlaceholderDefinitionInput.index?: number` is an optional explicit layout-local numeric identity. It is a nonnegative unsigned 32-bit integer, including zero, unique across all roles within its layout. This is an additive value field, not an OOXML node, relationship ID or package reference. Its numeric range interoperates with the native [`idx` unsigned-32-bit property](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.placeholdershape.index?view=openxml-3.0.1). Authored duplicate/invalid values receive located validation errors; atomic definition replacement retains old definitions and IDs on failure. Legacy definition creation retains its existing deferred validation behavior.

Normalization copies the field only when authored. Missing indices stay missing in IR v2 and authoring snapshots. Layout carries the same optional value; matching never obtains an index from a collection position or a previously generated PPTX ordinal. This extends the bounded matching decision in [ADR0058](0058-reviewed-layout-rebinding.md).

## Matching and export

Priority is explicit caller mapping, same logical key with the same role, declared index with the same role, then one unique same-role candidate. A same-key or same-index role conflict blocks instead of falling through. Missing index matches may use the existing unique-role fallback; ambiguity still needs review. Matches report declared source/target indices when present; choosing a different declared target index emits an informational finding. Ordinary objects gain no inferred binding.

Exporter owns a separate native-number allocation table. It reserves all explicit indices, prefers the existing one-based position for implicit entries, and uses the first free positive number when that position is reserved. Layout prompts and all local text/picture/shape/chart/table/connector placeholder declarations use the same table. The free-number cursor is monotonic; dense explicit reservations do not repeatedly scan from one. Allocation does not write generated values back to Core. With no explicit indices the old payloads and IR remain identical.

## Limits and evidence

This does not import native templates into Core or authorize source writes. Parser keeps its existing native index/type provenance and no-op authority; authoring definitions and source declaration identity remain distinct. Cross-document identity remapping and per-property reset remain separate original requirements.

Focused Core tests cover zero/max/invalid/duplicates, immutable snapshots, priority, order changes, conflicts, review/explicit detach, fallback, rollback and staleness. Export tests check mixed collision-free allocation and identical prompt/local bindings including pictures. Public consumer and native evidence status is recorded in [placeholder index verification](../../testing/placeholder-index.md).
