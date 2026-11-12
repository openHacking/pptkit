# ADR 0006 — Preserve source relationship dependencies across MCE branches

Status: accepted for the implemented inspection profile; preservation editing remains planned.

Date: 2026-10-25.

## Problem

An effective MCE projection can omit inactive source relationships. Using that graph as a complete preservation dependency graph risks dropping resources retained by another branch or by an outside owner. Treating root-unreachable parts as garbage would also discard source content without proof.

## Decision

Parser owns `inspectOpcClosure` with explicit roots. It reuses the bounded OPC reader and retains only raw relationship documents for a conservative source union, without adding Core/IR fields, an Exporter dependency or another package.

Ordinary XML is still validated but its tree is not retained for this analysis. The internal XML retention mode is separate from existing full-tree presentation readers. All source `Relationship` declarations are accounted for, including inactive choices/fallbacks; unknown markup and invalid/missing inactive dependencies produce located findings rather than disappearing.

Graph traversal includes relationship-set ownership, follows recognized internal targets, records outside incoming owners, and retains unreached identities in the report. Cyclic/shared targets are visited once. External targets are metadata only.

Source relationship IDs remain located by raw element index as well as owner/ID. Mutually exclusive branches can legitimately repeat IDs. Potentially coexisting declarations have an ambiguity finding; separate MCE expression correlations are conservatively unresolved rather than introducing another complete MCE solver.

Completeness flags are scoped to source relationship markup and directions. An outside unknown source prevents complete incoming ownership; a local unknown source prevents complete outgoing coverage. Neither flag grants garbage collection, extraction, source subtree mutation or edit-save permission.

Root/source-declaration/selected-part/selected-byte/finding/work limits and cancellation bound the call. Limit exhaustion returns no partial graph or truncated completeness result. Public reports are immutable metadata; private trees, byte retention and future serialization ownership stay in Parser.

## Consequences

The analysis supplies a reusable dependency foundation for IMP-07/09, SmartArt component/resource preservation, shared media ownership and inspection tooling. It deliberately leaves arbitrary implicit XML/binary references, source subtree edits, identity rewrites, session reconciliation and candidate writes to their owning follow-up tasks.

Whole-source no-op retains unreachable and opaque source content as before. No new cleanup or editing API is admitted by this decision.

See [the closure contract and evidence](../../api/preservation-closure.md) and [import/save ownership](0001-import-preservation-save.md).
