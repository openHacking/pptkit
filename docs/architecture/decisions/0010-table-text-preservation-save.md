# ADR 0010 — Physical table anchors in text preservation saves

Status: accepted for experimental `text-runs-v3`.

Date: 2026-10-25.

## Decision

Extend the [V1](0008-direct-text-preservation-save.md)/[V2](0009-text-owner-preservation-save.md) pipeline with an explicit V3 table-character profile. Keep earlier profile scopes/report shapes compatible. V3 permits only existing bound anchor run characters, not style/metric/merge/topology writes.

The existing table reader validates the full physical grid before publishing any cell runs. Each anchor callback carries its source cell and zero-based source row/physical column plus compact Core cell index. The private binder uses MCE effective-to-raw origins; it never guesses physical column from compact indices or searches text. Missing bodies and synthetic paragraphs have explicit unmapped records. Unsupported tables publish no bindings.

Public `tableTextRuns` is a separate immutable collection with the table element ID and row/cell/column coordinates. It does not invent cell element IDs or overload drawing/notes provenance. Private scalar identity includes row/cell indices. Source notes/drawing/table records share the same bounded run collection and source guards.

The pure snapshot analyzer locates real normalized cell content paths, rejects every unsupported mutation in the whole document and reports scalar table grid metrics as geometry and spans as structure. Raw table sub-target/range/opaque references block edited owners. Existing lexical changes, ZIP patching, shared-source reconciliation, exact candidate comparison, actual-output reprojection, budgets and lifecycle remain the only acceptance pipeline.

No new package, dependency, Core/IR field, table reconstruction or Exporter coupling is introduced. Source table-style references, continuation cells, unknown extensions and other resources remain retained.

## Evidence and consequences

The [public table-save contract](../../api/table-text-save.md) links independent merge partitions, dialect/encoding/mixed-mutation tests, three-round mixed decks, browser DOM/byte parity, installed consumers and headless native evidence. The native producer retains edited table characters and tables but changes geometry/representation and loses opaque content retained by PPTKit; these are recorded differences.

This advances the first text-only slice of IMP-09.2.5. Direct styles/metrics, structural transactions, range remapping, implicit-reference completeness, PowerPoint/interactive editing, rebasing, atomic file output and general rebuild remain separate open tasks. No inherited/native parent gate is closed.
