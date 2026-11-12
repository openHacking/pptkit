# ADR 0002 — Stored chart snapshots and embedded workbook data

Status: accepted for the experimental bounded content projection. Related tasks: IMP-06.2.1/06.2.2, with partial implementation evidence for IMP-06.2.3–06.2.5. This decision does not close their inherited gates or establish full AUT-06.2.

> Update: [ADR 0063](0063-chart-explicit-blanks.md) admits explicit blank policies while retaining the original cache-first and source-preservation boundaries. [Doughnut projection](../../api/chart-doughnut.md) extends the supported chart families with separately bounded geometry.

## Problem

A PPTX chart can contain complete, sparse, missing or stale cached points alongside a workbook or external data link. Treating these as interchangeable silently changes the displayed data or implies formula execution. Core's current bar/line/pie model stores nonempty category strings and dense finite arrays; sparse/error/date/format/formula semantics cannot be represented honestly by zero-filled numeric arrays.

## Decision

Complete chart cache/literal values own the stored snapshot. Compare supported local workbook stored ranges when available and report conflicts explicitly. Only absent caches may use complete compatible workbook values. Existing sparse/empty caches remain source-only. Unsupported formulas and external references are never evaluated/fetched. Formula cells use stored results with explicit findings. Malformed required relationships and supported nested packages fail the import, even when a cache exists; unsupported valid data profiles remain reported source state.

Source chart/workbook/formula/format identity remains Parser-owned. Parser reuses retained effective chart trees and bounded nested OPC readers. Core contains only public semantic arrays/styles. No chart-specific package, public parser dependency in Core/Layout/Exporter or new data-binding abstraction is introduced.

The importer caches distinct workbook inspection by canonical part and shares expanded-byte and data-work budgets across the invocation. Referenced cell reads count in the work budget, including repeat ranges; declared counts/ranges cannot drive unbounded work. Traversal is cancellable. Public structural/content reports are detached snapshots so later chart MCE interpretation cannot mutate or be blocked by an earlier frozen report.

Each source drawing produces an independent Core chart value. Fresh export rebuilds one chart/workbook pair per element from current arrays; shared source parts do not imply shared mutable Core data or preservation sharing. Source no-op retains exact package bytes. Source chart edits and rebuild stay blocked until chart reconciliation exists. The subsequent unrelated [text-run profile](0008-direct-text-preservation-save.md) retains chart/workbook payloads without editing them.

## Alternatives and consequences

Workbook-first import can replace a stored displayed snapshot with stale/recalculated values. Cache-only import cannot recover a legitimate missing-cache chart. Filling holes or inventing names/categories conceals unsupported data. Keeping two mutable datasets in Core would introduce a format-specific consistency problem without an authoring contract. The selected policy preserves explicit source disagreements and provides a bounded recovery path while exposing limitations.

Chart layout and unrepresented style/format properties remain approximations, reported per part/path/object. This policy does not promise native formula links, visual equivalence or arbitrary round-trip. Interactive data editing remains a separate application gate.

## Evidence

- [Independent cache/workbook fixtures](../../../packages/pptx-parser/test/charts.test.mjs): complete, conflicting, absent, sparse, stored-formula, external, invalid and budget cases.
- [Independent fresh-output cache/worksheet checks](../../../scripts/chart-import-parity.test.mjs): every mapped value and stale-data removal across three rounds.
- [Public contract](../../api/chart-import.md), [typed example](../../examples/chart-import.ts), [native evidence](../../testing/chart-import-evidence.json).
