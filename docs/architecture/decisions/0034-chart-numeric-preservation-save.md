# ADR 0034: Existing chart numeric preservation saves

- Status: accepted experimental Parser contract
- Date: 2026-10-26

## Problem

Chart content import prioritizes complete caches and may recover missing caches from stored XLSX values. Editing only that snapshot or one storage layer can silently leave stale data. Existing text profiles intentionally grant no chart authority.

## Decision

Introduce independent explicit `chartEditSaveProfile: "existing-values-v1"` for content imports. Parser binds existing numeric atoms only when complete cache and local workbook ranges for names/categories/values agree exactly. Public readonly chart-value provenance exposes Core indices, raw cache identity, workbook/sheet/cell identity and bound/unmapped state. It grants conditional eligibility; whole-candidate preparation remains authoritative.

Keep Core free of source XML and nested package details. Extend the private scalar analyzer/stager rather than add a second save/session framework. Stage cache and worksheet lexical patches, independently compare inner and outer allowed payload sets, then reproject and rebind actual candidate values and identities before full Core equality. A successful cache import alone is insufficient proof of XLSX synchronization. Retain unrelated payloads and lexical surroundings, including worksheet cell formatting and unknown adjacent chart markup.

Reject shared chart/workbooks, repeated numeric roles (including name/category aliases), formulas/named calculations/external dependencies, auto-update, chart MCE projections, missing/recovered/conflicting sources and unsupported normalized edits. No formula evaluation, networking, source repair or implicit shared-edit choice. Preserve default/text-profile authority and lifecycle. Reuse all session quotas with a dedicated bounded chart-binding work budget; nested comparison consumes the same validation budget as outer comparison. Exhaustion is a save limit failure, including exactly exhausted budgets.

## Evidence and limits

The [public contract](../../api/chart-value-save.md) links focused both-dialect/family three-round tests, second-layer failure/cancellation, shared/alias rejection, exact budget exhaustion, browser independent DOM and installed entry points. Signed zero is explicit and modified numeric spellings normalize deterministically. This is not a full XLSX validator, unknown-extension semantics proof, structural chart writer or native interactive-edit certification. IMP-09.2.6.1/.2 have a bounded slice; parents and native/structural tasks remain open.

Independent follow-up: [three-round source-save/native evidence](../../testing/chart-value-save.md) adds a separate Python cache/XLSX and payload oracle. Named LibreOffice resaves retain all numeric series but drop referenced XLSX sources; import remains read-only for numeric writes. This does not expand the profile or close interactive/native gates.
