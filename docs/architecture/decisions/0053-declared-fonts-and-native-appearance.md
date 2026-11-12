# ADR0053: Declared fonts and native appearance are separate contracts

Status: accepted for experimental source projection and named interoperability evidence, 2026-10-26. No runtime behavior or source-write permission changes.

## Decision

Keep `script-families-v3` as the conservative declared-source contract in [ADR0051](0051-declared-slot-supplemental-projection.md). Do not add an Office-font resolver or a replacement profile based on the current native samples. Preserve default/v1/v2/v3 snapshots and captured candidate options. Native fallback family names and inferred language must not replace source declarations in Core, provenance or saved XML.

Report three separate observations in interoperability evidence:

1. **Declared source:** actual source owner, selector, generic slots, script dictionary, language and local overrides, resolved through the real relationship chain.
2. **Projected selection:** the admitted source rules and Core selection under the named profile, including located blocked rules. This is not an application font-selection oracle.
3. **Native appearance:** actual face/glyph/visible ink and layout under a named application version, operation and environment. Record disagreement with the projection as a difference, even if native output remains stable across saves.

The existing Parser findings describe supported source projection. They must not be presented as warnings that enumerate every possible native substitution. No new automatic warning is added in this decision: a universal warning would imply knowledge of host fonts and application selection that the current pure Parser does not possess. Consumers requiring appearance parity must consult the named evidence and retain the original application as a separate validation gate. If a future resolver is justified, admit it through an explicit contract and independent cross-environment controls; do not silently reinterpret an existing profile.

## Evidence and limits

The [36 table controls](../../testing/native-table-fonts.md) on PowerPoint for Mac 16.104 (16.104.25121423) have 20 matches and 16 differences after three unedited native save/reopen rounds. Minor/major short `ko` differ from `ko-KR`. Explicit collections with empty generic EA/CS select SimSun/Batang/Arial for the seven non-Latin samples; filled generic EA/CS select ArialUnicodeMS. Supplemental declarations remain stored. Latin/Greek agreement in these collections does not prove supplemental-rule use.

These controls isolate source references, generic slots and language tags; they do not prove universal Office precedence, selection after editing, complete bidi/mark correctness, notes behavior or Windows compatibility. The same PostScript name in a local font and a PDF does not prove subset byte identity. Stable raster output does not prove unchanged source semantics.

Generic Core fallback, warning-bearing fresh map materialization and exact source-preserving saves retain their existing, separate contracts. Source scalar-save guards must continue to reject font/map/language/style edits atomically; this decision does not authorize a source style writer.

## Next admission evidence

Use new temporary copies for table text/style/theme editing. Capture new runs and local overrides rather than requiring edited cells to remain identical to the unedited baseline. Keep the strict baseline oracle unchanged and use a dedicated editing oracle with explicitly allowed changes, unchanged neighboring controls and actual ZIP/PDF mutation failures. Each claimed operation needs three native save/close/reopen rounds and candidate reprojection/no-op evidence.

Independent notes master/theme baselines precede notes editing. A slide-theme operation must not be assumed to update an independently themed notes master. The [execution cards](../../planning/advanced-capability-todos.md#表格与备注原生字体的可领取任务) own remaining acceptance and status. This ADR completes only the policy decision; complete native typography and the parent import/round-trip tasks remain open.

## Independent notes baseline evidence

The [30 notes controls](../../testing/native-notes-fonts.md) add two independent notes themes, mismatched local/master indices with matching body types, local concrete slots and near empty/bad EA declarations. Three actual PowerPoint Notes-print rounds preserve full PDF traces and rasters. Twenty-four declared selections match; six short-`ko`/empty/bad-EA differences remain. B's empty EA selects SimSun while its bad EA selects MicrosoftYaHei; neither case justifies filling unresolved source slots with an application fallback. The notes theme is relocated and redundant source declarations are removed by the application, so unchanged appearance is not source byte preservation. This baseline supports the existing separation decision; it adds no resolver, source writer, Windows, full notes-page or bidi guarantee. Notes text/style/theme editing remains a separate gate.
