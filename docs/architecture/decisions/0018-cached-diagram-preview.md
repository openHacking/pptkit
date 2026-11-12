# ADR 0018: Explicit cached diagram preview projection

Status: accepted for an experimental bounded content option, 2026-10-25. Partial ADV-01.1.2; full native/cache-freshness and inherited gates remain open.

The subsequent [ADR0019](0019-cached-diagram-text-transform.md) extends the original subset with independently positioned text rectangles and relative angles, retaining explicit unavailable cases and source-save boundaries.

The source inventory can associate cache parts without interpreting a data graph. Consumers need to display that existing drawing or an explicit unavailable placeholder. Adding XML to Core, mapping cache shapes to data nodes, or making an unknown-age cache the default would blur source authority and editing semantics.

Parser adds content-only `diagramPreview: "cache"`. A separate private cache drawing reader interprets DSP shape/group wrappers, reusing the existing DrawingML geometry/paint/text and image-loading helpers. It constructs ordinary Core groups and reports source/cache provenance, placeholder/partial results and unknown freshness separately. Cache root/nested child spaces are represented explicitly; unsupported geometry or independent text transforms do not gain guessed semantics. Empty cached trees are unavailable, even when relationship/root inventory is located. The renderer consumes ordinary public Core/Layout values and receives no XML or Parser dependency.

This is a cached visual approximation, not a SmartArt model. No source bindings are created for cached text/shapes, and native create/edit/export/edit-save remain unsupported. Ordinary text-save candidate reprojection retains the captured preview option so snapshots compare the same projection. The Node file adapter captures that option before IO. Fresh export of derived groups is explicitly ordinary-shape output, never native SmartArt preservation.

Repeated use of shared source XML is charged to a dedicated aggregate preview work budget, in addition to existing finding/snapshot/image budgets. External cache images are not loaded. Malformed mapped scalars fail with source provenance; budget/cancellation errors propagate. Missing, empty, unresolved/MCE and unsupported caches receive explicit placeholders or property findings. The implementation does not claim freshness detection or full theme/layout/geometry fidelity.

Validation covers both dialects, defaults and invocation-time capture, root/nested geometry, image sharing/budgets, unsupported/hidden objects, missing/empty/MCE/wrong-root caches, cache-part failures, blocked cache edits, unrelated text saves and Node file propagation. Browser public bundles render cached shapes/text/images with independently specified screen bounds and no network. Three-round source saves retain cache/image payloads; fresh export is checked to remain ordinary shapes. Complete/native-family fidelity and stale-cache policy remain later ADV-01–03 work. See [the public contract](../../api/diagram-preview.md).
