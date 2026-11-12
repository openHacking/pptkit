# Experimental slide-bound foreground timing

Slides can now retain authored finite gate stages and normalize them into an object-bound [event timeline plan](event-timeline.md). This advances ADV-06.1.2.2 for whole objects across Core, Canonical IR v2, JSON and Layout. Native effects, native timing import/serialization and actual playback remain unimplemented.

## Authoring and current ownership

`PresentationSlideInput.eventTimeline?: SlideEventTimelineInput` and `slide.setEventTimeline(value?)` accept `{ stages: readonly EventTimelineStageInput[] }`. The setter validates every timing tree, target and trigger against the actual owning slide ID and its owned object forest, then retains a detached deeply frozen source snapshot. It does not retain the caller's declared inventory, visit, event log, progress, clock or compiled cache. Omitting/clearing the field restores the earlier untimed behavior.

`slide.eventTimeline` exposes that immutable authored tree. `slide.getEventTimelinePlan()` compiles a fresh frozen plan from the current real IDs, or returns undefined when absent. Whole group objects and their descendants can be targeted. Shared layout objects, another slide's objects and paragraph/chart-series/SmartArt subtargets cannot be bound. Hidden state does not remove identity; eventual effect visibility policy is not inferred here.

The integrated profile uses the event utility's default budgets: 10,000 owned objects, 1,000 stages, 10,000 aggregate timing nodes, depth 128, a 24-hour finite-time ceiling, repeat/ID bounds. Inventories are collected within their object budget and sorted by exact ID before compilation, so visual reordering does not reorder timing identity. Long IDs reject before sorting. Attach a program after building its object forest; initial addSlide input attaches it only after its elements are materialized. Standalone numerical utilities still support their explicit larger limits; the slide profile has no stored quota override.

Core validation also checks the field on structural document consumers and reports located errors. A failed setter or initial slide attachment commits no timeline/slide and releases candidate element reservations. Generated ID counters may advance after failed candidate work, as with existing construction; failures do not reserve IDs or alter existing objects.

## Identity-preserving updates and structural operations

`slide.replaceElement(elementId, input)` atomically replaces a top-level object while retaining its root ID. Input id can be omitted or equal to the original. A different root ID, collisions with other objects, duplicate candidate child IDs or removal of a timed child/source reject before registry/collection commit. Retained child IDs must be explicit in a replacement group. Input and old snapshots remain detached; new IDs are registered only after timeline validation succeeds. Released snapshots cannot replace a live object or unregister an ID later reused elsewhere. Other document validation remains the existing Core validation contract; this method does not certify every unsupported connector/format reference.

| Operation | Timing behavior |
| --- | --- |
| Object reorder | Exact target/source IDs and stage order remain unchanged |
| Whole-slide copy | Allocator supplies a complete root/descendant ID map; target page/object and object-click source remap together |
| Object/group-only copy | Copies visual objects and their existing supported references; retains the original slide-scoped program without cloning gate stages |
| Remove timed object/source or a group containing one | Atomic rejection; explicitly replace/clear the program first |
| Remove unreferenced object | Allowed; later inventories reflect the remaining forest |
| Replace timed object | Root identity preserved; missing timed descendants/sources block the whole replacement |
| Remove whole slide | Its program and local object forest leave the active document together |

Stage IDs are local to their page program and remain stable on whole-slide copy; finite node IDs remain local to the stage. No stage/target ID is matched by label, array position or geometry. The existing document allocator owns new object IDs. Cross-document copy, general page/format reference reconciliation and external retained-plan revocation are separate gates. External compiled plans remain immutable snapshots; applications must still provide fresh visit/current inventory and avoid nominal ID recycling.

## Canonical IR and Layout

`NormalizedSlide.eventTimeline?: NormalizedSlideEventTimeline` is an additive optional field in experimental Canonical IR v2. The wrapper has `schemaVersion: "pptkit-slide-event-timeline-v1"`, the complete authored `source` and compiled `plan`. Group IDs, topology and empty groups remain in source instead of being discarded as intervals. Normalization derives both from owned authoring data and returns detached snapshots; absent fields do not become empty plans. Existing untimed IR/output shapes remain unchanged. The plan carries its own schema version and explicit typed targets/triggers/finite intervals, without raw format IDs/XML.

JSON transport preserves it. `LayoutSlide.eventTimeline` retains a detached copy without sampling or changing geometry. Direct normalized-layout input validates the wrapper/schema, complete stored plan and actual targets, recompiles the retained source against actual owned IDs and compares the entire plan before transforms. Coherent flat intervals alone cannot hide source/plan drift. Call `replayEventTimeline(layoutSlide.eventTimeline.plan, request)` to evaluate the stored plan; the wrapper itself is not a standalone event-plan input.

Single-page table pagination can retain the program. A requested split of a timed slide yields `pagination-event-timeline-unsupported` and applies no table/page changes: continuation target ownership has no admitted reconciliation policy. Geometry/measurement/composition retain identities and cannot invent new timing tracks. Callers may explicitly clear the program, or choose explicit exporter omission before requesting a split.

## Format boundaries and explicit loss

PPTX has no admitted native mapping for these authored stage trees. Default `generatePptx`/Node `writePptx` therefore reject with `PptxExportError`, code `event-timeline-export-unsupported`, before asset resolution, package output or file writes. Frozen `diagnostics` locate each stage by slideId, elementId, stageId and document path; an empty authored program gets one slide-level diagnostic. This prevents silent loss and does not grant native timing export.

Explicit `eventTimelineMode: "omit"` removes only the private normalized output candidate's event fields, reports `event-timeline-omitted` for every stage (or empty program), and generates static output. Omission precedes Layout, so requested table splits can proceed under that explicit loss choice. Authoring state remains unchanged. Invalid mode values reject. Other source-preservation/loss consent is not bypassed by this fresh-output option.

Static SVG emits unchanged authored geometry with `event-timeline-static` per timed slide. It neither applies progress nor embeds a player. A direct resolved Layout caller remains responsible for that Layout contract; document-facing rendering validates through Core/Layout.

Parser's native source timing remains format-owned/source-only. No native import becomes this Core field. Attaching or modifying authored timing on an imported session is unsupported for edit-save; a mixed timing/text candidate blocks completely. Explicit clear restores the supported text save path. Existing no-op and unrelated source preservation retain their earlier rules.

## Checked example

<!-- doc-test: docs/examples/slide-event-timeline.ts -->
```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function boundSlideTimingExample() {
  const document = createPresentation();
  const slide = document.addSlide({ id: "page" });
  slide.addElement({ id: "text", type: "text", box: { x: 30, y: 30, width: 300, height: 50 }, content: "Before" });
  slide.addElement({ id: "button", type: "shape", shape: "rect", box: { x: 30, y: 100, width: 100, height: 40 } });
  slide.setEventTimeline({ stages: [{
    id: "stage", target: { pageId: slide.id, elementId: "text" },
    trigger: { type: "object-click", elementId: "button" },
    timing: { kind: "span", id: "time", durationMs: 200 },
  }] });
  // The logical text ID stays the same; no temporary dangling target exists.
  slide.replaceElement("text", { type: "text", box: { x: 30, y: 30, width: 300, height: 50 }, content: "After" });
  const copy = document.duplicateSlide(slide.id);
  const normalized = normalizePresentation(document);
  const layout = resolveLayout(document);
  // Default PPTX generation rejects this unimplemented native mapping.
  // Explicit omission authorizes a static deck with located loss warnings.
  const staticPptx = await generatePptx(document, { eventTimelineMode: "omit" });
  return { document, copy, normalized, layout, staticPptx };
}
```

## Evidence and remaining gates

- [Nine Core groups](../../packages/core/test/slide-event-timeline.test.mjs): owned roots/descendants, immutable input and detached IR, setter/init/registry rollback, delete guards, atomic root/group replacement and conflicts, three whole-slide copies after source removal, object-only copy/reorder policy, layout-only target rejection, structural-consumer diagnostics and released-snapshot ownership checks.
- [Seven cross-package groups](../../scripts/slide-event-timeline.test.mjs): Core/IR/JSON/Layout/replay identity, static SVG bytes/warning, corrupt/stale input, atomic split refusal and single-page retention, default PPTX/file/asset-work refusal, exact static package payloads on explicit omission, omission before pagination and Parser whole-plan blocking/retry.
- [Real browser ESM/global](../../scripts/browser-bundles.test.mjs) and installed [runtime](../../tests/consumers/node.mjs)/[types](../../tests/consumers/types.ts): public methods, copying, retained plans, readonly trees, explicit output mode and refusal diagnostics.

[ADR0037](../architecture/decisions/0037-slide-event-target-transport.md) records ownership and compatibility. Native effects and timing adapters still require ADV-06.1.2.3/ADV-06.2's mapping and actual-playback gates. Subtargets, independent tracks, clone-stage policies for object-only copy, multi-page timing reconciliation, complete source structural saves and general cross-document transport remain open. No animation/SmartArt/native editable output claim is added by this field.
