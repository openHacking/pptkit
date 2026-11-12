# Experimental direct text-run preservation saves

`importPptx(source, { projection: "content", editSaveProfile: "text-runs-v1" })` enables a fixed-topology character-edit profile in Parser root, browser ESM and global APIs. Default imports retain their existing blocked edit behavior. The profile is not accepted with structure projection.

The separate [V2 owner profile](text-owner-save.md) additionally permits mapped shape text and direct notes bodies without changing this V1 contract.

The profile modifies existing imported **text boxes** (`type: "text"`), including text boxes inside mapped groups. Source shape text, notes, table cells, chart data, fields, breaks and synthesized insertion runs are not editable through this profile. It does not add an arbitrary OOXML editing API or use fresh export to reconstruct a source slide.

## Allowed operation and provenance

Only scalar text on a bound existing run may change. Paragraph/run counts, other run properties/actions, object identity/order, styles, frame properties, geometry, groups, assets, metadata, sections and slide order must remain equal to the imported normalized baseline. Setting an equivalent value or restoring the exact baseline makes original-byte no-op eligible again. Saving never advances the session baseline; reimport returned bytes to start a new baseline.

Opt-in reports add frozen `textRuns: readonly ImportedTextRunSource[]` metadata:

- Core `slideId`/`elementId`, zero-based `paragraphIndex`/`runIndex`.
- Source part, original drawing path/object identity and one-based raw `sourceElementIndex` (also `source.elementIndex`). Raw preorder includes inactive MCE branches and wrappers.
- `status: "bound" | "unmapped"` and an optional reason. Bound means a matching character-only source atom was established; it is not unconditional save permission.

MCE projection privately maps copied effective elements back to their exact raw source nodes. No source match is inferred by searching for text, prefix spellings or object IDs alone. Duplicate text, namespace aliases and selected/inactive branches remain distinct. A source atom containing child elements, comments or processing instructions is unmapped for editing; CDATA, ordinary character content and empty atoms are supported. Unmapped source content remains retained.

Normalization's derived `plainText` change is admitted only on actual mapped text owners and is not duplicated in the change list. A similarly named custom-data field receives no exemption. Multiple Core views of one raw atom must agree on the proposed text; conflicting proposals block with `save-source-shared-conflict`. Current presentation inspection separately rejects duplicate slide-part targets.

## Preparation, report and source guards

`prepareSave({ mode: "edit" })` analyzes the **whole** normalized candidate before writing. A supported text change combined with any unsupported change blocks the entire plan. Valid-state change analysis returns a complete `changes: readonly PptxSaveChange[]`; limit exhaustion throws instead of silently truncating it.

Change paths are escaped JSON Pointers into normalized IR. Kinds are `text`, `style`, `geometry`, `asset`, `structure`, `metadata`, and `other`. A scalar text change includes detached `text.before`/`text.after`; source/Core provenance is included when established. Unmapped authoring changes do not fabricate a source location.

Additional gates:

| Condition | Result |
| --- | --- |
| No normalized changes | Blocked `save-edit-no-changes`; use no-op to retain original ZIP bytes. |
| Unsupported mutation/unmapped atom | Blocked `save-edit-unsupported`. |
| Text/sub-object or opaque shape target referring to an edited object | Blocked `save-text-range-reference`, including inactive raw branches. |
| Malformed/unresolvable raw shape target | Conservatively blocks text edits in that source slide. |
| Whole-shape target with no sub-target/opaque attributes | Retains target identity and source XML; ordinary text replacement is allowed. |
| Sub-target owned by another object | Retains its source bytes; it does not block the unrelated text owner. |
| Recognized signature/VBA semantics | Existing import policy rejects active/recognized parts; additional raw/inactive source guards block `save-text-package-policy`. |
| Opaque, missing or ambiguous raw source relationships | Blocked `save-source-closure-incomplete`; the raw closure inspector supplies located findings. |
| Invalid XML replacement characters | Blocked `save-text-candidate-unsupported`. |

Text range remapping, including Unicode/paragraph/SmartArt/chart sub-target semantics, remains separate work. This profile guarantees the documented character transformation and surrounding byte preservation, not interpretation of arbitrary application references inside opaque XML/binary payloads.

Ready edit plans use `assurance: "scoped-preservation"`, an immutable change report, no losses and no diagnostics. Omit consent or pass `acceptLossIds: []`; foreign/nonempty consent does not authorize unsupported edits. General lossy rebuild remains blocked. Existing ownership, disposal, stale-plan, invalid-document and exact-consent rules from [import sessions](import-session.md) apply.

## Candidate construction and validation

1. Validate source relationship coverage across all raw branches. Read source payloads under captured ZIP/XML/OPC budgets.
2. Replace only the identified character regions using the private [lexical XML primitive](../architecture/decisions/0007-xml-character-edit-boundaries.md). Retain surrounding namespaces, order, comments, extension payloads, inactive branches, XML declaration and UTF-8/BOM-marked UTF-16 encoding. Changed CDATA/entity spelling may become escaped text within the approved atom.
3. Preserve untouched compressed local ZIP records verbatim. Changed entries use stored payloads with recomputed CRC/sizes and no old descriptor; all central offsets are rebased. Filename bytes, timestamps, entry attributes, extra fields and comments remain retained. Rewriting an edited entry can increase ZIP size; no-op still returns exact original ZIP bytes.
4. Revalidate the candidate archive and compare all OPC payloads/content types/relationships. Only the selected XML part payloads may differ, with no added/removed parts or hidden truncated comparison.
5. Reimport the actual candidate through structure/content projection and compare its **entire** mapped Core state with the prepared snapshot. Recheck current document freshness before exposing the plan.

Unreferenced/unknown parts, shared image bytes, chart/workbook pairs, tables, notes and relationships remain retained because no resource deletion or identity rewrite is performed. This does not authorize garbage collection or establish opaque-feature semantic fidelity. Candidate failure never updates source bytes, the authoring document or the imported baseline.

A [mixed-image regression](../../packages/pptx-parser/test/text-save.test.mjs) exercises a shared PNG, an uninterpreted SVG extension and a real WMF alongside three successive text edits in both dialects. Each round compares every unchanged part and the exact source XML outside the edited text. A browser/Node check verifies the first save's bytes and report. Image geometry changes, replacement-asset changes and mixed text/image edits remain blocked; this is preservation evidence for that text-only fixture, not image replacement support or a complete unknown-reference proof.

## Budgets and lifecycle

Captured [session limits](import-session.md#lifetime-concurrency-and-budgets) additionally provide:

| Limit | Default | Meaning |
| --- | --- | --- |
| `maxSaveAnalysisWork` | 16 Mi units | Diff values/path characters, source target and guard analysis. |
| `maxSaveChanges` | 10,000 | Complete change/diagnostic reports and per-part requested edits. |
| `maxSaveValidationBytes` | 128 MiB | Combined source/candidate expanded OPC payloads during comparison. |
| `maxSaveValidationWork` | 512 Mi units | Exact payload comparison work. |

Original XML budgets also apply to edited XML, and per-input ZIP/OPC budgets apply throughout. The private character-edit profile limits replacement text to 2 Mi UTF-16 units per part; source/new XML has the captured input/character/node/pending-token limits. Raw relationship closure retains its own default roots/declaration/part/selected-byte/work budgets; its finding cap uses `maxSaveChanges`.

An edit preparation reserves the full `maxPreparedOutputBytes` before analysis/candidate work, counts against `maxLivePlans` and shrinks to actual encoded size when ready. This conservative reservation makes concurrent output work explicit. XML replacements together and final ZIP output each obey the per-plan byte limit. Temporary source payloads, parser trees, snapshots and comparison reports consume additional memory; these are not hard peak-memory limits.

Caller cancellation and session disposal propagate into child parsing, XML edits, CRC/copy work and validation. A session shares one abort relay per caller signal and removes it when work finishes; session disposal detaches caller listeners synchronously while children unwind. Failure/cancellation/staleness releases the reservation and permits an independent retry. Plan bytes are detached; returning/saving them does not consume the plan or implicitly rebase the source.

Resource exhaustion throws `PptxImportError("save-limit-exceeded")`; candidate inconsistency throws `save-candidate-invalid`. These errors return no plan or partial output.

<!-- doc-test: docs/examples/text-edit-save.ts -->
```ts
import { importPptx, type PptxSaveReport } from "@pptkit/pptx-parser";

export async function editFirstImportedRun(
  source: Uint8Array,
  text: string,
  signal?: AbortSignal,
): Promise<{ bytes: Uint8Array; report: PptxSaveReport }> {
  const session = await importPptx(source, {
    projection: "content",
    editSaveProfile: "text-runs-v1",
    ...(signal ? { signal } : {}),
  });
  try {
    const slide = session.document.slides[0];
    const element = slide?.elements.find(item => item.type === "text");
    if (!slide || !element || element.type !== "text") {
      throw new Error("No imported top-level text box.");
    }
    const updated = structuredClone(element);
    if (typeof updated.content === "string" || !updated.content[0]?.runs[0]) {
      throw new Error("No imported first run.");
    }
    updated.content[0].runs[0].text = text;
    const index = slide.elements.indexOf(element);
    slide.removeElement(element.id);
    slide.insertElement(index, updated);
    const prepared = await session.prepareSave({
      mode: "edit",
      ...(signal ? { signal } : {}),
    });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    return await session.save(prepared.plan);
  } finally {
    session.dispose();
  }
}
```

## Evidence and open gates

- [Parser text-save tests](../../packages/pptx-parser/test/text-save.test.mjs): both dialects, direct/empty/CDATA/mixed atoms, MCE origins, full classifications, range guards, shared-source reconciliation, lifecycle/freshness/consent, exact budgets, reservations, abort relay cleanup, failure retry and three save/reopen rounds.
- The same suite rejects visually unchanged splits across combining-character and surrogate boundaries and a merge of duplicate runs in both dialects; restoring the baseline returns to exact-byte no-op. This is a guard against unproven topology changes, not support for editing them.
- [ZIP writer tests](../../packages/pptx-parser/test/archive-replacement.test.mjs): independent headers/payload/CRC inspection, both descriptors, reversed directory order, empty archives, UTF-8 names, extra fields/comments/timestamps, corrupt source rejection and exact budgets/cancellation.
- [Cross-package tests](../../scripts/text-save-parity.test.mjs): shared images, charts/workbooks, tables, notes, grouped text, untouched opaque/unreached content, UTF-16BE production saves, mixed-plan atomic blocking and three rounds.
- [Browser ESM/global](../../scripts/browser-bundles.test.mjs) and installed [Node](../../tests/consumers/node.mjs)/[type](../../tests/consumers/types.ts) consumers exercise the public profile. Browser DOM independently verifies raw source element indices; Node/browser output bytes agree.
- [Native experiment](../../scripts/inspect-text-edit-save.mjs) and [evidence](../testing/text-edit-save-evidence.json): one named edit is independently checked with Python ZIP/XML, LibreOfficeDev headless resave and PDF text inspection. The new and control text survive. PPTKit retains the opaque extension/orphan; the producer resave drops them. This is recorded as producer loss, not a full-fidelity pass. Run `node scripts/inspect-text-edit-save.mjs [soffice-executable]`.
- [PowerPoint interactive control](../testing/native-v1-interactive.md): one V1 candidate was edited in PowerPoint 16.104, saved and reopened three times, then edited again through PPTKit and reopened three more times. The application removed the orphan, and the first interactive/final PDFs differ slightly in glyph positions despite equal pixels; the post-native candidate's first/final PDF records agree. These are named-fixture observations, not a general native fidelity or reference-safety guarantee.

Broader interactive editing and PowerPoint coverage, repair-dialog behavior, complete appearance/fidelity, full source reference semantics, fields/table text, broader shape/notes editing beyond the separate V2 character profile, styles/geometry/images, structural operations, source rebasing, Node atomic file replacement and nonempty loss authorization remain open. IMP-07/08/09 and the complete backlog are not closed by this profile.


## Opt-in stable paragraph targets

The separate [paragraph-stable-v1 target profile](paragraph-target-save.md) can exempt proven direct shape paragraph-index references from their range guard during scalar edits. This requires explicit `textTargetSaveProfile`, raw owner/paragraph proofs and unchanged topology; all default behavior above, character/opaque guards and other save checks remain in force.
