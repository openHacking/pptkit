# Experimental shape-text and speaker-notes preservation saves

Content import with `editSaveProfile: "text-runs-v2"` admits scalar character changes on existing bound text-box runs, mapped shape-text runs and one directly mapped speaker-notes body per slide. The version is explicit: `text-runs-v1` retains its text-box-only scope and original report shape. Default imports still block edits; structure projection rejects all edit profiles.

The later [V3 table-character profile](table-text-save.md) extends owner coverage separately without changing this V2 contract. The [v1 preparation, validation, budgets and lifecycle contract](text-edit-save.md) applies to all profiles. V2 extends ownership only. Paragraph/run counts, formatting, actions, geometry, assets, identities and document structure must remain equal to the imported normalized baseline. A supported text edit mixed with an unsupported mutation blocks the entire plan. Save does not advance that baseline; reimport returned bytes for a new session.

## Owners and source identities

| Owner | Core location | Source mapping | Public provenance |
| --- | --- | --- | --- |
| Text box | `element.content[p].runs[r].text` | Existing mapped direct text box, including mapped groups | `textRuns`, `owner: "text-box"`, real `slideId` and `elementId` |
| Shape text | `element.text.content[p].runs[r].text` | Existing mapped shape text body, including eight supported presets and mapped groups | `textRuns`, `owner: "shape-text"`, real `slideId` and `elementId` |
| Speaker notes | `slide.notes[p].runs[r].text` | Exactly one direct body placeholder with a local text body in that slide's notes part | Separate `notesTextRuns: readonly ImportedNotesTextRunSource[]`, `slideId`, no fabricated `elementId` or public `owner` |

V2 emits both frozen arrays, including empty arrays. V1 emits only its existing `textRuns` array, with no `owner` tags; default imports emit neither array. All records include paragraph/run indices, part/path/source object identity, bound/unmapped status and the raw element index when available. Private baseline text, namespace and notes discriminators are omitted.

Binding follows effective-to-raw MCE origins, never matching by text, shape name, prefix or relative position alone. Repeated characters and equal drawing IDs in different parts remain distinct. Known range guards are scoped by **source part and source drawing identity**, so a notes target cannot accidentally block or authorize an unrelated slide shape. Inactive raw targets are scanned too. Whole-shape targets remain unchanged; range/sub-object or opaque targets on the edited owner block.

Notes-master prompts and inherited text are never synthesized into editable notes. Multiple bodies, absent local bodies, nested placeholders and unmapped shape families gain no bindings. Shared notes parts across slides retain the existing explicit import rejection. Private shared-atom reconciliation still requires all mapped views to agree. Formatting from unresolved notes-master/theme inheritance remains a projection limitation; preserving source bytes does not infer those values.

Tables, chart labels/data/workbooks, SmartArt, fields, breaks and synthetic insertion runs do not gain permission through V2. Unmodified opaque effects, unknown extensions, relationship sets and other source payloads remain retained. This retention does not interpret arbitrary implicit references or prove their application semantics.

## Saving shape text and notes together

The example edits the first run on an existing top-level mapped shape and the first direct notes run. Preparation remains authoritative: provenance does not guarantee approval, and source guards or unsupported concurrent mutations may block it.

<!-- doc-test: docs/examples/text-owner-save.ts -->
```ts
import { importPptx, type PptxSaveReport } from "@pptkit/pptx-parser";

export async function editShapeAndNotes(
  source: Uint8Array,
  shapeText: string,
  notesText: string,
): Promise<{ bytes: Uint8Array; report: PptxSaveReport }> {
  const session = await importPptx(source, {
    projection: "content",
    editSaveProfile: "text-runs-v2",
  });
  try {
    const slide = session.document.slides[0];
    const element = slide?.elements.find(item => item.type === "shape" && item.text);
    if (!slide || !element || element.type !== "shape" || !element.text) {
      throw new Error("No imported top-level shape text.");
    }
    const shape = structuredClone(element);
    const notes = structuredClone(slide.notes);
    if (!shape.text || typeof shape.text.content === "string" || !Array.isArray(notes) ||
        !shape.text.content[0]?.runs[0] || !notes[0]?.runs[0]) {
      throw new Error("Both existing first runs are required.");
    }
    shape.text.content[0].runs[0].text = shapeText;
    notes[0].runs[0].text = notesText;
    const index = slide.elements.indexOf(element);
    slide.removeElement(element.id);
    slide.insertElement(index, shape);
    slide.setNotes(notes);
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    return await session.save(prepared.plan);
  } finally {
    session.dispose();
  }
}
```

## Evidence and remaining gates

- [Parser owner tests](../../packages/pptx-parser/test/text-owners-save.test.mjs): both dialects, all mapped presets/groups, raw branch origins, source-scoped references/actions, ambiguous/missing notes, mixed/synthetic atoms, shared-part rejection, whole-plan mutation failures, freshness, combined provenance budgets, UTF-16 and three save/reimport rounds.
- [Mixed-document integration](../../scripts/text-owner-save-parity.test.mjs): generated tables/charts/workbooks/images and two notes parts; independent ZIP/SHA256 proves only selected slide/notes payloads change; exact OPC comparison, graph equality and full candidate reprojection agree over three rounds.
- [Browser ESM/global](../../scripts/browser-bundles.test.mjs): Node-identical bytes/reports/normalized documents, independent per-part DOM indices, stale notes plans and no network requests. Installed [Node](../../tests/consumers/node.mjs) and [declaration](../../tests/consumers/types.ts) consumers cover both owners and immutable notes identity.
- [Native runner](../../scripts/inspect-text-owner-save.mjs), [retained evidence](../testing/text-owner-save-evidence.json): independent Python ZIP/XML checks, LibreOfficeDev headless PPTX resave and two-page PDF text inspection. Edited shape text, notes and controls survive; PPTKit retains unknown extensions/orphan data that the native resave drops. Run `node scripts/inspect-text-owner-save.mjs [soffice-executable]`.
- [PowerPoint interactive control](../testing/native-v2-interactive.md): one mixed two-slide candidate was edited in shape and notes UI, saved and reopened three times, edited again by PPTKit and reopened three more times. The named shape/notes relationships, text, image, workbook, chart cache and links were checked. PowerPoint removed the orphan and its notes print used three pages before and after; initial interactive versus final note rendering differed. The later PPTKit candidate's first/final slide and notes PDF records agreed.

This is bounded progress on **IMP-09.2.1.2**. Broader PowerPoint and interactive editing coverage, repair dialogs, notes-page layout fidelity, complete appearance/reference fidelity, topology/property edits, rebasing, atomic file replacement, general rebuild and the inherited parent gates remain open. Headless producer text retention and the named PowerPoint control do not close those gates.


For drawing shape text, the separate [paragraph-stable-v1 opt-in](paragraph-target-save.md) admits proven paragraph-index preservation during scalar changes. Notes targets retain their conservative range guards; this is not a general subtarget or native playback API.
