# Speaker notes import and editing

`importPptx(bytes, { projection: "content" })` maps direct speaker-notes paragraphs through the public Parser entry points. Structural projection continues to leave notes in the inspected source only. The [checked example](../examples/notes-import.ts) demonstrates a deliberately fresh projection export after an explicit notes edit; it is not a preservation save.

## Supported content

The reader follows each effective slide's notes relationship and selects exactly one direct `p:sp` with an explicit `p:nvSpPr/p:nvPr/p:ph type="body"` in the notes drawing tree. Names and drawing order do not determine the speaker-notes body. Local offset/extent and shape geometry are not required for text import.

The body reuses the documented [direct rich-text subset](content-import.md): paragraphs, runs, explicit breaks, literal Unicode/whitespace, direct fonts/colors/styles, supported numbering/bullets, paragraph spacing/alignment and supported run click actions. Click relationships resolve relative to the notes part; slide links use effective Core slide IDs and URL links are never fetched or executed. Notes slides' implicit owning-slide relationship is checked separately from explicit click/hover slide relationships. A self-link can share the owning relationship; contradictory or multiple unreferenced owners reject.

| Source | Core projection |
| --- | --- |
| One direct body with local text | Owned rich-text `slide.notes` |
| Explicit body with an empty paragraph | Explicit empty notes input; direct formatting is retained |
| No notes part or no body placeholder | `slide.notes === undefined`; no text is synthesized from the master |
| Multiple direct body placeholders | No guessed body; source-only finding and `notes-body-unmapped` diagnostic |
| Body with no local text | Source-only finding; inherited text is not synthesized |
| Header/footer/date/slide-number/slide-image placeholders | Remain source-only |
| Arbitrary or nested shapes containing text | Remain source-only |
| Fields, page geometry/artwork, effects, notes master/theme/list-style inheritance | Separately reported as unsupported or unresolved |

Drawing identities are checked within each notes part, including omitted drawing families and group descendants. Identical numeric IDs in different drawing parts are valid. The reader reports the body as `ImportSourceIdentity.kind === "notes"`, with its **owning Core slide ID**, source notes part, body drawing ID and effective XML path. It does not create a fake slide element. Consumers that switch exhaustively over identity kinds must handle `"notes"`.

Each mapped body/click feature receives separate operation findings: Core creation/editing, import and fresh export are editable for the mapped fields; source no-op is preserved; slide preview and edit-save are unsupported. Standalone slide SVG does not render or expose speaker notes. Unknown body properties receive `notes.unmapped-property` findings with source identity/path. `notes.inheritance-and-defaults` marks unresolved inheritance as import fallback; default Core styling is not inferred source appearance.

## Editing and fresh export

`slide.setNotes(content?)` replaces notes with an immutable owned snapshot; `slide.setNotes()` clears them. Previous inputs, returned values and normalized snapshots remain independent. The Core `PresentationSlide` interface now requires this method; custom implementations must implement it. Existing notes constructor inputs and getters retain their behavior.

Notes edits participate in full normalized save-plan freshness checks. Editing notes invalidates an earlier no-op plan; exact semantic restoration restores eligibility. Session disposal leaves the Core notes available for normal authoring. Default notes preservation edits and all rebuilds remain blocked. The explicit [V2 owner profile](text-owner-save.md) permits existing bound direct notes-body run characters with fixed topology, source guards and full candidate validation. The [V1 text-box profile](text-edit-save.md) retains its original scope and unchanged notes payloads.

Fresh Exporter output writes normalized notes paragraphs, including all-empty paragraphs with styling or actions. Previously a text-length check discarded all-empty content. The plain default empty paragraph synthesized when notes are absent is omitted. Authoring still distinguishes `undefined` from an explicit empty input; normalized semantics and fresh export do not preserve that source shape-presence distinction. Exact source no-op retains it.

Fresh export uses the Exporter's notes page/master defaults. It does not reconstruct imported notes-page geometry, inherited style provenance, unknown shapes or master content. Retaining text and direct formatting must not be described as complete notes-page round-trip.

The current default fresh layout has [native slide-image and body placeholders](../testing/authored-notes-pages.md) with fixed notesSize and aspect-preserving image geometry. This does not change source import/preservation permissions. The original PowerPoint empty-run loss is retained as historical evidence; [the insertion-property correction](../testing/notes-insertion-styles.md) now verifies empty-mark styles and actual typing. An imported native paragraph containing only insertion properties still has no bound a:t character owner for source scalar writes.

## Verification

Focused Parser tests cover both namespace dialects, direct/empty/absent/ambiguous bodies, role selection, rich text, fields/effects, duplicate/invalid drawing IDs, notes-local URL/slide links, safe-action policy, backlink conflicts, MCE and aliases, malformed mapped styles, budgets/cancellation, stale plans and semantic undo. Core tests cover replacement/clear, isolation, copying/self-link remapping and failed snapshot atomicity.

Two cross-package cases inspect notes XML/relationships over three fresh projection rounds and after note edits/copy/delete. Chromium compares normalized documents, complete reports and output bytes with Node, exercises note edits and exact undo, checks notes are absent from slide SVG, and observes no link loading. Installed tarball consumers exercise the public setter, notes source identities and notes self-links.

The optional `node scripts/inspect-notes-import.mjs [soffice-executable]` creates a LibreOffice-produced fixture and independently inventories notes-body text and clicks using Python ZIP/XML. The [recorded evidence](../testing/notes-import-evidence.json) compares all produced notes text, the supported action subset, exact no-op bytes and fresh projection output; it also records a slide-PDF reopen. LibreOffice splits an authored line break into paragraphs, strips tooltips, normalizes URL escaping and converts two notes slide-jumps into relative external hyperlinks. Those two actions remain source-only with diagnostics, and are not present in the fresh projection. This is not native notes-page rendering, interactive editing/navigation, PowerPoint certification or preservation edit-save evidence. IMP-03.3 remains partial.

The separate opt-in [notes master text profile](notes-style-import.md) now resolves bounded notesStyle/type-matched placeholder list defaults and the independent notes theme. The direct-only default described above remains compatible; full notes-page geometry/artwork and native appearance remain open.

The separate [notes page source inventory](notes-source-inventory.md) now reads roles, raw geometry/artwork and stored field caches with explicit unresolved/evaluation markers. It leaves the direct-body import contract and existing save permissions unchanged.
