# Experimental stable paragraph-target preservation

`textTargetSaveProfile: "paragraph-stable-v1"` permits existing scalar run-text edits on a drawing with proven paragraph-index animation references. It requires content projection and an explicit `text-runs-v1`, `v2` or `v3` edit profile. Root, browser and Node file imports accept the same option. Omitting it retains the existing conservative range guards. It can be combined with `firstSlideNumberSaveProfile: "attribute-v2"` without granting additional timing or topology edits.

```ts
const session = await importPptx(bytes, {
  projection: "content",
  editSaveProfile: "text-runs-v2",
  textTargetSaveProfile: "paragraph-stable-v1",
});
```

See the [typed example](../examples/paragraph-target-save.ts) and [ADR0039](../architecture/decisions/0039-stable-paragraph-target-preservation.md).

## Admitted transformation

Only already-bound scalar text atoms may change. The existing edit profile still fixes all paragraph/run topology, ordering, IDs, styles, properties, geometry and assets. Parser preserves the complete original timing XML. It neither emits a new effect nor rewrites a range.

The admitted raw target is exactly `p:spTgt@spid / p:txEl / p:pRg@st,@end`, with no extra attributes, child alternatives, comments or non-whitespace character content at these three target nodes (namespace declarations and whitespace are allowed). The indices must be ordered, nonnegative supported decimal integers inside the source paragraph count. Numeric spelling is retained; lookup uses the existing content projection's numeric drawing identity domain, including optional plus/leading zeros/XML whitespace. This is not general OOXML schema validation or the separate timing inventory's stricter canonical lookup profile.

The source owner must be a uniquely identified raw `p:sp` with one direct text body whose element children are only DrawingML `bodyPr`, `lstStyle` and direct paragraphs. Every imported paragraph must correspond to one of these raw direct paragraphs. Actual source text atom indices must belong to that shape and paragraph. At least one run must be bound. All raw `cNvPr` identities participate in ambiguity checks, including inactive/extension branches, lexical aliases and negative-zero collisions. Proofs are part-local; identical IDs on other slides do not merge. Nested group children can qualify independently.

These checks establish that replacing text inside a paragraph cannot change which raw paragraphs the preserved indices select. Unicode length, combining marks and newlines inside a scalar run do not change XML paragraph identity. Paragraph/run insertion, removal, property changes or reordering that changes normalized topology/properties remain blocked by full-candidate analysis. Scalar changes to otherwise identical paragraphs are changes to their existing text atoms; the API does not infer semantic paragraph moves from strings.

## Guards and lifecycle

| Source/candidate | Result |
| --- | --- |
| Proven exact paragraph target and admitted scalar changes | Scoped-preservation plan; unchanged range/timing XML |
| Character range, mixed targets, unknown target attributes/children, malformed or out-of-bounds paragraph range | Existing `save-text-range-reference` blocking diagnostic |
| Raw owner ambiguity or MCE/foreign wrappers among direct text-body children | No paragraph proof; existing range guard remains |
| Notes, table, SmartArt, chart or group-level range target | No new authority; owning profile's existing guards apply |
| Another unsupported mutation or another guarded reference on the edited owner | Entire mixed plan blocked; no partial output |
| Guard belongs only to an unrelated owner/part | Does not block an otherwise admitted owner |
| No normalized changes/restored baseline | Use original-byte no-op |

Inactive raw range references are checked too. A transparent inactive paragraph target can use the same proven unique owner; it does not become an effective animation. If any raw reference on the owner still requires a guard, a supported paragraph reference cannot cancel it. Missing/unresolvable owners retain the existing part-wide text guard.

Bytes/options are captured before asynchronous work, including Node file reads. Proofs use a cumulative `maxReportFindings` cap over eligible bindings, candidate owners and direct paragraph records across all inspected source parts. Existing binding/guard/report and XML/package limits also apply; this is not a hard peak-memory bound. Proof traversal cooperatively checks cancellation. Proof exhaustion throws `import-limit-exceeded` with no session. Save analysis, reservations, cancellation, failure retry, stale plans and exact consent retain the [text-save contract](text-edit-save.md).

The lexical writer changes only approved `a:t` character regions, compares the actual package and reprojects the complete mapped Core candidate before returning a plan. No save rebases the original session: reimport returned bytes for the next round.

## Evidence and native limits

[12 Parser groups](../../packages/pptx-parser/test/paragraph-target-save.test.mjs) cover both dialects/all text profiles, three Unicode save rounds, numeric spellings, malformed/mixed/opaque targets, inactive identities, direct-body ambiguity, multiple parts, groups, notes/tables, topology guards, quotas/cancellation/retry and Node file capture. [Two browser groups](../../scripts/paragraph-target-save-parity.test.mjs) compare complete Node/ESM/global bytes over three rounds and independently verify ranges, source owners, paragraph content/order and timing with Chromium DOM. Installed runtime/types consumers exercise the opt-in contract.

Run `node scripts/inspect-paragraph-target-save.mjs [soffice-executable]` for the optional [independent/native experiment](../../scripts/inspect-paragraph-target-save.mjs). The [manifest](../testing/paragraph-target-save-evidence.json) records three source-save rounds checked by Python ZIP/ElementTree: exact unselected payloads and complete XML fingerprints except two approved text atoms. Deliberately changed ranges and unrelated attributes are rejected by that independent verifier.

The named LibreOfficeDev version retains the three paragraph texts but changes this specimen's range from `0..2` to `0..0` in each native resave. The unedited-source control exhibits the same rewrite. This is a producer observation for this fixture; the native range result remains **partial**, and cannot certify preserved playback. Its saved output reopens as a two-page PDF containing the changed and control text. PDF text, headless reopening and source index preservation do not prove animation playback, visual fidelity, repair-dialog absence or interactive effect editing.

This advances fixed-paragraph source preservation under ADV-07.8 / ADV-06.2.1.3. Character/word offset remapping, paragraph insertion/deletion/moves, full native semantics and real playback remain open. This profile adds no Core event timeline, public animation effect, animated SVG or source structural editing.
