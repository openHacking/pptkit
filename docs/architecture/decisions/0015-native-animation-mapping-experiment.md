# ADR 0015 — Native animation specimens before effect admission

Status: research observations under ADV-06.2.1. No public/native animation support admission.

## Specimens and sources

The private [fixture builder](../../../scripts/research/native-timing-fixture.mjs) constructs one hand-authored timing tree on each of four slides: `animEffect` fade entrance, `animRot` clockwise 360-degree emphasis, `animEffect` fade exit and `animMotion` straight path. Each behavior targets the fixture's single named drawing, has 500 ms duration and sits inside a tmRoot/mainSeq/clickEffect wrapper. This is a reproducible research package, not a general serializer or a schema-certified animation preset.

Source references are Microsoft's [AnimateEffect](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animateeffect), [AnimateRotation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animaterotation), [AnimateMotion](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animatemotion) and [Office timing-tree restrictions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/1e399b8c-b3c4-41b3-970c-510909879f59). The rotation uses `by="21600000"` (60,000 units/degree) and attribute `r`. Motion uses layout origin, `M 0 0 L 0.25 0 E` and explicit relative path editing. These references describe storage, not observed native playback.

## Experiment and observations

After building, run `node scripts/inspect-animation-mapping.mjs [soffice-executable]`. The runner injects named specimens into a repository-generated deck, verifies public exact no-op, saves through an isolated LibreOfficeDev profile and reopens the saved PPTX to PDF. Python ZIP/ElementTree independently asserts each source behavior/parameter/target/duration, resolves target IDs to names and rejects duplicate timing IDs. It inventories native conditions and behavior records without requiring equality, and checks saved-output PDF has four pages. Hashes, exact application/environment and logs are retained in [the evidence](../../testing/animation-mapping-evidence.json).

LibreOfficeDev 26.8.0.0.alpha0 retains the entrance/exit filter, rotation amount, line path, named targets and 500 ms durations. It drops motion `pathEditMode` and rewrites time IDs/conditions; complete raw timing inventory equality is false. This headless experiment did not observe initial hidden state, actual clicks, sequence/parallel playback, final visibility, motion coordinate interpretation, native property editing, repair dialogs, PowerPoint or visual fidelity. Subsequent [limited PowerPoint observations](../../testing/animation-playback-protocol.md#preliminary-powerpoint-observation) are recorded separately without retained video. Neither evidence closes ADV-06.2.1 or authorizes a public effect API.

## Preservation and portable evidence

[Two portable integration groups](../../../scripts/animation-mapping-parity.test.mjs) verify raw native storage independently through Chromium DOM, unique timing IDs/resolved targets, and three rounds of unrelated mapped shape-text edits. Original no-op is exact; timing XML remains byte-identical through the character edits, and every non-slide payload remains identical. This establishes retention for these whole-object-target specimens, not semantic editing of timing trees or arbitrary range references. Existing raw text/subobject guards continue to block unsafe text edits.

The research fixture/module is imported only by test/tooling scripts; no OOXML, live timing state or new dependency enters Core/Layout/Renderer. Next admission requires actual native event/visibility/path evidence, more timing/source variants, initial/final-state decisions and target reconciliation. Unknown timing must remain source-owned or block relevant edits; none may be silently regenerated as Fade. All broader/inherited tasks and the full goal remain open.
