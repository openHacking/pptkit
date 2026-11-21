# ADR0065: Caption track research and admission boundary

Status: research decision, 2026-11-21. This does not admit caption creation, semantic import, fresh export or browser playback.

## Evidence

In PowerPoint for Mac 16.104, two WebVTT files were inserted on the first of two video instances sharing an MP4. The first instance was trimmed by 150 ms at the start and 50 ms at the end. After saving and reopening, the caption editor listed English (United States) and a French (France) track with the label `French test`. The other video has no track entry. The single-track and two-track packages, original-byte no-op and independent ZIP/XML inspection are described in [native caption observations](../../testing/native-captions.md).

A separate plain UTF-8 SRT inserted on the other video was converted by this PowerPoint version into a `text/vtt` part with the same two cue times and text. Reopening retained the Spanish track and left the original two tracks intact. A manual trimmed-video playback attempt showed video, but did not establish subtitle selection or timing. See the [experiment](../../testing/native-captions.md#srt-conversion-and-playback-attempt).

The observed video `p14:media` contains an extension with `p173:tracksInfo displayLoc="media"`, a `trackLst`, and ordered `track` elements with GUID-like `id`, `label`, `lang` and `r:embed`. Each relationship has the Office 2017 track relationship type and targets a separate package part of content type `text/vtt`. The saved tracks contain the same cue timecodes as the inputs, including one cue beginning before the trim start. This shows no storage-time offset in this experiment; it does not establish rendering or seek-time offset. The observed XML has no explicit default-track flag; order and language must not be interpreted as a default. Separate media instances can share the MP4 while having different track lists.

## Decision

Keep tracks as source-owned, per-media-instance data. Raw bytes, relationship identity, language spelling, label, order and unknown extensions require preservation until separately admitted. An unchanged source save retains the original package. Current Parser reports the track parts as unsupported; a fresh export of projected content does not preserve tracks and is not a caption round trip. Do not infer a default track, display state or playback timing from this sample. External targets must not be fetched implicitly. Do not introduce generic Core subtitle fields or claim WebVTT cue editing based solely on file extension or a header.

The research inventory searches picture descendants for the observed track namespace and tests UTF-8 plus the WebVTT header; it is not an OPC security boundary, complete MCE inventory, WebVTT parser, validator, or production API. The separate [bounded Parser source inventory](../../api/caption-source-inventory.md) now recognizes the observed extension chain with OPC/XML budgets, relationships and raw source locations. It includes inactive MCE branches and distinguishes whether the owning picture is present in the effective drawing tree, without interpreting the opaque extension's track activation. Unobserved owner/extension forms and semantic interpretation remain open. External targets are never fetched.

A synthetic mixed package made from the retained PowerPoint SRT sample and a generated text object preserves all three tracks through three unrelated text saves, checked by exact part comparison, relationship graph and the independent research inspector. See [the regression](../../testing/native-captions.md#unrelated-text-preservation). This does not admit track edits or prove native playback.

## Outstanding admission criteria

Determine default selection and disabled state with native playback and save/reopen evidence; observe seek, cue boundary, trim and looping against actual video and compare a second producer. Investigate SRT markup and encoding conversion beyond the one plain example, inspect audio and other extension forms, define language handling, broken/shared/duplicate tracks and explicit access to external sources. Only then choose a bounded VTT grammar and lossless handling of unsupported cue settings, and define actual source-relative time mapping. Test browser synchronization and accessible alternatives separately. ADV-04.4.1 and ADV-04.4.2 therefore remain open.

## Alternatives

Inferring a default from the first track would invent behavior not evidenced by XML. Normalizing all tracks to cues now could discard unknown WebVTT syntax, metadata and bytes. Exposing the research helper as the public API would bypass existing OPC limits and MCE policy. Preserving source bytes and explicitly reporting unsupported projection keeps those options available without overstating capability.
