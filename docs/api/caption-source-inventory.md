# Caption source inventory

`inspectPptxCaptionTracks(bytes, options?)` is a read-only, bounded source inventory in `@pptkit/pptx-parser`. It recognizes the specific PowerPoint 2017 `tracksInfo/trackLst/track` form observed under a picture's Office 2010 media extension. Each record contains the slide and raw XML location, lexical picture/track identities, language, label, relationship and a status. It does not create Core caption objects, infer a default or selected track, validate cue syntax, extract track text, enable editing, or claim playback support.

```ts
import { inspectPptxCaptionTracks } from '@pptkit/pptx-parser';

const result = await inspectPptxCaptionTracks(pptxBytes, {
  captionLimits: { maxTracks: 100, maxTrackBytes: 500_000 },
});
for (const track of result.tracks) console.log(track.source.part, track.language, track.status);
```

`view: "raw-only"` applies to every track record. The scan covers raw XML of slides in the effective presentation order, including inactive markup-compatibility branches. `pictureView` separately reports whether the owning picture's exact raw source node occurs in the effective drawing tree (`effective-drawing`) or only in raw XML (`raw-only`). This is picture presence, not visibility or track activation: the media extension remains opaque, and the result does not interpret compatibility choices within it. Masters, layouts, unlisted slides, audio or non-picture owners, other extension shapes and malformed chains outside the observed form are not enumerated. An absent record is therefore not proof that a presentation has no captions. Duplicate IDs and shared targets remain separate source records and are not resolved into semantic identities.

An internal whole-part relationship of the observed track type is inspected by content type, fatal UTF-8 decode and the `WEBVTT` header followed by LF or CRLF. `vtt-header-only` is not cue validation. Other statuses identify invalid UTF-8 or header, unsupported content type, an external relationship or an unsupported relationship. External targets are reported but never fetched. A structurally invalid package, including an internal target missing from the ZIP, fails package inspection rather than producing a partial inventory. Track content is not returned.

The default limits are 1,000,000 work units, 10,000 tracks, 1,000,000 output text characters, 1 MiB per distinct track and 8 MiB total distinct track bytes. Positive safe-integer overrides are explicit; the usual ZIP, XML and OPC limits apply independently. Shared parts are inspected and charged once, while every source record and path is charged. Limits fail the whole call; abort is checked through the existing package and traversal checkpoints. Package parsing retains bounded package payloads before the per-track limits are applied, so these limits are not a peak-memory guarantee. Results and records are frozen.

Tests use the [PowerPoint 16.104 sample](../testing/native-captions.md), malformed header and UTF-8, wrong MIME/relationship, an external target, missing package part, shared targets, duplicate object IDs, limits, cancellation and selected versus inactive MCE picture branches. Node and Chromium ESM/global outputs match with network blocked. This inventory does not change import or no-op preservation behavior; more producers, track-level active-branch interpretation, audio, language semantics, default selection and actual subtitle timing remain open.
