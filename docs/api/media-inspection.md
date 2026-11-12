# Bounded media inspection

`@pptkit/media-inspection` contains environment-neutral media structure checks
shared by format adapters. `validateJpegStructure(bytes, signal?)` accepts a
`Uint8Array` (including Node `Buffer` and cross-realm `Uint8Array`) and returns
`{ precision }` for a supported JPEG. The function snapshots up to 64 MiB of
input, checks cancellation between parsing chunks, and does not mutate the
caller’s bytes. Concurrent writes while a shared buffer is being copied are not
an atomic snapshot.

The accepted profile is standalone Huffman DCT JPEG with SOF0, SOF1 or SOF2,
8- or 12-bit precision where allowed by the frame type, bounded marker count,
valid segment lengths, frame/component declarations, quantization and Huffman
tables, supported scan ordering, restart sequence and a final EOI with no trailing
bytes. Unsupported or invalid structures throw; oversized input or marker count
throws `RangeError`. An aborted signal propagates its reason. The check does **not**
decode entropy coefficients, verify all pixel values, color profiles, orientation,
image fidelity or compatibility with a particular reader. Acceptance is not a
complete JPEG validity or safe-decoding certification.

PPTX export uses this same check for its existing `jpegValidation: "structure"`
option. PDF can apply it through its own stricter 8-bit policy; neither consumer
depends on the other format package.

`inspectPng`, `inspectApng` and `inspectGif` are the bounded structural
inspections also re-exported by the PPTX exporter; `inspectGifImageData` is the
additional strict GIF LZW-checking entry. They snapshot their input, reject shared
buffers and require an `instanceof Uint8Array` input in the current realm. PNG
inspection checks chunk framing and CRC but does not inflate pixels; APNG checks
animation-control sequencing and frame bounds. GIF inspection counts image
frames and relevant extensions without decoding LZW; `inspectGifImageData` also
checks its LZW stream. These functions do not certify rendering or playback.
`captureImageInput(input, format, maxBytes, signal?, limitName?)` exposes the
same bounded snapshot helper: `maxBytes` must be a nonnegative safe integer,
labels must be nonempty strings, shared buffers are rejected, and cancellation
propagates its reason. Callers choose a suitable byte limit.

`inspectStaticImage(input, options?)` classifies a bounded image by its bytes,
not its declared MIME, returning `{ format, status }`. A PNG is `static` only
after the PNG/APNG checks find no animation control and all chunks belong to a
named safe set; unrecognized chunks remain `unverified`. APNG is
`animation-declared`. GIF is `static` only with one image frame, no looping,
plain-text or unknown extensions, dangling control, user-input flag, reserved
bits, or disposal beyond 0/1; multiple image frames are `animation-declared`,
and ambiguous single-frame structures are `unverified`. JPEG is `static` only
when it passes the supported JPEG structure profile. Other signatures are
`unknown`/`unverified`; malformed recognized structures throw. This checks
structure, not image pixels, entropy decoding, frame identity or all possible
consumer-specific extensions. It snapshots input and may make additional bounded
copies. `maxBytes` defaults to 64 MiB and must be between 1 and 64 MiB; cancellation
is cooperative and synchronous work cannot be fully preempted.
