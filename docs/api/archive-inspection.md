# Experimental ZIP archive inspection

`@pptkit/pptx-parser` now provides `inspectZipArchive(bytes: Uint8Array, options?: ZipInspectionOptions): Promise<ZipArchiveInspection>` from its root and `/browser` entries, plus `PackageInspectionError`, inventory/limit types and frozen `DEFAULT_ZIP_LIMITS`. The global browser bundle adds these names to `PPTKit`.

Input is checked against `maxInputBytes` using the actual typed-array view, then copied before asynchronous work. Shared-memory-backed views, including buffers from another JavaScript realm, are rejected because an atomic snapshot cannot be guaranteed. Subclass getters and iterators are not used for the size check or copy.

This is the first IMP-02 substrate, **not PPTX import or an OPC validity guarantee**. An empty valid ZIP returns an empty inventory. Per-part XML inspection is available separately through [inspectXmlPart](xml-inspection.md). The separate [OPC inspector](opc-inspection.md) adds content types, relationship targets and package-wide XML budgets. Strict/Transitional markup interpretation and supported semantic conversion remain unimplemented. Do not interpret a successful inventory as proof that the presentation is safe to open or round-trips correctly.

## Implemented profile

The reader accepts a single ZIP32 disk with stored or deflate entries, ASCII names or explicitly UTF-8 names, comments, and signed/unsigned ZIP32 data descriptors. Central entry order may differ from local record order. It verifies local/central agreement for names, flags, methods, sizes and CRC declarations, validates descriptor values and extra-field boundaries, and rejects overlapping/unindexed local ranges, prefixed executables and trailing archive junk. Each entry's actual decoded count and CRC-32 are checked independently.

The implementation follows the record structure in [PKWARE APPNOTE 6.3.10](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT), particularly sections 4.3 and 4.4. Unsupported profile choices are deliberate limits, not claims that other ZIP forms are invalid.

ZIP64, multi-disk/encrypted archives, unsupported flags/methods, alternate Unicode filename extra fields, non-ASCII legacy filename encoding, and entries marked as Unix or OS X symlinks or special filesystem objects are explicitly rejected. Other creator-host-specific file types are not interpreted, and no path is extracted to disk. Unknown extra fields with valid lengths are not interpreted. Signed central-directory structures and other extensions are outside this initial profile.

Paths cannot contain absolute/drive syntax, backslashes, controls, traversal/empty segments, URL query/fragment syntax or nested percent encoding. Duplicate keys are rejected after percent decoding, NFC normalization and lowercase comparison; file/directory aliases and files used as directory ancestors are rejected. Directory names must agree with explicit DOS directory flags and, for Unix and OS X creator systems, regular-file or directory type bits. Unix type bits are not inferred from other creator hosts' high attribute bits. [Go's ZIP reader](https://github.com/golang/go/blob/master/src/archive/zip/struct.go) likewise interprets these two creator systems using Unix mode bits; this inspector deliberately rejects ambiguous directory declarations. This is a conservative package path policy, not a complete OPC URI implementation. No path is extracted to disk.

[Archive tests](../../packages/pptx-parser/test/archive.test.mjs) and [browser integration](../../scripts/archive-directory-type-browser.test.mjs) also exercise OS X type declarations through structure/content PPTX import: ordinary files permit byte-exact no-op saves, links fail as unsupported, and directory/name conflicts fail as invalid packages. A separate [Go comparison](../../scripts/archive-mode-go.test.mjs) checks four synthetic OS X modes against Go's ZIP reader when Go is available; it skips if Go is absent. These tests do not establish behavior for every creator system or native application.

## Budgets and ownership

| Limit | Default |
| --- | --- |
| `maxInputBytes` | 64 MiB compressed input |
| `maxEntries` | 20,000 entries |
| `maxEntryBytes` | 64 MiB per expanded entry |
| `maxExpandedBytes` | 256 MiB total expanded output |
| `maxCompressionRatio` | 1,000 expanded/compressed, with denominator at least one byte |
| `maxStalledInputBytes` | 128 KiB compressed input without decoded progress |

All limits must be finite and positive; byte/count limits are safe integers. Limits are copied before asynchronous work. Input bytes are detached before the first await; returned inventory is detached and readonly in TypeScript. The actual AbortSignal is captured once, so replacing an options property cannot change cancellation identity.

Declared sizes are screened before decoding, but they are not trusted as actual output counts. An independent structural scan counts DEFLATE output before allocating decoded output and enforces the same entry/total/ratio budgets. Deflate then receives compressed chunks no larger than 1 KiB, and decoded chunks are budget-checked before retention. The public inventory API does not retain expanded payloads. Private payload collection is reserved for the future XML/package reader. Decoder working memory and one callback chunk are additional overhead; these logical byte limits are not a hard process-memory sandbox. The stalled-input budget bounds streams that make no output progress. Cooperative task yields permit cancellation during large inputs; header/index work is bounded by input and entry budgets but is synchronous.

Errors return no partial inventory. Malformed options use TypeError/RangeError. `PackageInspectionError` carries `code`, optional `entry` and `message`: `archive-invalid`, `archive-unsupported`, `archive-path-invalid`, `archive-duplicate-path`, `archive-crc-mismatch`, `archive-limit-exceeded`, or `operation-aborted`.

## Remaining verification boundary

An independent [RFC 1951](https://www.rfc-editor.org/rfc/rfc1951.txt) structural scan validates stored/fixed/dynamic blocks, Huffman alphabets and repeats, end-of-block/final-block markers, reserved symbols, and backward distances against available history. The final block must end in the last declared compressed byte: concatenated streams and whole trailing bytes are rejected; unused bits in the final byte are permitted padding. The scanner retains no decoded output and uses no inflater-private state. Both output-producing symbols and empty blocks participate in cooperative cancellation.

The [OPC inspector](opc-inspection.md) adds a bounded logical-name and relationship profile with cumulative XML limits. Broader physical-item mapping remains outside this archive profile. Per-part XML limits are documented in the [XML inspector](xml-inspection.md); dialect/MCE readers and [import/save sessions](import-session.md) are separate higher-level APIs. ZIP structure and CRC validation do not establish XML validity or semantic PPTX compatibility.

## Checked consumer

<!-- doc-test: docs/examples/archive-inspection.ts -->
```ts
import { inspectZipArchive, PackageInspectionError } from "@pptkit/pptx-parser";

export async function inspectArchive(bytes: Uint8Array, signal?: AbortSignal) {
  try {
    return await inspectZipArchive(bytes, {
      limits: { maxInputBytes: 32 * 1024 * 1024, maxEntries: 10000 },
      ...(signal ? { signal } : {}),
    });
  } catch (error) {
    if (error instanceof PackageInspectionError) return { code: error.code, entry: error.entry, message: error.message };
    throw error;
  }
}
```

The package tests use independently built ZIP records and bit-level DEFLATE fixtures, plus corruption, truncation, budget and cancellation cases. An 80-case seeded corpus is cross-checked against Node zlib across compression levels, four strategies and seven window sizes, covering all three supported block types. Corpus tests inspect real PPTKit-generated PPTX envelopes. Installed Node/browser/global consumers exercise the public entries without workspace-private imports.
