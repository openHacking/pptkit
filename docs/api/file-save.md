# Node.js atomic file import and output

The experimental `@pptkit/pptx-parser/node` entry imports a local PPTX into a preservation session and publishes a prepared candidate by create-only or snapshot-guarded replace. The package root, `/browser` entry and browser bundles contain no Node filesystem imports. Use the [byte session API](import-session.md) when the application owns a virtual filesystem or downloads the result itself.

```ts
import {
  importPptxFile,
  writePptxFile,
  type ImportedPptxFile,
  type PptxFileSnapshot,
  type PptxFileWriteOptions,
} from "@pptkit/pptx-parser/node";
```

## Import a stable source

`importPptxFile(path, options)` captures import options before its first filesystem await, canonicalizes the parent directory, then reads the named regular file without following a final symlink. Reads and metadata checks require `O_NOFOLLOW` and `O_NONBLOCK`; the import rejects with `file-platform-unsupported` when those platform flags are unavailable.

The captured options include `masterArtwork` and `pictureSourceProfile`; their output is the same as byte import for the same source. The master artwork profiles are checked across both dialects by [file/byte parity tests](../../packages/pptx-parser/test/master-artwork-import.test.mjs), including caller mutation after invocation.

```ts
declare function importPptxFile(
  path: string,
  options: PptxImportOptions,
): Promise<ImportedPptxFile>;

interface ImportedPptxFile {
  readonly session: ImportedPresentation;
  readonly file: PptxFileSnapshot; // process-local overwrite receipt
}

interface PptxFileSnapshot {
  readonly path: string;       // canonical absolute path
  readonly byteLength: number;
  readonly sha256: string;
}
```

The returned session uses the same content/structure projections and edit profiles as `importPptx`. The source is captured at a fixed byte length in chunks of at most 64 KiB. Start, end and re-resolved name identities must agree, including device, inode, size, mode, modification time and change time. The complete bytes must equal the size at capture and must still have the same descriptor/name identity and content. Growth, shrinkage, replacement and cancellation reject with `file-source-changed` or `operation-aborted`; no usable session is returned.

`PptxFileSnapshot` is a frozen, nominal, process-local receipt. It is issued only by one Parser Node adapter instance and the adapter keeps its native identity metadata privately. Copying its visible fields, JSON serialization or using a snapshot from another adapter does not grant overwrite authority. A session baseline never advances after writing its file; re-import output when a new source baseline is intended.

## Create or replace a file

```ts
declare function writePptxFile(
  session: ImportedPresentation,
  plan: PptxSavePlan,
  path: string,
  options: PptxFileWriteOptions,
): Promise<PptxFileWriteResult>;

type PptxFileWriteOptions = { signal?: AbortSignal; acceptLossIds?: readonly string[] } & (
  | { mode: "create" }
  | { mode: "replace"; expected: PptxFileSnapshot }
);
```

The write first validates the save session, plan, consent, ownership, liveness and normalized freshness through the [session contract](import-session.md), then obtains detached candidate bytes. It creates a random `0600` temporary regular file in the same canonical parent directory, writes it in chunks of at most 64 KiB, and calls `fsync` on the file. It reads and hashes the full flushed temporary file before publication. A short write is completed; zero progress and same-length corruption reject. The named temporary identity is checked again immediately before commit.

`create` publishes by hard-linking the flushed temporary file to the new name without replacement, then removes the temporary name. An existing file or symlink remains untouched and rejects with `file-target-exists`. This preserves an ordinary default `0600` mode subject to the platform's link permissions.

`replace` requires an authentic expected snapshot and checks that the canonical path matches. It rejects a changed file, equal-byte new inode, missing target, symlink, special file, path parent change and content conflict with `file-target-changed`. The final freshness check resolves the complete source identity, name and SHA256 twice. After creating and flushing the candidate, the adapter rechecks the target, its named parent and temporary identity. It preserves ordinary permission bits (`mode & 0777`) on the new inode. It does not preserve owner, ACLs, extended attributes, alternate streams or creation/birth time.

Both modes call the public synchronous `session.assertSavePlanCurrent(plan, { acceptLossIds? })` immediately before dispatching the link or rename. `session.save()` is asynchronous, so the adapter cannot rely on its result for a final no-await freshness check. The guard performs the same ownership, consent, disposal and full normalized snapshot checks without copying or consuming the plan. Session edits, plan/session disposal and semantic changes in this final boundary reject before publication.

## Commit and failure outcomes

The following failures occur before publication. The original target remains unchanged, the temporary name is removed when possible, reservation/path locks are released and the same plan can be retried:

- Source read, options, snapshot, platform, save preparation or final plan freshness failure.
- Parent/temp open, write, file sync, file close, readback/hash, final parent/temp identity or target verification failure.
- Cancellation observed before syscall dispatch.
- `create` losing the exclusive publication race or `replace` finding a detected target change.

An injected filesystem failure rejects with `PptxFileError` code `file-io`; source, target, consent, contention and platform failures use distinct codes. If precommit cleanup itself also fails, the error is `file-cleanup-failed`, has `committed: false`, and exposes the `temporaryPath` when it still exists. The original target is not reported as replaced.

Publication is the commit point. Cancellation after the link/rename syscall is dispatched, temporary cleanup failure, directory `fsync` failure or a directory-handle close failure returns `{ status: "committed" }`. The reused result reports `directorySynced: false` when directory synchronization failed and immutable `warnings` for postcommit failures. Cancellation cannot roll back a committed file. Callers must inspect the output or its receipt even when warnings are present.

```ts
interface PptxFileWriteResult {
  readonly status: "committed";
  readonly file: PptxFileSnapshot;
  readonly report: PptxSaveReport;
  readonly directorySynced: boolean;
  readonly warnings: readonly PptxFileWriteWarning[];
}
```

Any successful result can authorize a later replace of that exact current file. Snapshots remain reusable only while the named file identity and bytes still match; a successful save does not consume them and does not rebase the session.

## Concurrency and filesystem limits

One shared adapter in the public Node entry fails fast with `file-write-busy` for concurrent writes to the same canonical path or expected source inode. Independent processes or adapter instances cannot share that process-local lock. `create` remains exclusive through the hard-link operation. Replacement is **optimistic**: the second SHA256/stat check detects changes up to the final check, but standard portable Node filesystem calls provide no atomic compare-and-swap between that check and rename.

A different process can therefore replace the target in the final check-to-rename interval and have its update overwritten. The adapter assumes a trusted/stable parent directory and cooperating writers that honor the documented `writePptxFile` contract. Directory replacement, hostile local mutation, remote filesystems with nonportable identity semantics and arbitrary external editors are outside the guarantee. Applications requiring a stronger cross-process transaction must add an external lock/CAS protocol and separately validate it.

Source/input and parent-directory paths are canonicalized; the final existing input must be a regular file. The temporary file and publication target share that parent directory so the rename cannot cross filesystems. There is no weaker flags or symlink-following fallback. Windows has not been verified, and a filesystem that cannot represent the required operations rejects rather than degrading.

## Bounded resource and lifecycle behavior

The Node adapter adds no new byte or decompression budget. `importPptxFile` applies the captured `zipLimits.maxInputBytes` before opening the source. `writePptxFile` uses the existing plan/output reservations, normalized snapshot work and per-operation cancellation; the candidate may temporarily coexist with the prepared plan, flushed temporary file and readback buffer. These operations are bounded per chunk and plan output, but this is not a hard process-memory cap.

Directory and file handles are closed on success and failure. After commit, a cleanup/close failure is a warning because the output already exists; before commit it is an explicit cleanup failure. `create` leaves a temporary entry if unlink fails after its published hard link succeeds, but reports that path in warnings. Callers remain responsible for application policy and cleanup of reported leftovers.

## Checked example

The example accepts a current source snapshot so replacement is explicit. It keeps the source session/plan alive through publication and disposes them after the write.

<!-- doc-test: docs/examples/file-save.ts -->
```ts
import {
  importPptxFile,
  writePptxFile,
  type PptxFileSnapshot,
} from "@pptkit/pptx-parser/node";

export async function preserveImportedFile(
  sourcePath: string,
  destinationPath: string,
  expectedDestination?: PptxFileSnapshot,
) {
  const imported = await importPptxFile(sourcePath, {
    projection: "content",
    editSaveProfile: "text-runs-v3",
  });
  try {
    const prepared = await imported.session.prepareSave({ mode: "no-op" });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    const result = await writePptxFile(
      imported.session,
      prepared.plan,
      destinationPath,
      expectedDestination
        ? { mode: "replace", expected: expectedDestination }
        : { mode: "create" },
    );
    return { result, source: imported.file };
  } finally {
    imported.session.dispose();
  }
}
```

## Verification and open gates

- [17 focused Parser groups](../../packages/pptx-parser/test/file-save.test.mjs) cover public create/replace, permissions, authentic/invalid snapshots, invocation-time options, source races, bounded reads, target content/inode/symlink changes, every injected precommit stage, short/zero/corrupt writes, final-boundary edits/disposal, cancellation timing, path/inode contention, independent-adapter create races, postdispatch warnings and cleanup failures.
- [Two integration groups](../../scripts/file-save-parity.test.mjs) perform three real mixed-deck replacement rounds, compare Node file candidates with byte-session output, inspect unaffected payloads independently and verify source/retry state. [Browser bundles](../../scripts/browser-bundles.test.mjs) confirm the root/browser entries omit Node exports and retain synchronous guard semantics; isolated [installed package](../../scripts/installed-packages.test.mjs)/[Node](../../tests/consumers/node.mjs)/[type](../../tests/consumers/types.ts) consumers exercise the packed `/node` export.
- The file-adapter suite additionally verifies defaulted metadata removals and reinsertions through three guarded in-place replacements in both dialects. Published bytes equal the prepared in-memory candidate, ordinary permissions persist, temporary files are cleaned up, and an externally changed destination is rejected without replacement.
- [Retained mixed-deck evidence](../testing/file-save-evidence.json) rechecks create and guarded replace of a real validated source: exact bytes, mode `0640`, directory sync, no unresolved temporary files and rejection of a receipt after a later mode change.
- [ADR 0011](../architecture/decisions/0011-node-file-publication.md) owns the transaction boundary, synchronous guard, durability and concurrency decisions.

This is implementation progress on **IMP-09.5.3**, not completion of its broad mode/platform matrix. Cross-process CAS, Windows/native filesystem certification, network/remote filesystem semantics, owner/ACL/xattr preservation and application-level repair behavior remain open. The inherited no-op/text-run candidate scope is unchanged, as are loss-bearing rebuild, structural reconciliation, SmartArt, media and animation.
