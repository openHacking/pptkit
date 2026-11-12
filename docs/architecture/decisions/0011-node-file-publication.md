# ADR 0011 — Node file publication and synchronous plan guard

Status: accepted for experimental `@pptkit/pptx-parser/node`.

Date: 2026-10-25.

## Context

The preservation session deliberately returns caller-owned bytes and owns no filesystem policy. Applications need an explicit local-file adapter that can import a stable source, create a new output without clobbering an existing name and replace only the file the caller inspected. A file writer also cannot rely on an asynchronous `save()` call for its final freshness check: JavaScript could run caller mutation or lifecycle work after that promise resolves but before the operating-system publication call is dispatched.

Standard portable Node filesystem operations offer exclusive hard-link creation and atomic same-directory rename, but no compare-and-swap rename against a previously observed inode/SHA256. Durability also differs: a synced file survives a file write, while directory entry durability requires a directory sync. Postcommit failures cannot be represented as a precommit rejection without denying a file that actually exists.

## Decision

Add a Node-only public entry, `@pptkit/pptx-parser/node`, exposing `importPptxFile`, `writePptxFile`, `PptxFileError` and their types. The root and browser entries expose byte import/save only. Node filesystem code remains in `src/node` and imports only through package internals; the browser bundle continues to build from the root entry and rejects Node/governance leakage.

Add the runtime-neutral synchronous `session.assertSavePlanCurrent(plan, { acceptLossIds? })`. It validates session/plan ownership, plan/session liveness, exact consent and the complete normalized candidate snapshot. It neither copies bytes nor consumes the plan. The private `currentPlan` implementation is shared by this guard and `save()` so the contracts cannot drift. The Node adapter calls the guard after all asynchronous temporary-file work and immediately before invoking `link` or `rename`.

`importPptxFile` captures nested options and the signal before the first filesystem await. It canonicalizes the parent, rejects a non-regular/symlink source and takes a stable, bounded snapshot. Its receipt is a frozen nominal interface whose native identity remains in a private adapter WeakMap. It records the canonical path, byte length and SHA256; JSON serialization is rejected.

`writePptxFile` requires an explicit union:

```ts
{ mode: "create" }
| { mode: "replace"; expected: PptxFileSnapshot }
```

A replace caller must present a receipt issued by the adapter and to the same canonical path. Create mode cannot carry a stale/foreign snapshot as overwrite authority.

The adapter writes a random temporary regular file in the target's canonical parent directory. Required no-follow/nonblocking flags are checked without a weaker fallback. It writes and reads at most 64 KiB per chunk, forces the temp to `0600` before writing and syncs it, closes it, then reopens/readbacks it with stable identity. It verifies a SHA256 match and the named temp identity. Replace preserves only ordinary `0777` permission bits from the expected inode.

For create, a hard link from the validated temporary inode to the destination is the exclusive publication operation. It has no truncate/check-then-create window. Temp unlink follows.

For replace, the adapter checks expected canonical path, regular-file identity (device/inode/size/mode/mtime), parent identity and complete SHA256. It performs that check before and after temp construction, then rechecks all state with the synchronous plan guard. A same-directory `rename` publishes the validated inode atomically.

Both file and directory are synced. After publication, cancellation is too late to roll back. Directory sync, temporary unlink or directory close failures yield `status: "committed"` with immutable warning codes and, where relevant, `directorySynced: false`. Before publication, temporary/handle cleanup failure rejects with `file-cleanup-failed`, `committed: false` and a `temporaryPath` if one remains. The original file is never described as replaced in that case.

The public entry shares one adapter instance, which fails fast on same-process canonical-path or expected-inode contention. Independent instances/processes still race only through their filesystem primitives.

## Consequences

The writer permits no hidden source rebasing, accepts no implicit "overwrite" or filename-derived policy, and retains prepared plans for exact retry. A successful file receipt can be used as the next expected snapshot while the existing bytes still match. Sessions, plans and receipts stay process-local; transfer requires re-import.

Replacement is optimistic across processes. Stat/hash verification can detect all changes before its final check, but an external process may write after that check and before rename. Portable Node offers no atomic expected-content rename. The adapter therefore assumes a trusted, stable parent and cooperating writers; it does not promise universal protection against hostile local mutation, another process's final-interval write, malicious directory replacement, remote-filesystem identity quirks or editors that do not use this protocol. A stronger guarantee requires an external lock/CAS mechanism and its own tests.

Temporary file and target names share a parent, so publication does not cross filesystems. The adapter preserves content, ordinary permission bits and the selected candidate report; owner, ACLs, extended attributes, alternate streams and creation time are outside the contract. Windows and nonportable filesystem operations remain unverified and reject if the required platform flags are unavailable.

This decision keeps Core, Layout, Exporter and the Parser root independent of filesystem policy. It advances IMP-09.5.3 but does not close its complete platform/mode matrix, cross-process transaction work or inherited rebuild/SmartArt/media/animation tasks.
