# ADR 0008 — Opt-in direct text-run preservation save

Status: accepted for experimental `text-runs-v1`; broader preservation remains planned.

Subsequent [ADR 0009](0009-text-owner-preservation-save.md) adds an explicit V2 shape-text/notes profile while retaining the V1 scope and report shape described here.

Date: 2026-10-25.

## Decision

Parser enables a bounded preservation edit only when content import explicitly selects `editSaveProfile: "text-runs-v1"`. Default edit blocking, no-op behavior and rebuild blocking remain compatible. The profile changes existing direct text-box run character values with fixed topology; it is not a semantic rebuild.

MCE privately retains effective-to-raw element origin maps. Source bindings carry original raw preorder, source/Core identities and baseline characters; no source XML/handles enter Core or IR. An independent snapshot analyzer classifies the whole document, skips derived plainText only on actual text elements, reconciles shared raw targets and rejects unsupported mixed plans. Raw guards include inactive shape sub-targets and signature/VBA relationships. Unknown relationship closure findings block.

The lexical character primitive from [ADR 0007](0007-xml-character-edit-boundaries.md) supplies bounded source-local modifications. A Parser-owned ZIP32 patch writer preserves untouched local records and central metadata, stores changed payloads and rebases offsets. It validates complete input/output and never removes source parts. XML source and container handling remain separate from Core snapshot analysis.

Candidate acceptance requires exact source/package comparison showing only approved XML payload differences, full mapped candidate reprojection equality and final plan freshness. No Exporter dependency is added. Save plans retain private candidate bytes and existing ownership/consent/disposal behavior. Edit preparation reserves its maximum output budget before work, shrinks when ready, and propagates caller cancellation/session disposal into child operations. One signal relay per caller signal prevents listener amplification.

Reports expose complete located changes and either a scoped-preservation loss-free plan or explicit blocked diagnostics. Exhaustion/inconsistency returns no partial plan. The source baseline does not advance after save. Disposal aborts child operations and detaches caller signal relays synchronously; operation finalizers can safely release them again. Nonempty loss authorization, structural reconciliation and file replacement belong to later tasks.

## Evidence and consequences

See the [public contract and linked tests](../../api/text-edit-save.md). Independent byte/ZIP/XML/DOM readers, three-round tests, real installed consumers and the named LibreOfficeDev/PDF experiment cover this profile. The producer resave drops source extension/orphan data retained by PPTKit, so it does not close native fidelity.

This is concrete progress on IMP-09.1, IMP-09.2.1 and IMP-09.5. Their complete inherited/native requirements remain open, as do arbitrary implicit references, text-range remapping, broader text owners, styles/inheritance writes, geometry/assets, structural operations and rebuild. Source-preserving bytes outside explicit atoms do not imply that arbitrary opaque-feature semantics have been interpreted.
