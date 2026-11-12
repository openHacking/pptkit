# ADR 0007 — Preserve lexical XML around bounded character edits

Status: accepted for the private character-edit primitive. Subsequent opt-in session wiring is defined in [ADR 0008](0008-direct-text-preservation-save.md); this primitive itself grants no save permission.

Date: 2026-10-25.

## Problem

Rebuilding a complete slide from its current Core projection would drop source content that the projection does not understand. Serializing the effective MCE tree would also discard inactive branches. A global replacement cannot distinguish repeated text, similarly named elements in different namespaces, extension payloads or source text mentioned in a comment.

The first text-save task therefore needs a smaller proven transformation: change explicitly identified character content and retain every other lexical region of the original XML part. This primitive alone cannot prove the selected source content belongs to a particular Core run or that changing it preserves application-level references.

## Decision

Parser owns a private [XML text source](../../../packages/pptx-parser/src/preservation/xml-text-source.ts). It has no public entry-point export and adds no dependency, Core/IR field, package writer or Exporter coupling.

Capture first snapshots input/options and validates the complete document with the existing bounded XML reader. A lexical scanner then binds source spans to the validated tree's full preorder, retaining all branches rather than an effective projection. The scanner is not an independent XML parser and never accepts an unvalidated document.

Every requested edit provides the raw element index, expanded namespace/local name, exact expected baseline text and replacement text. Stale text, missing/mismatched identity or repeated target rejects the entire candidate. Only character-only content is admitted: text, CDATA and empty elements. Child elements, comments and processing instructions inside a target prevent replacement. The future format-level mapper must separately prove that an admitted character-only element is an allowed editable target.

Replacement escapes XML delimiters and represents carriage return as a character reference, so it survives XML 1.0 line-ending normalization. Invalid XML characters and lone surrogates reject. A changed CDATA/entity representation can become escaped ordinary text inside the selected region. Equal-text requests leave the original representation untouched. Self-closing elements expand only at their `/>` suffix, preserving their prefix and attribute spelling.

Unselected decoded regions are copied verbatim and encoded with the source UTF-8 or BOM-marked UTF-16 byte order. BOM, declaration, namespace bindings, QName values, whitespace, comments, processing instructions, extension content and inactive MCE markup outside selected ranges remain byte-exact. Candidate XML is reparsed under the original XML budgets; element identities and requested text are checked before output is returned.

Edits are canonicalized to source order and each candidate starts from the captured original. There is no implicit baseline advance or mutable shared source graph. Private source text can be disposed; in-flight work checks disposal/cancellation and returns no partial output. Returned bytes are detached. Frozen metadata and changed source ranges contain no mutable private parser trees or source buffer.

The independent edit defaults are 10,000 requests, 2 Mi UTF-16 replacement units, 8 MiB encoded output and 32 Mi charged work units. Exact encoded output size, including escaping and BOM, is counted before candidate allocation. Work charges requests and candidate-segment code units; original XML budgets separately bound capture and output revalidation. These limits are not a peak-memory measurement or an OPC package budget.

## Evidence

[Private primitive tests](../../../packages/pptx-parser/test/xml-text-source.test.mjs) compare explicit lexical output, independent Buffer UTF-8/UTF-16 encodings and 64 generated template cases. Tests cover repeated text, namespace shadowing, unknown siblings, all MCE branches, quoted `>` attributes, comments/PI/CDATA, empty/self-closing targets, Unicode, injection attempts, stale identities, immutable input/options capture, exact budgets, cancellation, disposal and failure retry.

Three patch/repack/reimport rounds in both PPTX dialects compare every untouched part using a separate ZIP reader and SHA256. They retain unknown effects/extensions, source object IDs and all relationship/incoming/unreached inventories. Existing sessions still reject edit mode. The test fixture repacker is not a production preservation writer.

A test-only private bundle in the [Chromium suite](../../../scripts/browser-bundles.test.mjs) verifies all three source encodings against independent Buffer byte expectations and browser DOM parsing. DOM text, expanded names and opaque attribute/content values agree; output-budget and disposal failures remain controlled and no network requests occur. This does not add the primitive to published browser exports.

## Remaining gates

IMP-07.2.1 remains open. Next work must bind source targets to imported Core runs with stable provenance, classify changes and prove operation-level ownership. Edits involving text fields, mixed child content, target ranges used by animations, shared source parts and unknown application references need explicit decisions and counterexamples. Formatting/inheritance changes, arbitrary XML subtree mutation, identity remapping and relationship rewrites are outside this primitive.

IMP-09 still requires a bounded production package writer, candidate validation/comparison, snapshot-bound plans, exact loss handling, concurrent cancellation/failure recovery and native application reopening. Signed-package admission and signature invalidation also need package-level decisions; a single XML-part transform cannot make that decision. No public edit-save, native fidelity or completed preservation task is claimed by this ADR.
