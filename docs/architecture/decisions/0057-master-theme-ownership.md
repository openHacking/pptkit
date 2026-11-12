# ADR0057 — Explicit master theme/background ownership

- Date: 2026-10-27
- Status: accepted for the authoring/output foundation; native and complete template gates remain open

## Problem

Core had one document theme and layouts, while source notes/master inspection already exposed distinct native domains. Adding notes fields or inherited artwork without a format-independent ownership model would duplicate theme/cascade logic and could leak values between templates.

## Decision

Add immutable initializer master declarations (`id`, `name`, optional partial theme/background) and optional layout `masterId`. Keep existing document/layout methods and the legacy default path. The new optional document registry also permits unchanged custom document implementations. Do not encode OOXML IDs, filenames or relations in Core.

Core materializes each master's theme from document defaults and binds layout themes once. Background origins distinguish slide/layout/master/default. Optional IR v2 additions are absent on old inputs; explicit none/zero remain values. Layout consumes resolved themes consistently during sizing and diagnostics; SVG and fresh PPTX use the same owner domain.

PPTX retains native ownership with explicit master registrations, independent theme parts, globally unique layout IDs and bidirectional layout/master relationships. The default document theme and independent notes theme keep their existing identities. An unused explicit master is preserved without manufacturing an authored layout. Parser source structure and exact no-op remain separate from a future Core master import.

Registry definitions are immutable at initialization. Adding a second mutation/override graph before AUT-07.1.2 would create hidden coupling, so parent edges, artwork definitions and unknown master keys reject. This makes unsupported declarations observable instead of silently dropping artwork. Master artwork, reusable theme identities, shared-definition editing and template rebinding remain the original follow-on tasks, not alternate parallel models.

## Verification and consequences

[API](../../api/multi-master-authoring.md), [tests/evidence](../../testing/multi-master-authoring.md) cover same placeholder keys/theme roles under distinct masters, interleaving, unused masters, copies, capture/immutability, malformed declarations, local none/zero, measurement ownership, SVG, native XML, public browser/installed consumers and exact legacy bytes.

The `backgroundSource` union gains `master`; exhaustive callers need the new case. All new collection/IR fields are optional for legacy structural compatibility. Resolved IR is detached state, not a live registry; consumers must re-normalize authoring changes. The native relationship model follows Microsoft's [slide master documentation](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-slide-masters). Independent graph tests do not replace actual native save/edit/reopen evidence.


## Native admission follow-up

[Actual PowerPoint 16.104 controls](../../testing/multi-master-native.md) found the initial allocator's master/layout ID overlap triggered repair. The corrected exporter uses a joint disjoint namespace, placing extra master IDs beyond every native layout ID while retaining legacy default IDs. The initial failed source remains. Independent source/2×2 native controls and three corrected no-edit save/close/reopen rounds verify the change; actual template/layout/theme editing remains open. Native filename/relationship renumbering is resolved by identity rather than silently ignored.

[Actual native operations](../../testing/multi-master-operations.md) now cover cross-master layout rebinding, placeholder append and selected-slide Office replacement over nine reopen rounds, including independent bystander/glyph checks and current scalar source-candidate boundaries. Complete Core shared-edit/override/artwork/template contracts remain open.

## Shared-definition editing follow-up

The [shared-edit slice](../../testing/shared-definition-edit.md) now supplies a `ManagedPresentationDocument` factory return type. The base structural contract stays unchanged. Add/replace master definitions and replace layouts by stable identity with full replacement semantics; missing declarations resume inheritance and none/zero/false remain explicit values. Input capture precedes staging; layout element collection IDs and their allocator commit only after candidate/affected-slide validation. An explicit validation data view avoids fake document instances or Proxy dispatch. Pure validation consumes data, not mutation capabilities. No new package dependency or source-format state enters Core.

This implements shared editing within the existing flat ownership graph. It does not add parent edges, change normalized IR shape, introduce a second cascade in downstream consumers, or close property-origin, artwork, source import/write or native admission requirements for these newly authored edits.

## Theme/background origin follow-up

`resolvePresentationInheritance` now exposes ordinary IR plus an immutable [theme/background origin report](../../api/inheritance-origins.md). Core's shared resolver selects effective values and records declaration owners together; ordinary normalization uses that same path. The report is separate from Canonical IR and Parser source provenance. Targets use logical kind/ID; paths are snapshot authoring locations. Built-in and synthesized opacity defaults are explicit. Downstream APIs and serialized output are unchanged. Element/preset/notes/artwork origins and general reset graphs remain follow-on work.


## Text-format and native reset follow-up

[Text-format origins](../../api/text-format-origins.md) now record actual placeholder/preset/local selections and optional-font null barriers within the existing normalization selectors. New text target identities preserve the original target interface. Source-only metadata stays outside ordinary IR.

Independent DOM and native opening found two serialization gaps: false run flags were omitted, and layout-owned placeholder-bound text was treated as a duplicate native prompt. Fresh output now writes explicit off flags and resolved placeholder paragraph/run defaults; layout artwork keeps ordinary identity. The [actual failure, source-only correction and three reopen rounds](../../testing/text-format-origins.md) are separate from the prior byte-equal baseline. Native run formatting is still materialized resolved state; complete live template inheritance is not claimed. Named marker font/outdent fixes also advance existing typography work; general advances/wrapping and source style writers remain open.


## Chart and direct-notes origin follow-up

The same selector/report path now traces chart common/title/legend/axis text and direct authored notes. Chart origins retain symbolic declaration locations while current owning themes resolve concrete values; title fontSize/bold built-ins are independent defaults. Notes have slide-scoped logical targets and their own content paths, without inheriting slide placeholders or introducing a native notes-master authoring model. The original target interface remains unchanged; report unions add notes and chart paths. [Verification](../../testing/chart-notes-origins.md) covers independent authored source reads, browser output/XML and installed consumers. Regional table provenance still requires integration at its merge boundary.


[Regional/direct table formatting origins](../../testing/table-format-origins.md) now extend the same report through actual raw merges and normalized cell selectors. Explicit data flow keeps metadata outside IR; geometry/content provenance and complete template/artwork/reset semantics remain follow-on requirements.


## Ordinary artwork ownership and visibility follow-up

[Master artwork](../../api/master-artwork.md) now uses the existing element collection and global identity store. Initial declarations register their elements; shared replacement stages collection IDs/allocator, semantic validation and affected layout/slide references before commit. Master asset references participate in lifecycle protection even when hidden or unused. No opaque format state or extra package enters Core.

Optional `elements` preserves the old runtime path when absent. `ManagedPresentationSlide` supplies a strict display setter through factory slide access while legacy custom slide requirements remain unchanged. Layout resolves the shared collection once into `LayoutSlideMaster`, including master-themed analysis/sizing/composition; Renderer draws master/layout/local layers once in order. Independent geometry dependencies remain valid when a layer is hidden.

Fresh master objects and resources stay in their native owning part. Serialization stages master charts before emitting complete chart/workbook resources while preserving legacy ZIP ordering. Layout and page display declarations map to separate native `showMasterSp` scopes. [Actual six-page, three-round evidence](../../testing/master-artwork.md) verifies named shapes/text/images/groups, all owner/media relationships, source-frame geometry, layer pixels and exact native reopen output; explicit true/default-size labeling changes are recorded. Advanced master object native appearance, native shared editing and full template/source-write gates remain open.
