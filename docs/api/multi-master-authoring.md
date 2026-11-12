# Multiple authored slide masters

`createPresentation({ masters })` captures a readonly registry of format-independent master declarations. A layout's optional `masterId` selects its shared theme/background/artwork owner. Layout IDs remain document-global; placeholder keys are local to each layout, so the same key can occur under different masters without crossing domains.

```ts
const document = createPresentation({
  masters: [{ id: "brand", name: "Brand", theme: {
    colors: { accent1: "CC2200" }, fonts: { body: "Georgia" },
  }, background: { type: "solid", color: { theme: "accent1" } } }],
});
document.defineSlideLayout({ id: "title", name: "Title", masterId: "brand" });
document.addSlide({ layoutId: "title" });
```

See the [checked example](../examples/multi-master-authoring.ts), [ADR0057](../architecture/decisions/0057-master-theme-ownership.md), and [verification](../testing/multi-master-authoring.md).

## Ownership and defaults

Each master has a nonempty stable `id` and display `name`, with optional partial theme, background and [ordinary artwork elements](master-artwork.md). Inputs and nested values are cloned/frozen at creation. `document.masters` is a getter with a frozen registry; there is no new method requirement for existing custom `PresentationDocument` implementations. The interface's optional property treats an absent registry as empty.

A master theme inherits omitted name, color roles and heading/body font roles from the document theme. Master objects do not inherit from one another and cannot create cyclic graphs. Unknown master declaration keys, including parent graphs and alternative artwork containers, reject instead of silently inventing semantics. Invalid/duplicate IDs, names, theme components/fonts/colors and missing layout master references reject before normalization succeeds. IDs are logical strings, never output filenames.

A layout without `masterId` retains the legacy default master and document theme. Old normalized IR omits all new fields; supplying `masters: []` generates the same PPTX bytes as omitting the registry. Existing layout methods, element identities, placeholder keys, copying and notes behavior remain unchanged. Copying a slide retains its layout/master sharing.

Background precedence is explicit slide paint, explicit layout paint, declared master paint, then the existing default background. `none` and opacity zero are explicit values, not inheritance. Normalized slides may now report `backgroundSource: "master"`; exhaustive consumers of that union must handle the new case. Bound normalized layouts record the effective theme and background source; normalized master definitions retain native ownership. The data is detached resolved IR, not a live editable registry. Re-normalize authoring input after changes; mutating a resolved theme does not recompute other resolved values.

## Downstream behavior

Core resolves owner themes once. Layout transports them and uses the matching theme for intrinsic text sizing, table row sizing and overflow diagnostics, including reusable layout objects. SVG uses the same theme for slide/layout objects and inherited background. Charts receive the matching normalization theme. Neither Layout nor SVG solves a separate master cascade.

Fresh PPTX produces a distinct slide-master/theme part for each explicit declaration, presentation registrations, master-to-layout lists and layout-to-master relationships. Native master and layout identities use disjoint globally unique ranges even with interleaved master ownership; this prevents the PowerPoint repair behavior found during native admission. An inherited background remains on its master rather than being copied into each native layout/slide. Unused explicit masters are emitted with `preserve="1"` and no empty layout list. The legacy default master/theme remains; notes retain their separate existing global-theme domain.

Manually transported IR with a master binding but missing resolved master/theme fails instead of borrowing the document theme. Public normalized/Layout data remains caller-supplied resolved state; arbitrary manual mutations are not an authoring workflow.

## Editing shared definitions

Factory-created documents return `ManagedPresentationDocument`, extending the unchanged `PresentationDocument` contract with `defineSlideMaster`, `replaceSlideMaster` and `replaceSlideLayout`. Legacy custom implementations need no new methods. A variable explicitly typed as the base interface exposes only its existing capabilities.

```ts
document.defineSlideMaster({ id: "alternate", name: "Alternate" });
document.replaceSlideMaster("brand", { name: "Updated brand",
  theme: { colors: { accent1: "118855" } },
});
document.replaceSlideLayout("title", {
  name: "Title", masterId: "brand", background: { type: "none" },
});
```

Replacements retain the target ID and registry position. They replace the whole definition rather than merge a patch. Supply every element/placeholder you want to retain. Omitted theme roles resume document inheritance; omitted background resumes its owner; omitted `masterId` returns to the document's default master. Explicit `none`, opacity zero and false style flags remain declarations. `null` is not a reset command for these master/theme/background declarations; existing element script-font reset fields retain their own contract. Replacing a shared master affects all its layouts; replacing a layout affects every bound slide on the next normalization/preview/export. Previously returned immutable snapshots remain stable; resolved IR snapshots must be regenerated.

New edits capture finite plain data before staging changes. Accessors, cycles, symbol/non-enumerable fields, unsupported prototypes, sparse/custom arrays and unknown top-level definition fields reject. Data is limited to one million visited values and depth 128, including repeated aliases. Master parent graphs remain unsupported; the existing `elements` field owns ordinary artwork.

Master and layout edits validate the replacement declaration and stage their entire element collections, permitting existing IDs to move between its roots/groups while rejecting collisions with other owners. Definition validation and validation of every bound slide (including placeholder/connector dependencies) must pass before registration changes. Rejection keeps old definitions, ID ownership, asset references and automatic-ID allocation unchanged. Existing errors on affected slides must be repaired first; unrelated semantic draft errors do not prevent the edit. Malformed runtime data outside the typed authoring contract may still abort validation. Failures throw `PresentationValidationError` for semantic diagnostics, and ordinary `TypeError`, `RangeError` or `Error` for malformed input, limits, identity and duplicate/unknown-target failures. Removing definitions, document-theme mutation and automatic slide rebinding remain separate. Master artwork now follows the same owned collection transaction; omission removes it.

See [shared-edit verification](../testing/shared-definition-edit.md) and the [theme/background property-origin API](inheritance-origins.md). This is an AUT-07.1.2 implementation slice; current text/preset/chart/direct-notes/table formatting origins are available through the shared resolver; broader geometry/content/reset provenance remains pending.

## Admission boundary

This is the master/theme/background authoring foundation for AUT-07.1.1 with the shared-edit slice above, with a later [ordinary artwork/visibility slice](master-artwork.md), not complete template editing. Master-level placeholders, advanced native artwork admission, complete clear/reset provenance, template rebinding, independent theme registries, complete source-to-Core master import and source master/style writes remain open. Parser structure/no-op can inspect/preserve the native graph; a content projection/fresh rebuild does not reconstruct this authored registry. [Native no-edit admission](../testing/multi-master-native.md) now covers the real ID-collision failure, corrected fresh source and three save/close/reopen rounds. [Named actual layout/text/selected-theme operations](../testing/multi-master-operations.md) now have nine native reopen rounds; full shared editing, artwork and generalized template operations remain pending. AUT-07.1.1/.2/.3/.4 and complete notes/field parents are not closed by automated output checks.
