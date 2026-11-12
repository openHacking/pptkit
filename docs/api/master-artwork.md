# Master artwork ownership and visibility

Slide masters can now own ordinary `elements`, using the existing text, shape, image, group, connector, table and chart inputs. They retain a shared collection in Core, IR and Layout. Fresh PPTX writes that collection into its own slide-master part and relationships. Slide object arrays do not receive copies of master or layout artwork.

See the [typed example](../examples/master-artwork.ts), [shared-definition editing](multi-master-authoring.md#editing-shared-definitions), [multi-master themes](multi-master-authoring.md) and [verification](../testing/master-artwork.md).

## Create and replace

Use `createPresentation({ masters: [...] })` for initializer declarations, or `defineSlideMaster` after registering resources/presets. `elements` is optional. Old declarations keep their previous shape; an explicit empty array remains an explicit empty collection. Initializer objects are captured/frozen and register nested IDs. A missing resource in an initializer can be registered before validation/normalization, just as other draft authoring references can be completed later.

`replaceSlideMaster(id, input)` fully replaces the definition. Omitted elements remove the old artwork; omitted theme/background declarations resume their existing fallbacks. New definition edits require finite plain data, with no accessors. IDs, nested groups, validation and the allocator stage together through the existing collection transaction. Failed replacement keeps the previous definition and ID ownership. Dependent layout/slide connector references, including a slide using the authored default layout through an omitted `layoutId`, prevent removal of their target.

Element IDs remain **document-global**, across master/layout/slide/group owners. Equal IDs in two Core collections reject; they do not select an override by coincidence. Native drawing IDs are a different, part-local namespace. Master elements have no placeholder-definition registry: a `placeholderKey` on a master object rejects. Layout placeholder definitions continue to supply geometry/style to bound slide content, and ordinary layout artwork remains ordinary artwork rather than a native prompt.

No parent-master graph or per-object override graph is introduced. Shared order is the order of each owner's collection. Replace a definition to change its shared order or artwork.

## Display switches

| Declaration | Omitted / true | False |
| --- | --- | --- |
| `SlideLayoutInput.showMasterShapes` | Show the owning master's ordinary artwork | Hide master artwork beneath this layout; keep layout artwork |
| `PresentationSlideInput.showInheritedShapes` | Show the layout's allowed inherited layers | Hide ordinary layout and master artwork on this page |

These switches affect artwork, independently of background paint and local slide content. Explicit true on a slide does not undo a layout's `showMasterShapes: false`. Existing bound slide placeholder content stays local and visible. This does not define visibility for unmodeled imported header/footer fields or master placeholders.

Factory `slides`, `addSlide`, `insertSlide` and `duplicateSlide` expose `ManagedPresentationSlide.setShowInheritedShapes(value?)`. Pass `undefined` to restore the default. Nonboolean values reject without changing the flag. Copies retain false/true/absence. The base `PresentationSlide` gains only an optional readonly declaration; legacy custom slides do not need a new mutation method.

Fresh XML maps these declarations to `p:sldLayout@showMasterSp` and `p:sld@showMasterSp`. Microsoft's [layout schema explanation](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-slide-layouts) documents the master-shape attribute, and its [DisplayMasterShapes reference](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.displaymastershapes) distinguishes background objects from background paint. The named PowerPoint controls independently verify the table above, including the slide-level suppression of ordinary layout artwork.

## Resolution and resources

Drawing order is background, master artwork, layout artwork, then page elements. Each visible collection is drawn once per rendered page. Groups retain their internal order/transforms. A local background override changes the paint beneath artwork; it does not hide inherited objects. Each owner uses its resolved theme, with no cross-master fallback.

`LayoutResult.masters` uses exported `LayoutSlideMaster`, whose optional elements are resolved `LayoutElement`s. Images, groups, charts, table borders, analysis, intrinsic/table sizing and master composition containers use the owning master domain. Use `resolveLayout` / `resolveNormalizedLayout` for transported data containing artwork rather than casting normalized connector/chart elements into resolved geometry. Existing objects without artwork keep the legacy runtime/serialized path.

Layout/slide connectors can use their owning master's geometry. Hiding artwork does not remove that geometry from the authoring graph: a connector may still point to a hidden inherited shape's position. Cross-owner native connector attachment retains the existing fallback rules; it does not become a native cross-part connection. Master connectors stay within their own collection/group scope. Master tables do not gain slide-table pagination support.

`getAssetReferences` adds optional `masterId` records with actual master/group paths. Hidden and unused master collections still retain their assets. Asset removal and cleanup include those references, and replacing one definition does not delete media still used by another master/layout/page. One asset ID is packaged once, with relationships from its actual owners; different asset IDs retain the existing deduplication policy. Load failures remain explicit, and master-owned omissions/materialization warnings carry `masterId`.

## Verification boundary

The named shape/text/RGB-PNG/group deck has actual first-open and three save/close/exact-path-reopen controls. Whole graph, media bytes, source-frame geometry, visible layers, text color/font selection and page correspondence are checked independently. Explicit true can become an omitted native default after saving; it is not a retained lexical override claim.

Master table/chart/workbook emission and shared sizing/composition have automated integration coverage. Their full native appearance/editing, arbitrary template resets, source artwork/visibility projection, source style/structure writes and independent authored notes-master/theme models remain separate requirements. Parser's exact no-op checks preserve source packages and do not grant those editing capabilities. Full parent tasks remain open.

## Actual shared-heading control

A [named native/shared replacement control](../testing/master-shared-edit.md) now covers a master ordinary-text edit across two visible pages, an inherited-artwork-hidden page and a second master. Fresh `replaceSlideMaster` output and actual manual native master editing produce identical four-page native prints; three later native edit/reopen rounds preserve bystanders. The same fixture verifies master-owned grouped merged tables, positive bar/line/pie charts and a straight connector at named boundaries. This does not grant source master writes or establish general table/chart/image editing and complete reset/rebinding support.
