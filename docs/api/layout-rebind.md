# Reviewed layout rebinding

`prepareLayoutRebind` prepares an immutable authoring value plan; factory-created `ManagedPresentationDocument.applyLayoutRebind` applies all selected pages atomically. Legacy `PresentationDocument`/`PresentationSlide` implementations need no new methods. Normalization, inheritance resolution and validation now consume read-only authoring `PresentationData`/`PresentationSlideData`; managed documents remain compatible. These data views contain values, not edit methods or Parser package fragments.

```ts
import { createPresentation, prepareLayoutRebind } from "@pptkit/core";
const document = createPresentation();
document.defineSlideLayout({ id: "a", name: "A", placeholders: [
  { key: "body", kind: "body", box: { x: 20, y: 20, width: 400, height: 200 } },
] });
document.defineSlideLayout({ id: "b", name: "B", placeholders: [
  { key: "copy", kind: "body", box: { x: 40, y: 40, width: 350, height: 180 } },
] });
const slide = document.addSlide({ layoutId: "a", elements: [
  { type: "text", placeholderKey: "body", content: "Retained content" },
] });
const plan = prepareLayoutRebind(document, [{ slideId: slide.id, layoutId: "b" }], {
  localAppearance: "adopt", reset: ["geometry", "text"],
});
if (plan.status === "ready") document.applyLayoutRebind(plan);
// Review matches/findings before applying; discard the plan to cancel.
```

See [typed example](../examples/layout-rebind.ts) and [decision](../architecture/decisions/0058-reviewed-layout-rebinding.md).

## Inputs and matching

Each request has `slideId`, target logical `layoutId`, optional `bindings` (top-level element ID → target placeholder key or null detach) and optional `references` (old inherited object ID → target inherited object ID). Requests have unique existing page IDs and are evaluated in document order. Unknown IDs/options, malformed/sparse/cyclic data and accessors reject before mutation; accessors are not executed.

Matching priority after explicit caller decisions is same logical key with same role, same declared `index` with same role, then one unique same-role candidate. Candidates are reported by key. Multiple candidates and same-key/index role conflicts never choose by array order. `unmatched` defaults to `block`; `detach` permits a missing role to retain local content/old geometry. Ambiguity still blocks unless explicitly mapped or detached. Explicit role changes are reported. Ordinary objects are retained and not automatically bound by names or positions; an explicit mapping may bind them.

The implicit default is logical `layout-blank`. A no-op retains an originally omitted `layoutId`; rebinding to an unauthored default clears the explicit ID. A placeholder may declare `index` from 0 through 4294967295, unique within its layout. Matches expose `sourceIndex`/`targetIndex` only when declared and use `method: index` for numeric matching. Key/explicit decisions may choose another index with an informational finding; missing numeric matches may still use unique-role fallback. Duplicate/invalid indices fail validation. Exporter-generated ordinals remain outside Core identity. See [index policy and native verification limits](../testing/placeholder-index.md); complete native template import and cross-document/dimension adaptation remain open.

## Appearance and reset

`localAppearance` is required:

- `preserve` fixes effective foreground geometry, text/frame styles, colors/fonts/scripts, shape/image/connector/table/chart values and background to the old owning theme. IDs, content/actions, image resources and existing local references survive. Regional tables become explicit effective cells; their live region rules are not retained and the plan reports materialization.
- `adopt` retains local declarations and symbolic theme references under the target template. `reset` may contain unique `geometry`, `text`, `background` entries. Geometry resumes the matched placeholder box; text clears matched text/shape frame, preset and paragraph/run style declarations while preserving text/actions; background clears the local paint. Ordinary unbound object formatting and table/chart rules are retained. Reset is rejected in preserve mode.

Target master/layout artwork always follows the target owner. Preserving foreground appearance does not flatten shared artwork or promise identical complete pages. Host font availability, measured shaping/intrinsic sizing/composition, source fields and cross-size adaptations are not inferred. Core issues deterministic overflow/font-family/shrink-policy findings; it never silently reduces font size to fit. An existing authored target shrink policy remains explicit and reported.

Missing inherited connector IDs require explicit reference maps. Slide-level targets must be inherited roots, not group descendants in a different coordinate scope. Preserve mode blocks changed inherited targets because Core does not own Layout anchor/transform solving. Adopt mode follows an explicit valid target. Unsupported cases return review findings rather than silently producing detached or guessed geometry.

## Review, freshness and atomicity

Plans expose `status`, captured `requests`/`options`, `matches`, `findings`, ready-plan effective property `changes`, `changedSlideIds` and `sourceSnapshot`. The last field is the exact serialized finite authoring value snapshot, not a hash/security token. Plans and nested data are frozen. `needs-review` has no guessed after-value preview and application leaves every page unchanged.

Application recaptures current authoring data and recomputes the plan from its inputs. Forging `status`, findings or candidate reports does not bypass checks. Any source change, even unrelated notes or an invalid new draft value, throws `LayoutRebindError` with `code: stale-plan`. Malformed plan input descriptors throw `invalid-plan`; malformed captured requests/options throw TypeError. Plans may be serialized and applied to equivalent authoring values; there is no hidden registry or object-identity ticket.

One element-store transaction stages every changed page, validates full candidate data and commits IDs/allocator before updating existing managed slide handles. No page is committed early. Unselected pages/shared templates/assets and notes/timing/transitions remain stable. Local object IDs are retained; repeated ready no-ops preserve object pointers and consume no IDs. Calling no apply is cancellation. Returning to a prior template follows the selected policies; clearing overrides is not an undo history that resurrects discarded declarations.

## Output and source boundary

Layout/Renderer/fresh PPTX consume ordinary re-normalized values; there is no new IR version, native state or output default. Public consumer tests independently read actual target layout/master, text/image placeholder indices, hidden/display flags, geometry, effective theme fonts, table fills and chart colors. Parser sessions retain their source-only no-op and existing writer profiles. A new target definition/rebinding remains unsupported for source preservation writes; the authoring transaction does not grant that permission.

[Verification](../testing/layout-rebind.md) states automated and actual native scope. Full template/field/source gates remain open.

### Selective text style reset

In `adopt` mode, `reset` may include individual run selectors such as `text.fontFamily`, `text.eastAsianFontFamily`, `text.complexScriptFontFamily`, `text.scriptFontFamilies`, `text.fontSize`, `text.bold`, `text.italic`, `text.underline`, `text.strike`, `text.color`, `text.colorOpacity`, and `text.language`. Paragraph selectors are `text.paragraph.align`, `.indent`, `.hanging`, `.lineSpacing`, `.spaceBefore`, `.spaceAfter`, and `.bullet` (each with the `text.paragraph` prefix). Frame selectors are `text.frame.margin`, `.verticalAlign`, `.wrap`, and `.autoFit` (each with the `text.frame` prefix).

For matched text and shape-text placeholders, each selector removes the corresponding local declaration and resumes target placeholder inheritance, including theme references. Other declarations, actions and text remain unchanged. Ordinary and detached elements are not reset. Margin sides can also be reset individually with `text.frame.margin.top`, `.right`, `.bottom`, or `.left` (each with the full prefix). A numeric margin is expanded so the other sides retain their values. Script maps can also be reset by supported ISO 15924 key, for example `text.scriptFontFamilies.Hans`. A whole-map `null` clear is expanded so other scripts remain cleared. `bullet` and `autoFit` remain whole-property operations. The existing `text` reset clears all local text styling and takes precedence over individual selectors.

A local `textStylePreset` remains referenced. `textStylePresetReset` records bypassed preset properties using the same names without the `text.` prefix, so other preset values remain live. It accepts a unique array of supported properties and applies to text or shape text. A later local declaration still overrides inheritance. Table cells, fields, and general transformed geometry need separate contracts.

Fresh PPTX export defaults to materializing effective values and emits `native-text-inheritance-materialized` for matched shapes carrying reset selectors. Call `generatePptx(document, { nativePlaceholderText: "inherit" })` to opt into native placeholder inheritance for eligible `text.fontFamily`, `text.color`, `text.bold`, `text.italic`, `text.underline`, `text.strike`, `text.fontSize`, and `text.language` run declarations. A later local run declaration always remains local. Script font maps, font-slot null barriers, and explicit or combined color opacity can prevent this mapping; affected properties remain explicit with a warning. Eligible paragraph alignment, indent, hanging indent, line spacing, space before/after and bullet can also inherit independently; list-generated indentation that cannot be proven to inherit remains explicit with a warning. Eligible frame wrap, vertical alignment, margins and the whole autofit property can also inherit; later local frame declarations remain explicit. Unsupported reset properties remain explicit with a warning. A named combined frame control has three native save/reopen rounds. Six independent property controls have three-round PowerPoint evidence, including zero insets and wrapped long text; other hosts remain unverified.

The opt-in writes only selected eligible level-one layout defaults; a font default can change text metrics for other objects sharing that layout; `native-placeholder-metrics-may-change` reports the affected layout. Review the resulting presentation. Native inheritance of other properties and editing layout formatting in the PowerPoint UI are not yet verified.
