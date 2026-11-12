# Cached SmartArt preview

The experimental content option `diagramPreview: "cache"` explicitly projects associated cached drawing shapes into ordinary Core groups for preview. The default projection continues to omit native diagrams. Structure mode and other option values reject. Selecting the option accepts that cache freshness is unknown; it does not request semantic SmartArt import or native editing.

<!-- doc-test: docs/examples/diagram-preview.ts -->
```ts
import { importPptx } from "@pptkit/pptx-parser";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

export async function previewCachedDiagrams(bytes: Uint8Array) {
  const session = await importPptx(bytes, { projection: "content", diagramPreview: "cache" });
  try {
    return { rendering: await renderPresentationToSvg(session.document), diagrams: session.report.diagramPreviews };
  } finally {
    session.dispose();
  }
}
```

The outer group keeps the graphic frame's local position, extent, rotation, flips, hidden state and accessibility description. Cached root and nested group transforms preserve their child origins/extents; without a root transform, cache coordinates use the source frame's point extent. EMUs convert to points using 12,700 EMUs per point. Cache IDs use raw source preorder identities inside a frame-specific prefix, not data-model node IDs.

The supported cache subset is DSP shape/group drawing with complete local geometry, the existing seven presets, direct fills/strokes, direct rich-text paragraphs/runs/frame settings, rectangular PNG/JPEG/GIF image fills using explicit stretch/crop, and the independent text rectangle profile below. Image assets share the normal import resolver and byte budgets, including shared slide/cache uses. Cache text uses local styles and Core defaults; full diagram theme/layout/font inheritance, custom geometry, native node semantics and source actions are not implemented. Source actions are neither executed nor attached to cache text. Unsupported attributes/subtrees produce located `smartart.cache.unmapped-property` preview findings. Unsupported shapes use visible placeholders; supported neighbors stay in order.

An explicit `dsp:txXfrm` with complete nonzero offset/extent creates a separate text sibling in the parent's coordinate space. Its rectangle is rebased by the parent child origin, and its local rotation combines shape, text-transform and text-body rotation modulo 360 degrees. Parent group transforms apply once. A compensating `-90` degree text angle therefore keeps text upright relative to its group when the shape is rotated `90` degrees. Unknown visual presets can still expose this known text rectangle, with the visual reported as partial. Image captions can also use an independent rectangle instead of the image frame.

Incomplete/zero text rectangles, flips on the shape or text transform, and `bodyPr/@upright=true` retain an unavailable text marker and located findings; they do not gain guessed placement. Unavailable text bodies are not interpreted or validated as mapped text. Malformed mapped transform numbers or duplicate `txXfrm` nodes fail with cache-part provenance. This profile's geometry does not establish native application fidelity or support for flipped/upright text, vertical flow, all fonts/paints or arbitrary cache transforms. See [ADR0019](../architecture/decisions/0019-cached-diagram-text-transform.md).

The report's optional readonly `diagramPreviews` records source frame, Core group ID, available cache part, `freshness: "unknown"`, and one of:

| Status | Result |
| --- | --- |
| `cache` | The bounded drawing reader consumed the preview subset without property findings. This still means a cached approximation, with defaults and unknown freshness. |
| `partial-cache` | Cached drawing fields plus placeholders/defaults or reported unmapped properties. |
| `placeholder` | No usable associated drawing: missing/unresolved cache, wrong type/root, empty/missing tree, cache MCE, or incomplete root coordinates. The reason remains in the report. |

Missing source frame geometry still follows the existing content-import rule: the object is omitted with capability findings. A mapped nonzero frame is needed for cache coordinates. A cache with unknown freshness can be previewed only through this explicit option. The optional [direct text comparison](diagram-text-comparison.md) can report exact source/cache paragraph disagreements; it does not determine general cache freshness or change this preview policy. Malformed scalar fields in the mapped subset fail import with cache-part provenance; malformed package/XML continues to fail in the existing inspection layer. Budget/cancellation failures are not converted to successful placeholders.

All `smartart.cache-preview` create/edit/native-export/edit-save capabilities remain unsupported; import/preview are fallback, and no-op is preserved. The derived ordinary shapes are mutable Core authoring values, but their changes receive no SmartArt writeback permission. No-op retains original bytes. With a text-run save profile, unrelated supported text changes preserve cache parts, and candidate validation reprojects using the same captured preview option. Editing cached text, geometry or children is blocked by the ordinary unbound-change guard. Direct fresh export emits ordinary shapes/groups; it does not produce native SmartArt or retain the source diagram. Original source identity and provenance remain in Parser reports.

`sessionLimits.maxDiagramPreviewWork` defaults to 1,000,000 traversal checkpoints, charged across all cache uses, including shared-source repetition. Existing image-byte, XML/OPC, finding, snapshot and save-validation budgets also apply. The byte and Node file APIs capture the option before their first await. Root/browser/Node public consumers need no new package or renderer dependency.

Evidence: [focused tests](../../packages/pptx-parser/test/diagram-preview.test.mjs), [browser geometry and three-round save integration](../../scripts/diagram-preview-parity.test.mjs), installed consumers and [ADR0018](../architecture/decisions/0018-cached-diagram-preview.md). Browser tests independently check visible rectangle/image bounds and block all network access. These checks do not establish native SmartArt fidelity or complete ADV-01.1.2 acceptance.

[Independent text source evidence](../testing/smartart-text-transform-evidence.json) additionally pins three nonempty-cache source files and verifies seven text rectangles, strings and combined angles through Python ZIP/ElementTree. All three previews remain partial because other presets/styles are unsupported. No native application was launched. Reproduce with the source runner's `--preview --output <evidence-file>` options and the retained corpus/provenance directory; third-party fixture redistribution admission remains pending.
