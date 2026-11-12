# Document Models

PPTKit uses three document layers with different audiences and stability needs. Keeping them distinct prevents a convenience API, a shared semantic contract, and a file-format package graph from collapsing into one model.

## Authoring Model

The Authoring Model is the application-facing representation in `@pptkit/core`. It is ergonomic, method-managed, and allowed to accept conveniences such as omitted IDs, string text, partial themes, and placeholder-derived boxes.

It owns construction and editing intent, not package storage. See [Core Authoring Model](core-authoring-model.md) for state, identity, validation, and normalization rules.

## Canonical Presentation IR

Canonical IR is the validated, detached, fully explicit document consumed across package boundaries. It resolves authoring ambiguity while remaining independent from PPTX.

It owns normalized slides/elements, stable references, consistent geometry/style representation, and a shared contract for Layout and exporters. See [Canonical Presentation IR v2](canonical-ir-v2.md) for exact invariants.

## PPTX Package Model

The PPTX Package Model exists inside format-specific parser/exporter boundaries. It may represent parts, relationships, content types, slide/master/layout XML, notes, themes, media, extensions, and preservation fragments.

It must not become the public authoring API. OOXML contains storage and compatibility concepts that are neither ergonomic nor portable presentation semantics.

## Implemented forward flow

```text
Authoring Model
      │ Core validation + normalization
      ▼
Canonical Presentation IR v2
      │ Layout resolution
      ▼
Layout Result
      │ PPTX serialization
      ▼
PPTX Package Model
      │ ZIP encoding
      ▼
editable .pptx
```

## Planned reverse flow

A future parser will inspect a PPTX Package Model and classify structures by capability. Supported semantics can become editable normalized structures; unknown or partially understood structures may require preservation or fallback representations. Silent data loss is not an acceptable default.

The reverse flow is not implemented and is not documented as universal round-trip support.


Optional master artwork remains one owned collection in Authoring/Canonical IR and one resolved collection in Layout. Ordinary rendering selects visible master/layout layers without appending copies to slide elements; fresh PPTX writes original native owners. `LayoutSlideMaster` carries resolved connector/chart/table/image geometry. Legacy inputs omit the new fields. Visibility declarations do not alter background selection or remove geometric/reference ownership; see [the contract](../api/master-artwork.md).


Read-only `PresentationData` and `PresentationSlideData` describe authoring values without methods; they do not add a fourth model or native package state. [Reviewed layout plans](../api/layout-rebind.md) consume those values and produce a separate frozen review sidecar. Factory apply commits authoring state, then ordinary normalization recomputes Canonical IR.
