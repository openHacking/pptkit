# Assets

Core owns asset identity and descriptive metadata. It deliberately does not read local files, download URLs, store binary payloads, hash bytes, or create PPTX media relationships.

## `registerAsset(input)`

```ts
registerAsset(input: PresentationAssetInput): PresentationAsset;

interface PresentationAssetInput {
  id?: string;
  kind: "image";
  source: { type: "path" | "url"; value: string };
  mimeType?: string;
  width?: number;
  height?: number;
  accessibility?: {
    description?: string;
    decorative?: boolean;
  };
  dedupeKey?: string;
}
```

| Field | Required | Description |
| --- | --- | --- |
| `id` | no | Stable ID; generated when omitted. |
| `kind` | yes | Currently only `image`. |
| `source` | yes | Environment-neutral path or URL descriptor. |
| `mimeType` | no | Declared media type; exporter may infer common raster types when omitted. |
| `width`, `height` | no | Positive source dimensions used for `contain` and `cover` layout. |
| `accessibility` | no | Asset-level description and decorative flag. |
| `dedupeKey` | no | Caller-controlled identity for repeat registration of the same source. |

The returned asset is immutable and belongs to the document.

```ts
const hero = presentation.registerAsset({
  kind: "image",
  source: { type: "path", value: "./assets/hero.png" },
  mimeType: "image/png",
  width: 1600,
  height: 900,
  accessibility: { description: "Product launch overview" },
  dedupeKey: "launch-hero",
});

slide.addElement({
  type: "image",
  assetId: hero.id,
  box: { x: 420, y: 80, width: 460, height: 260 },
  fit: "cover",
});
```

## Deduplication and conflicts

- An explicit `id` identifies an existing asset.
- Without an explicit ID, `dedupeKey + source.type + source.value` identifies an existing asset.
- Re-registering the same identity with identical metadata returns the existing immutable asset.
- Re-registering the same identity with different metadata throws an error.
- Reusing an ID for a different asset throws an error.

## `getAsset(assetId)`

```ts
getAsset(assetId: string): PresentationAsset | undefined;
```

Returns the registered immutable asset or `undefined`. It does not load or verify the source.

## Runtime loading boundary

The browser-neutral exporter loads URL sources and rejects path sources with a warning. The Node.js exporter loads both URL and local path sources. See [PPTX exporter](../pptx-exporter.md) for runtime behavior.

## Validation

Validation reports empty source values, non-positive dimensions, and image elements that reference missing assets. Asset loading failures happen later in the exporter and are returned as export warnings so other valid content can still be packaged.
