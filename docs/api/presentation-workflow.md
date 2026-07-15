# `@pptkit/presentation-workflow`

`@pptkit/presentation-workflow` provides the browser-neutral contracts and deterministic recipes used by the browser-first presentation skill and its Node fallback.

## Deck sessions

`DeckSessionV1` is the portable source of truth for browser review:

```ts
import type { DeckSessionV1 } from "@pptkit/presentation-workflow";

const session: DeckSessionV1 = {
  schemaVersion: 1,
  id: "quarterly-review",
  revision: 1,
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
  deck: { brief, slides },
  sources: [],
  assets: [],
};
```

Images use stable `assetId` values. `authorPresentation(deck, resolver)` asks the runtime adapter to resolve each ID into a browser URL or Node path without adding environment-specific fields to the deck spec.

## Source extraction

`extractSource` and `extractSources` accept `{ name, mimeType, bytes }`. Runtime adapters provide PDF, DOCX, workbook, and image parsers. Plain text parsing, source IDs, result normalization, and failure reporting remain shared.

`blobToSourceInput` is a convenience adapter for browser `File`/`Blob` values.

## Validation and package inspection

- `validateDeckSpec` checks semantic role requirements, IDs, density, and asset availability.
- `inspectStructure` checks resolved slide bounds and risky overlaps.
- `inspectPptxPackage(bytes)` checks required ZIP/XML parts directly from `Uint8Array` in either runtime.
- `parseDeckSession` validates schema version and enforces 5 MB per-inline-asset and 20 MB total-inline-asset limits.

The package has no filesystem, process, IndexedDB, DOM, UI, or network behavior.
