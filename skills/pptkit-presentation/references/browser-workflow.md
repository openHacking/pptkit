# Browser workflow

Use this workflow only after the confirmation gate.

## Preconditions

- Resolve the HTTPS application from `PPTKIT_PREVIEW_URL` or the host's configured PPTKit preview URL. Never invent a deployment URL.
- Require modern browser support for `fetch`, `Blob`, `URL`, typed arrays, `structuredClone`, and IndexedDB.
- Fall back to the Node workflow when the URL is missing or unreachable, IndexedDB or required APIs are unavailable, one inline asset exceeds 5 MB, all inline assets exceed 20 MB, or strict Office/LibreOffice rendering is required.

## Create the session

Create these artifacts without initializing a Node project:

```text
project/
├── deck-brief.md
├── deck-session.json
└── content/sources.json
```

Use this top-level shape:

```json
{
  "schemaVersion": 1,
  "id": "stable-deck-slug",
  "revision": 1,
  "createdAt": "ISO-8601",
  "updatedAt": "ISO-8601",
  "deck": { "brief": {}, "slides": [] },
  "sources": [],
  "assets": []
}
```

Use stable source, asset, and slide IDs. Put small images in `assets[].dataUrl`; do not inline an asset over 5 MB or more than 20 MB in total. Use the browser's local-file control for larger user-selected assets or switch to Node when unattended automation is required.

## Open and verify

1. Open the configured HTTPS URL with the host's browser capability. In Codex, load and follow the in-app Browser skill.
2. Paste the complete `deck-session.json` into **Import DeckSessionV1** and activate **Import and preview**.
3. Confirm the title, theme, revision, slide count, one SVG per slide, IndexedDB save status, and the complete findings list.
4. Inspect every slide in the stage or thumbnail gallery. Treat blocking findings as failures and renderer warnings as required review items.
5. Keep the preview tab as a deliverable tab so the user can review it without downloading a file.

Do not activate **Generate & download PPTX** for the user. That control is the user's explicit export decision.

## Revise

Translate feedback into the brief or slide plan, increment `revision`, refresh `updatedAt`, and import the full session again. Preserve semantic slide IDs so the application keeps the same slide selected and reports changed pages.

## Browser source extraction

The preview application accepts TXT/Markdown, PDF, DOCX, CSV/XLS/XLSX, PNG/JPEG/GIF, and SVG through `File`/`Blob`. It stores extracted evidence and binary assets in IndexedDB without uploading them. PPTX input remains an unsupported manual reference.

When the host cannot automate a native file picker, inspect attached sources with host-native file tools and place normalized evidence in `sources`; use the Node fallback for large binary assets that cannot be transferred safely.

## Export

When the user clicks **Generate & download PPTX**, the application generates bytes, inspects ZIP/XML parts, downloads `build-report.json`, and downloads the PPTX only when package checks pass. Browser preview does not replace a final PowerPoint or LibreOffice review.
