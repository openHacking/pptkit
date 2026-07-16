# Browser workflow

Use this workflow only after the confirmation gate.

## Preconditions

- Resolve the HTTPS preview application in this order:
  1. A URL explicitly supplied by the user or host for the current task.
  2. `PPTKIT_PREVIEW_URL`, when available.
  3. The official PPTKit preview application at `https://openhacking.github.io/pptkit/`.
- Require the resolved application to support `DeckSessionV1` with `schemaVersion: 1`. Never invent another deployment URL.
- Require modern browser support for `fetch`, `Blob`, `URL`, typed arrays, `structuredClone`, and IndexedDB.
- Fall back to the Node workflow when the resolved URL is unreachable or incompatible, IndexedDB or required APIs are unavailable, one inline asset exceeds 5 MB, all inline assets exceed 20 MB, or strict Office/LibreOffice rendering is required.

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

1. Open the resolved HTTPS URL with the host's browser capability. In Codex, the listed in-app Browser skill is an available preview channel and must be loaded and followed.
2. If Codex browser controls are not directly visible, discover `browser:control-in-app-browser` (or the equivalent in-app Browser skill) and `node_repl js`. Follow the Browser skill to initialize its runtime, explicitly select the `iab` browser, make it visible for the user-facing preview, and open or reuse the resolved URL. Do not give up solely because the initial tool list omits browser controls.
3. Treat a successful open or focus operation as proof that the preview channel is available. Only fall back after the Browser skill's setup or navigation actually fails; name the failed step and preserve the resolved preview URL as a direct review link.
4. Paste the complete `deck-session.json` into **Import DeckSessionV1** and activate **Import and preview**.
5. Confirm the title, theme, revision, slide count, one SVG per slide, IndexedDB save status, and the complete findings list.
6. Inspect every slide in the stage or thumbnail gallery. Treat blocking findings as failures and renderer warnings as required review items.
7. Keep the preview tab as a deliverable tab so the user can review it before export. In Codex, finalize the browser session with this tab marked `deliverable`; do not clean it up as an intermediate research tab.

Do not download automatically. Export is allowed when the user clicks **Generate & download PPTX** or explicitly asks the agent to trigger the export/download after preview.

## Revise

Translate feedback into the brief or slide plan, increment `revision`, refresh `updatedAt`, and import the full session again. Preserve semantic slide IDs so the application keeps the same slide selected and reports changed pages.

## Browser source extraction

The preview application accepts TXT/Markdown, PDF, DOCX, CSV/XLS/XLSX, PNG/JPEG/GIF, and SVG through `File`/`Blob`. It stores extracted evidence and binary assets in IndexedDB without uploading them. PPTX input remains an unsupported manual reference.

When the host cannot automate a native file picker, inspect attached sources with host-native file tools and place normalized evidence in `sources`; use the Node fallback for large binary assets that cannot be transferred safely.

## Export

When the user clicks **Generate & download PPTX**, or explicitly asks the agent to export/download, the application generates bytes, inspects ZIP/XML parts, downloads `build-report.json`, and downloads the PPTX only when package checks pass. Browser preview does not replace a final PowerPoint or LibreOffice review.
