# Quality gates

## Required automated checks

In browser mode, require a successful session import and SVG render. In Node mode, run `npm run build`. Require:

- no PPTKit validation errors
- no element outside the slide boundary
- no unexpected pairwise overlap classified as high risk
- no missing required image
- no exporter warning unless explicitly accepted and reported
- before export: one preview SVG per planned slide and no blocking session/layout/asset issue
- after the user requests export: a readable ZIP containing content types, presentation, slide, and relationship parts
- after export: XML-shaped package parts with matching slide count

Warnings are deliverables, not console noise. Keep them in browser findings and `build-report.json`, or in the Node `output/build-report.json`.

## Optional rendering

The SVG browser preview is the default review surface and is not pixel-identical to Office. In Node fallback, run `npm run render` when strict application rendering is requested. The script checks for LibreOffice, `pdftoppm`, and ImageMagick `montage`:

- LibreOffice converts the PPTX to PDF.
- Poppler renders page PNGs.
- ImageMagick creates `contact-sheet.png` when available.

If any capability is absent, record `skipped` or `partial`; do not fail an otherwise structurally valid build. Ask the user to open the deck manually in PowerPoint or LibreOffice.

## Visual review

Inspect every rendered slide for:

- clipped or unexpectedly wrapped text
- low contrast or unreadable captions
- collisions and accidental tangencies
- inconsistent margins, alignment, or page rhythm
- stretched/cropped evidence and substituted fonts
- unsupported glyphs or missing media

## Adversarial review

Attack at least these failure modes before delivery:

1. Dependency or PPTKit version drift prevents a clean project install.
2. Long translated text overflows a layout that worked in another language.
3. Temporary, remote, or renamed image paths disappear at export time.
4. Font substitution changes wrapping or hierarchy across PowerPoint and LibreOffice.
5. IndexedDB, browser API, HTTPS URL, or asset-size failures leave a session that appears saved or complete.
6. A host without structured forms skips a required product decision.

Fix the underlying source or layout. Do not dismiss a failure because a `.pptx` file exists.
