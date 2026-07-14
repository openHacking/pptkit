# Quality gates

## Required automated checks

Run `npm run build` and require:

- no PPTKit validation errors
- no element outside the slide boundary
- no unexpected pairwise overlap classified as high risk
- no missing required image
- no exporter warning unless explicitly accepted and reported
- a readable ZIP containing content types, presentation, slide, and relationship parts
- XML-shaped package parts with matching slide count

Warnings are deliverables, not console noise. Keep them in `output/build-report.json`.

## Optional rendering

Run `npm run render`. The script checks for LibreOffice, `pdftoppm`, and ImageMagick `montage`:

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
5. A host without structured forms skips a required product decision.

Fix the underlying source or layout. Do not dismiss a failure because a `.pptx` file exists.
