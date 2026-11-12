# Exact-face SVG font evidence

The [optional runner](../../scripts/inspect-real-font-svg.mjs) consumes the earlier [real-font shaping evidence](real-fonts.md) and verifies original/extracted font hashes before loading the same bytes. It renders eight named regular-face samples twice through PPTKit's actual SVG API: native `<text>/<tspan>` and mixed-run `foreignObject` with a smaller Latin prefix. These are SVG specimens, not replacement HTML-only samples.

Each font registers under a unique hash-derived family alias. The alias deliberately contains a numeric name component; it is not an installed host family. CDP checks the actual custom PostScript identity and positive glyph count on the rendered SVG body node. The measured native SVG text length / XHTML range advance must agree with the original-face HarfBuzz advance within **0.05px**. All 16 cases pass, with no render warnings or network requests.

The harness independently embeds those exact bytes in a copy of each SVG for an SVG-as-image control; this historical experiment did not use the renderer's later optional embedded-font input. Each inline SVG and image is captured at the same unscaled 550×120 viewport. The [Python PNG oracle](../../scripts/research/oracles/real-font-svg.py) requires visible dark ink and exact raster equality for every pair, records ink bounds and PNG hashes, and verifies that the Latin alias image **without** embedded font differs. This negative control prevents a coincidentally installed family or the parent document's FontFace registry from masking missing image-local fonts.

The [recorded manifest](real-font-svg-evidence.json) identifies current artifacts and source-evidence/font hashes. The original no-wrap probe emitted six native lines; after [ADR0052](../architecture/decisions/0052-svg-text-boundaries-and-font-names.md), it emits one. The initial 16-case contact sheet was viewed, showing visible Han variants, Hangul, Arabic joining/marks, Hebrew marks and the Latin accent; the final alias-based sheet was separately inspected. Each exact-face image matches its inline SVG pixel-for-pixel in the recorded Chromium environment. Native and XHTML baselines differ by their documented line-box policies; equality is checked within each path's inline/image pair, not asserted between the two paths.

The [PNG corruption checks](../../scripts/research/oracles/real-font-svg-checks.py) replace an actual image with white, alter a pixel or change dimensions. The same PNG oracle must reject each. Width alone, CSS family declarations alone and invisible text extraction are insufficient evidence.

Reproduce after building browser bundles and generating the shaping runner's local artifacts:

```sh
node scripts/inspect-real-font-shaping.mjs
node scripts/inspect-real-font-svg.mjs
# Use the artifactsDirectory from the newly recorded manifest:
python3 scripts/research/oracles/real-font-svg-checks.py ARTIFACTS_DIRECTORY 16
```

Only code, metadata and hashes belong in the repository; licensed font binaries and rendered artifacts remain local. No proprietary fonts are installed, downloaded or redistributed. The runner's data-font injection is an experiment operation, not a library resource-resolution contract.

This advances IMP-05.1.3.3.2.5.2's actual SVG slice. The eight string samples and one Latin-prefix mixed-run form do not certify mixed bidi paragraph ordering, emoji color/ZWJ rendering, every script, browser/native line-layout equality, Firefox/WebKit or non-browser SVG import. Structural grapheme regressions preserve source sequences without claiming emoji font coverage. Native fresh-PPTX/edit/theme checks and complete font parent tasks remain open.
