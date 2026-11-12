# ADR0052: SVG text boundaries and literal font names

Status: accepted for the existing SVG preview APIs, 2026-10-26. No public option/type or dependency is added.

## Problem and decision

[Exact-face SVG controls](../../testing/real-font-svg.md) found that native SVG ignored `frame.wrap: false`, collapsed authored whitespace and repeated/trailing hard breaks, and split extended grapheme clusters under narrow wrapping. Separate controls found that unquoted native text/chart font-family attributes interpreted literal names as CSS syntax. The XHTML path also collapsed wholly empty paragraphs and the final empty line after a break.

Keep line emission inside Renderer until Layout owns a resolved line-fragment contract. Pass the actual wrap flag into native line construction: disabling soft wrapping must still honor explicit CRLF, CR and LF breaks. Preserve authored leading/interior/trailing whitespace; whitespace used as a soft line-wrap opportunity is omitted at the new line. Emit explicit whitespace-preserving SVG text attributes. Repeated and final hard breaks contribute empty lines to vertical placement.

Segment joined paragraph text with `Intl.Segmenter` in grapheme mode **before** redistributing it into the selected-font fragments. A font-family boundary cannot split a combining, ZWJ or regional-indicator cluster across lines. Each cluster is indivisible, including when wider than the frame; overflow is preferable to splitting its source sequence. Existing deterministic width estimates and baseline calibration remain approximate. Nonspacing/enclosing marks and join controls add no separate estimated advance. Grapheme behavior follows the runtime's Unicode/Intl implementation; this is not a font shaper, full Unicode line-break algorithm or bidi engine. A runtime without Intl grapheme support uses the existing escaped XHTML fallback instead of failing module import.

For XHTML, append an `aria-hidden` zero-width display placeholder only for a wholly empty paragraph or a final hard-break line. It uses the last effective run's style. It is outside authored `span[lang]` runs and does not alter Core input or source text. This aligns visible paragraph heights with the existing browser measurement adapter's empty-line policy.

Move the existing quoted CSS family encoder into the Renderer escaping module and use it for native text, chart labels and XHTML. Family strings are literal names, not caller-supplied CSS stacks. Escape CSS quotes/backslashes and XML separately, including names containing numeric words, commas and quotes. Native SVG font-family attribute values now contain quoted CSS family syntax; they should not be read as raw OOXML typeface values.

## Boundaries and evidence

At the time of this decision the renderer performed no font I/O, installation,
registration or embedding. The optional experiment owns its exact font bytes and
aliases. A later [explicit embedded-font input](../../api/svg-renderer.md#explicit-embedded-fonts)
allows caller-supplied programs in SVG; it does not perform font I/O or validate
the complete program. Author input and detached Layout state remain unchanged.
Analysis diagnostics remain informational and do not supply shaped line geometry.

[Six browser regressions](../../../packages/svg-renderer/test/measurement.spec.mjs) cover no-wrap/space advances against pinned Canvas and XHTML metrics, repeated CR/LF and final-empty placement, grapheme integrity across selected-family boundaries, safe fallback without Segmenter, pinned-font advances for literal names in native text/chart labels, and XHTML empty-line heights against browser measurement. Existing baseline/wrap/visual expectations remain checked; no visual baselines were replaced.

[Sixteen exact-face controls](../../testing/real-font-svg.md) exercise eight regular-face samples across seven installed font families through native and mixed-run SVG. Actual custom PostScript faces, HarfBuzz advances, visible PNG ink and exact inline/image raster equality are checked with independent corruption controls. A missing embedded alias font changes the isolated image, ruling out an inherited document-font false positive. This certifies the named Chromium observations only. Fresh-PPTX native editing, complete shaping/bidi/emoji and non-browser `foreignObject` portability remain open.
