# Installed-font shaping and native font observations

These experiments advance IMP-05.1.3.3.2.5/.6. They introduce no production API, change no source selection rule and grant no source font/style write permission. All fonts were already installed on this Mac. Font binaries and extracted TTC faces stay under `.tmp`; only code and metadata belong in the repository.

## Exact faces in HarfBuzz and Chromium

The [optional runner](../../scripts/inspect-real-font-shaping.mjs) records source file SHA-256, face index, Fontconfig font version, extracted-face SHA-256 and actual cmap ranges. `hb-subset --keep-everything` extracts individual TTC faces locally. HarfBuzz shapes both the original indexed face and extracted bytes; complete glyph IDs, clusters, offsets and advances must agree. This is a shaping equivalence check for these strings, not byte identity of font containers.

The same extracted bytes load through `FontFace` into Chromium with network requests blocked. The application inventory derives coverage from actual cmap ranges and admits only regular style. The public browser measurement adapter measures the samples. CDP `CSS.getPlatformFontsForNode` independently checks the actual PostScript name, custom-face identity and glyph count; a CSS family declaration or `document.fonts.check()` alone is insufficient evidence.

The [manifest](real-font-shaping-evidence.json) records eight samples: simplified/traditional/Japanese Han variants, Korean Hangul, marked Arabic, lam-alef, marked Hebrew and a combining Latin acute. All eight use their exact intended face, produce no Layout diagnostic and agree with HarfBuzz advances within **0.013px** at 48px. Layout and visible DOM advances agree within **0.02px**. Source/extracted HarfBuzz results agree exactly and contain no glyph zero. The screenshot was inspected: Han variants, Hangul, Arabic joining/marks, Hebrew marks and the Latin accent are visible. These are observed samples, not certification of every glyph or bidi sequence.

Actual Arial cmap coverage rejects Han and the emoji ZWJ control; the inventory also rejects unprovided bold. Each produces a located `font-unavailable`. Explicitly permitting Hiragino Sans GB for the Han request produces a located `font-fallback`, while the author run remains unchanged. Emoji is **unavailable** in this experiment, not supported shaping. Later [actual SVG controls](real-font-svg.md) separately exercise these faces through native/mixed-run SVG and isolated image rasters. Mixed LTR/RTL paragraph ordering, emoji variation/color fonts and native line-layout parity remain open.

Reproduce after `pnpm build:browser`, with installed fonts, HarfBuzz, Fontconfig and Playwright Chromium available:

```sh
node scripts/inspect-real-font-shaping.mjs
```

The runner fails on a missing requested font, changed extraction shaping, zero glyph, unexpected actual browser face, advance disagreement or missing negative-control diagnostic. It does not install fonts or change global font configuration. Local artifacts are listed in the manifest; the final run is `.tmp/real-font-shaping/run-iAgNLm`.

## LibreOffice font discovery is part of the environment

The [native runner](../../scripts/inspect-real-font-native.mjs) builds [20 isolated source controls](../../scripts/research/font-native-cases.mjs), with direct slots, explicit/empty generic theme slots, supplemental maps, missing EA/CS, cross-slot references and local overrides. Its [independent oracle](../../scripts/research/oracles/real-font-native.py) verifies actual source hashes, named source run text/slot attributes and actual PDF body fonts/glyph traces before/after native save.

The bundled LibreOfficeDev **26.8.0.0.alpha0** does not discover the host fonts with its default environment in the initial six-slide probe: its CJK text was visually absent. Supplying a per-run Fontconfig file with `/System/Library/Fonts` and `/Library/Fonts` makes the intended direct CJK fonts visible. The override and cache are confined to child-process environment and the run directory. This finding does not retroactively turn earlier missing-glyph observations into passes. See [Fontconfig configuration and environment](https://fontconfig.pages.freedesktop.org/fontconfig/fontconfig-user.html).

The [LibreOffice manifest](real-font-native-evidence.json) retains all 20 observations and process failures. Direct SC/Arabic/Hebrew and literal local overrides select intended faces. Many theme supplemental cases use Arial Unicode MS (Greek uses Arial), differing from the current Parser v2 selection. A successful conversion or extracted text does not establish shaping, visibility or PowerPoint behavior.

```sh
node scripts/inspect-real-font-native.mjs /absolute/path/to/soffice
```

## PowerPoint comparison and three native save/reopen rounds

Microsoft PowerPoint for Mac **16.104** opened the controls, exported PDFs and saved three distinct PPTX copies. Each copy was closed, reopened and exported through the native UI. No repairs were accepted. The initial attempt to assemble controls used colliding master/layout IDs and nonstandard nested locations; PowerPoint requested repair, so it was cancelled. The corrected named-fixture assembler uses conventional part locations and unique IDs. Its actual slide and theme payload hashes must equal the original per-case sources; all internal relationships must resolve.

The [selection manifest](powerpoint-font-selection-evidence.json) identifies actual PDF body fonts. **15/20** match the Parser-selected family. The five differences remain explicit:

| Control | Parser-selected face | PowerPoint PDF face | Interpretation boundary |
| --- | --- | --- | --- |
| `theme-explicit-ko` (Han string with `ko`) | Apple SD Gothic Neo | Microsoft YaHei | Later paired controls below establish coverage and isolate `ko` versus `ko-KR`; this difference cannot be attributed to missing Han glyphs. |
| `latin-only-tc` (EA/CS absent) | Songti TC | Microsoft YaHei | The experimental Latin-collection supplemental fallback is not native selection parity for this input. |
| `cross-slot-ea` (`ea=+mn-lt`) | Arial | Microsoft YaHei | Actual native fallback must remain distinct from a raw slot reference. |
| `bare-zh-empty` (`lang=zh`) | Arial | Hiragino Sans GB | Core deliberately does not infer a Han region from bare `zh`; native environment chooses a face. |
| `missing-glyph-latin` | Arial | Microsoft YaHei | A requested family is not a glyph-coverage guarantee. |

Explicit and empty matching EA/CS theme references do select the intended SC/TC/ja/Arabic/Hebrew supplemental faces in PowerPoint; Greek also matches. The LibreOffice differences therefore cannot justify replacing v2 wholesale with LibreOffice behavior. The default source profile and v2 declared fallback semantics remain unchanged; native-parity corrections need a separately tested contract. No OS language heuristic is written back into source rules.

The [native rounds manifest](powerpoint-font-native-evidence.json) independently follows actual presentation → slide → layout → master → theme relationships after every save. It compares source run count/text/language/slot values, local list font declarations and theme minor generic/supplemental entries. All 20 retain those selected declarations in all three rounds. Each page's actual PDF raster equals its initial PowerPoint raster at the fixed PyMuPDF default scale. Every PPTX/PDF hash is recorded. A real ZIP with a deliberately wrong EA family is reread and rejected by the same comparison; unaffected pages remain equal. This is selected declaration/raster stability, not complete package byte preservation.

Reproduce the combined controls and independent checks using the LibreOffice runner's source directory:

```sh
python3 scripts/research/oracles/font-native-deck.py assemble SOURCE_DIR OUTPUT_DIR/font-controls.pptx
# Open the generated deck in PowerPoint and export font-controls.pdf.
python3 scripts/research/oracles/font-native-deck.py inspect SOURCE_DIR OUTPUT_DIR/font-controls.pptx OUTPUT_DIR/font-controls.pdf RESULT.json
# Save as font-controls-round1/2/3.pptx; close/reopen and export each PDF.
python3 scripts/research/oracles/font-native-deck.py native-rounds SOURCE_DIR OUTPUT_DIR/font-controls.pptx OUTPUT_DIR ROUNDS.json
```

The actual run is `.tmp/powerpoint-font-probe/font-controls-v2.pptx`; the filename reflects the corrected fixture, not a different Parser profile. The PowerPoint window and the 20-page PDF contact sheet were inspected. CJK is visible; the Arial Unicode Hebrew control shows different mark/spacing behavior, so text extraction is not promoted to general Hebrew shaping correctness. Text/style/theme edits, materialized author-map editing, native notes typography and full bidi/emoji support remain [separate TODOs](../planning/advanced-capability-todos.md#真实字体证据之后的执行卡).

## Paired native precedence controls

The later [32-control manifest](powerpoint-font-precedence-evidence.json) separates language tags, explicit/empty/absent native slots, concrete cross-slot references, changed/omitted supplemental entries and major/minor collections. Its actual source directory is `.tmp/real-font-native/run-lPDVJa`, with `.tmp/powerpoint-font-probe/font-precedence32.pptx` and matching PDF. All 32 source text/language/slot values, theme generic/supplemental declarations, slide order and actual slide→layout→master→theme chains are independently checked. **20/32** selected families match; all 12 differences remain observations. The contact sheet was inspected and PDF glyph traces contain no zero glyph; these checks are not complete shaping certification.

The earlier Korean interpretation needed correction: actual Apple SD Gothic Neo cmap ranges include both the Han and Hangul sample characters. A direct `ko-KR` Han control renders with that face. Theme Hangul with short `ko` uses Malgun Gothic, while identical theme references with `ko-KR` use Apple SD Gothic Neo, including empty generic EA. A `ko-KR` Han control with its Hang supplemental entry changed to Hiragino Sans GB uses that changed face. A short-`ko` control with a covering Arial Unicode MS Hang entry still uses other fonts. These are native language-tag differences for the tested PowerPoint environment, not evidence that `ko` is invalid BCP 47 or a reason to mutate imported language tags. Microsoft's [MinorFont example](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.minorfont?view=openxml-3.0.1) uses `Hang` as a supplemental script; there is no basis for replacing that script identifier to fit this application behavior.

With Latin explicitly resolving to the covering Arial Unicode MS, absent EA selects Arial Unicode MS for the TC sample, while a present `+mn-ea` selects Songti TC. For Arabic/Hebrew, absent or explicitly empty CS selects Arial in PowerPoint, while present `+mn-cs` selects the declared supplemental face. Thus the v2 Latin-collection supplemental fallback is contradicted as a **native selection parity** claim for these inputs; retaining source rules and existing v2 compatibility does not make that heuristic native-correct. The later [explicit v3 contract](../api/declared-slot-font-import.md) excludes that heuristic across drawings/tables/notes and captured candidates, reporting missing selections without guessing Office fonts. It preserves v2 compatibility and does not turn native implicit-font differences into resolved claims.

Covered `ea=+mn-lt` and the equivalent literal EA both use Hiragino Sans GB: the earlier cross-slot mismatch is consistent with actual glyph fallback, not a need to enable supplemental substitution on cross-slot references. Explicit Hant map mutation is honored; removing it exposes the generic face. Explicit major SC/TC/ja controls also honor their supplemental entries.

Bare `zh` with a present Hans entry follows that entry in this environment, including an alternate Songti TC value; removing Hans produces Microsoft YaHei. Core deliberately retains its region-ambiguous behavior for bare `zh`; explicit `zh-Hans`/`zh-Hant` controls match intended faces. This is an application/locale observation rather than a universal language-region inference.

The [negative-control runner](../../scripts/research/oracles/font-native-deck-checks.py) mutates actual ZIP slot/language/supplemental values, redirects a slide to another valid theme chain and swaps actual presentation slide order. It deliberately updates integrity hashes, requiring the semantic checks to reject each mutation. The slide-order mutation initially exposed a verifier gap: the inspector used manifest order without checking the actual presentation list. That check is now explicit, and all five mutations are rejected. Both the original 20 and new 32 controls pass the strengthened source checks. The original 20 generated slide/theme payloads remain byte-identical after extending the builder.

```sh
node scripts/inspect-real-font-native.mjs --precedence --prepare-only
python3 scripts/research/oracles/font-native-deck.py assemble SOURCE_DIR OUTPUT_DIR/font-precedence.pptx
# Open in PowerPoint and export the matching PDF.
python3 scripts/research/oracles/font-native-deck.py inspect SOURCE_DIR OUTPUT_DIR/font-precedence.pptx OUTPUT_DIR/font-precedence.pdf RESULT.json
python3 scripts/research/oracles/font-native-deck-checks.py SOURCE_DIR OUTPUT_DIR/font-precedence.pptx OUTPUT_DIR/font-precedence.pdf
```

`--prepare-only` invokes no native producer and needs no LibreOffice executable. Omit it to run LibreOffice conversions for the new controls and record a separate LO manifest, without replacing the historical baseline manifest. The separate v3 source contract has [independent preservation evidence](declared-slot-font-import.md); table fontRef/default-appearance native controls and native style/theme edits remain open under .6.1/.6.3.
