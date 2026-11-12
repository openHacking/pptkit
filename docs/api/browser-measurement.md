# Browser font-backed text measurement

`@pptkit/layout/dom` exports `createBrowserTextMeasurement(options)`. This opt-in adapter implements the [Layout measurement contract](layout.md#opt-in-measurement-and-diagnostics) using real browser CSS typography and loaded fonts. The default Layout entry never creates DOM nodes or loads fonts. Importing the DOM entry is safe without a browser; constructing it requires a live browser `Document`.

## Inputs and lifecycle

- `document`: a live document with a window, CSS layout and Font Loading API. Its root must not be transformed or CSS-zoomed. Detached/hidden documents fail explicitly.
- `fonts`: unique `{ family, faces, covers }` entries. Faces must belong to that exact family, already be loaded and be registered in the supplied document's `fonts` set when used. The adapter neither loads fonts nor registers them. The inventory arrays are copied; the FontFace objects remain application-owned and their live status/family is checked on each use.
- `covers(request)`: synchronous boolean confirming glyph **and style** coverage for family, effective size, bold/italic, language and text. Derive it from a trusted font inventory/cmap and supported font style variants. A simple “font exists” check is insufficient. Browser font synthesis is disabled. Unloaded or unregistered fonts, unverified glyphs and unsupported styles are rejected, not silently measured with a system fallback.
- `fallbackFamilies`: optional ordered family names from the inventory. Layout reports `font-fallback` when the selected family differs from the requested one. Without this list there is no permitted family substitution.
- `tabWrap`: optional `"approximate"` (default), `"powerpoint"`, or `"reject"` policy for tabbed lines requiring automatic wrapping, as described below.

The returned object contains synchronous `measure`, `measureEastAsianVertical`, and `resolveFont`, suitable for `LayoutOptions.textMeasurement`, plus idempotent `dispose()`. It owns one hidden, closed-shadow measurement host. Per-call nodes are removed after measurement; disposal removes the host. Calls after disposal/detachment fail with Layout diagnostics. The application retains ownership of FontFace registration and any font buffers/URLs.

The factory rejects malformed inventory/fallback configuration before attaching a host. Runtime provider failures use the existing `font-resolution-failed` or `text-measurement-failed` diagnostics. Unavailable fonts return `font-unavailable`. Empty inventories are permitted and report all requested fonts unavailable.

## Measured semantics

The adapter handles mixed runs/sizes and families, CJK browser line breaking, explicit and consecutive line breaks, empty paragraphs and trailing blank lines, wrapping/no-wrap, numbered and character markers, indentation/hanging, alignment, paragraph base direction, and paragraph spacing. Layout applies margins, normalized shrink scaling and vertical frame alignment. Alphabetic numbering continues beyond Z (AA, AB, …). Marker glyphs must also be covered by the first run's resolved font; otherwise the body is explicitly unmeasurable. Explicit RTL tabbed paragraphs are rejected because the bounded tab placement algorithm assumes LTR geometry.

Explicit tab stops are measured using loaded browser font advances for left-aligned paragraphs when hanging indent does not exceed the paragraph left indent. A no-wrap frame is measured as authored. By default, a wrapping frame supports explicit lines that fit, including table cells, and a bounded automatic-break case: with no list or indent, one tab separates an ASCII alphanumeric prefix from at most 256 ASCII alphanumeric words or decimal numbers separated by single spaces. Each word must fit the frame; the aligned segment must not overlap the prefix. The first word stays at its aligned position if it fits, otherwise moves to the next line; subsequent words may wrap at spaces. Other lines requiring automatic wrapping remain unmeasurable under the default policy. Without a list marker, stops use the paragraph left margin while hanging indent moves the first-line start; one [LibreOfficeDev ODF control](../testing/tab-stops-libreoffice.md) supports this coordinate relationship. For a `•` bullet or `arabicPeriod` number, the first resolved run must be nonempty and cover the marker. The marker uses that run's font, occupies at least the hanging width, and stops use the content frame's left edge, following the named [LibreOfficeDev list controls](../testing/tab-list-libreoffice.md). Other markers and initial empty runs are rejected in this tabbed profile. Marker font selection is application-dependent and not proven identical to Office.

This profile accepts printable ASCII, Unicode Han, precomposed hiragana/katakana in U+3041–3096 and U+30A1–30FA, and precomposed Hangul syllables in U+AC00–D7A3, with mixed runs, explicit CR/LF breaks and up to 256 explicit stops. Stops restart on each hard line; CRLF across run boundaries is one break. Each tab requires an available stop, decimal alignment requires a period, and the aligned segment must not overlap preceding text except for the opt-in PowerPoint case below. A line cannot end with a tab. For bounded DOM work, a tabbed paragraph permits at most 4096 tab, CR and LF characters in total. Unsupported cases—including default tab intervals, automatic wrapping outside the bounded case above, hanging beyond the left indent, other paragraph alignments, other scripts and combining marks, and line-ending tabs—report `text-measurement-failed`; no measurement or overflow conclusion is returned for that body. Other application-supplied measurement providers retain their own contracts. This browser profile is not a claim of PowerPoint equivalence.

The default `tabWrap: "approximate"` preserves the browser estimate described above. [PowerPoint narrow-frame controls](../testing/tab-stops-libreoffice.md) show that it can split a word and alter decimal positioning where the approximation moves whole words, so the estimate can differ materially from PowerPoint. The opt-in `tabWrap: "powerpoint"` models bounded cases: no list or indent, one left, center or right tab after an ASCII alphanumeric prefix followed by at least two ASCII alphabetic words, or one decimal tab followed by an ASCII decimal number and at least one alphabetic word. Items have single separating spaces, each must fit the full frame, and at most 256 items are allowed. For left tabs, when the first word does not fit after the stop, the tab moves to the next line and the word may split there; if even its first character does not fit after the stop, the text moves past an empty line. For center, right and decimal tabs, the longest initial sequence that fits after the prefix may shift within the frame before remaining words wrap; if even the first item cannot fit beside the prefix, measurement fails. This overlap adjustment applies only in wrapping frames; no-wrap overlap remains unmeasurable. Other automatically wrapped or overlapping cases report `text-measurement-failed`. Native controls cover two- and three-item strings; behavior for longer sequences is an extrapolation and this is not a general PowerPoint layout engine. Browser fonts and line spacing still determine its metrics.

Set `tabWrap: "reject"` when an application prefers no measurement in that situation: explicit lines that fit and no-wrap frames remain measurable, while a tabbed line requiring automatic wrapping according to the browser's measured advances reports `text-measurement-failed`. This does not prove that PowerPoint will fit a line the browser considers fitting: font selection and native layout can differ. The policy is captured when the provider is created and does not change other providers or non-tabbed paragraphs.

One logical point is represented by one CSS pixel during measurement, matching the renderer's internal user-unit convention. The returned bounds include typographic text rectangles and paragraph line/spacing extent, not a pixel scan of painted glyph ink. Decorations, shadows, stroke extents and other effects are outside this adapter's bounds. Extremely small line spacing can allow font rectangles outside their line boxes; these are included in overflow analysis.

Text is assigned as literal `textContent`, never HTML. The shadow tree and reset host styles isolate measurement from ordinary document styles. Root transforms/zoom or missing CSS layout are diagnosed. Callers must keep font resources and font-family definitions stable during an operation; the same text is not guaranteed to produce identical metrics across different browser/font versions.

## Example

<!-- doc-test: docs/examples/browser-measurement.ts -->
```ts
import { createPresentation } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { createBrowserTextMeasurement } from "@pptkit/layout/dom";

// The application supplies licensed font bytes and their verified cmap coverage.
export async function measureInBrowser(fontBytes: ArrayBuffer, codePoints: ReadonlySet<number>) {
  const face = await new FontFace("Application Font", fontBytes).load();
  document.fonts.add(face);
  const adapter = createBrowserTextMeasurement({
    document,
    fonts: [{
      family: "Application Font",
      faces: [face],
      covers: ({ text, bold, italic }) => !bold && !italic && [...text].every(
        (character) => /[\n\r\t]/.test(character) || codePoints.has(character.codePointAt(0)!),
      ),
    }],
  });
  try {
    const presentation = createPresentation({ theme: { fonts: { body: "Application Font" } } });
    presentation.addSlide().addElement({
      type: "text", content: "Measured text", box: { x: 20, y: 20, width: 100, height: 30 },
    });
    return resolveLayout(presentation, {
      textMeasurement: adapter,
      diagnostics: { textOverflow: true, slideBounds: true },
    });
  } finally {
    adapter.dispose();
    document.fonts.delete(face);
  }
}
```

## Evidence and fidelity boundary

[Browser tests](../../packages/svg-renderer/test/measurement.spec.mjs) use pinned Roboto Mono and a licensed [Noto Sans SC subset](../../packages/layout/test/fixtures/README.md), block external requests, compare Latin advances to an independent Canvas measurement, and verify CJK/mixed-run wrapping, margins, spacing, lists, missing coverage, disposal, font mutation and host CSS isolation. Chromium is the currently exercised measurement engine; Firefox/WebKit remain untested.

The adapter supplies browser text analysis. It does not change the SVG renderer's current native/simple-text line-breaking algorithm or the text saved into PPTX. Native Office metrics and shared resolved line rendering remain separate work. No font is embedded into exported presentations by enabling measurement.

An isolated [48-case LibreOfficeDev PDF comparison](../testing/tab-stops-libreoffice.md) checks the supported tab profile with four indent controls and named, metric-compatible Liberation font candidates. It does not certify PowerPoint or broader tab layout.

The font-loading policy follows the [CSS Font Loading specification](https://www.w3.org/TR/css-font-loading/): `check()` concerns whether rendering would require unloaded fonts, and alone does not establish glyph coverage. Text rectangles follow [CSSOM View](https://www.w3.org/TR/cssom-view/), which applies ancestor transforms; the adapter rejects transformed/zoomed document roots instead of interpreting screen coordinates as presentation coordinates.

The optional [installed-font experiment](../testing/real-fonts.md) additionally derives regular-style coverage from actual cmap ranges, verifies TTC extraction with HarfBuzz glyph/position equality and checks Chromium's actual PostScript faces through CDP. Eight named samples agree in advance within 0.013px; missing Han, unprovided bold and emoji produce located rejection, and explicitly permitted Han fallback retains the author run. These observations do not expand the adapter's contract to mixed bidi, emoji shaping, native Office layout or every installed font.

Run styles with `rightToLeft: true` request DrawingML selective run-level direction and complex-script font selection. SVG preview reports `text-run-rtl-approximation`; the built-in browser text measurer rejects these runs because browser layout cannot reproduce the selective override. Explicit `false` does not trigger this limitation.

For fixed East Asian vertical frames and cells, a custom provider may implement `TextMeasurementOptions.measureEastAsianVertical`. The immutable request gives `availableInlineSize` (physical height) and `availableBlockSize` (physical width), excluding frame insets. Its returned bounds use inline `x/width` increasing downwards and block `y/height` increasing leftwards, before frame alignment. Layout maps those bounds back to physical coordinates and aligns the columns from the right edge. The browser adapter implements this callback with CSS `vertical-rl` and mixed glyph orientation. It rejects tabs, list markers, RTL paragraphs and DrawingML run-level RTL overrides in this mode. Loaded-font and coverage requirements are unchanged. Without a vertical callback, Layout reports that vertical bounds are unavailable. The callback also supports intrinsic sizing of rectangular text frames. Intrinsic sizing probes with a zero available block size to request natural column thickness; zero inline size in an unwrapped probe does not clip the text. Automatic row sizing for East Asian cells also uses this callback: it probes the natural inline extent, can search for wrapping under a row cap, and remeasures the resolved cell span.

The browser adapter also exposes `measureVertical`, which accepts a `textDirection` discriminator for East Asian, Mongolian, and both stacked modes. Logical block coordinates advance from the starting column edge: right for East Asian and stacked RTL, left for Mongolian and stacked LTR. If both callbacks are supplied, `measureEastAsianVertical` takes precedence for East Asian text. All four modes support fixed bounds, rectangular intrinsic sizing, and automatic table row sizing. Mongolian script shaping depends on the browser and loaded fonts; the adapter does not substitute native Office metrics.
