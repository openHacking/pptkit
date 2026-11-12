/** One paragraph contains all seven requested script/language cases plus Greek. */
export function createRealFontCombinedFixture(createPresentation, families) {
  const sourceMap = { Latn: families.latin, Hans: families.sc, Hant: families.tc, Jpan: families.ja, Hang: families.ko, Arab: families.ar, Hebr: families.he, Grek: families.greek };
  const expectedRuns = [
    ["Latin ", "en-US", "latin"], ["骨 ", "zh-CN", "sc"], ["骨 ", "zh-TW", "tc"], ["骨 ", "ja", "ja"],
    ["한 ", "ko", "ko"], ["لا ", "ar", "ar"], ["שָׁ ", "he", "he"], ["Ω", "el", "greek"],
  ].map(([text, language, fontId]) => ({ text, language, fontId, family: families[fontId] }));
  const document = createPresentation({ size: { width: 720, height: 260 }, theme: { fonts: { body: families.latin, heading: families.latin } } });
  const slide = document.addSlide({ id: "combined-scripts" });
  slide.addElement({ id: "label-combined", type: "text", name: "Label", box: { x: 20, y: 20, width: 680, height: 40 }, content: [{ runs: [{ text: "combined-scripts", style: { fontFamily: families.latin, fontSize: 16 } }] }] });
  slide.addElement({ id: "probe-combined", type: "text", name: "Probe", box: { x: 20, y: 90, width: 680, height: 110 }, frame: { margin: 0, wrap: false }, content: [{ runs: expectedRuns.map(r => ({ text: r.text, style: { fontFamily: families.latin, fontSize: 48, color: "000000", language: r.language, scriptFontFamilies: sourceMap } })) }] });
  return { document, sourceMap, expectedRuns };
}
