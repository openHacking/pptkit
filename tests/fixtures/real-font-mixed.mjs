/** Each author run intentionally crosses three script-selected literal families. */
export function createRealFontMixedFixture(createPresentation, families) {
  const sourceMap = { Latn: families.latin, Hans: families.sc, Hant: families.tc, Jpan: families.ja, Hang: families.ko, Arab: families.ar, Hebr: families.he, Grek: families.greek };
  const samples = [
    ["sc", "zh-CN", "骨直令返"], ["tc", "zh-TW", "骨直令返"], ["ja", "ja", "骨直令返"],
    ["ko", "ko", "한글"], ["ar", "ar", "السَّلَامُ"], ["he", "he", "שָׁלוֹם"],
  ];
  const document = createPresentation({ size: { width: 720, height: 260 }, theme: { fonts: { body: families.latin, heading: families.latin } } });
  const controls = [];
  for (const [fontId, language, text] of samples) {
    const id = `mixed-${fontId}`, fullText = `Latin ${text} Ω`;
    const slide = document.addSlide({ id });
    slide.addElement({ id: `label-${id}`, type: "text", name: "Label", box: { x: 20, y: 20, width: 680, height: 40 }, content: [{ runs: [{ text: id, style: { fontFamily: families.latin, fontSize: 16 } }] }] });
    slide.addElement({ id: `probe-${id}`, type: "text", name: "Probe", box: { x: 20, y: 90, width: 680, height: 110 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: fullText, style: { fontFamily: families.latin, fontSize: 48, color: "000000", language, scriptFontFamilies: sourceMap } }] }] });
    controls.push({ id, fontId, language, text: fullText, nativeText: text, slideId: id, elementId: `probe-${id}`, expectedRuns: [{ text: "Latin ", family: families.latin }, { text: `${text} `, family: families[fontId] }, { text: "Ω", family: families.greek }] });
  }
  return { document, sourceMap, controls };
}
