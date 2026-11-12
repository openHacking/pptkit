/** Generated directly through author APIs; no source XML patches. */
export const freshFontSamples = [
  { id: "sc", text: "骨直令返", language: "zh-CN", script: "Hans", bucket: "eastAsianFontFamily" },
  { id: "tc", text: "骨直令返", language: "zh-TW", script: "Hant", bucket: "eastAsianFontFamily" },
  { id: "ja", text: "骨直令返", language: "ja", script: "Jpan", bucket: "eastAsianFontFamily" },
  { id: "ko", text: "한글", language: "ko", script: "Hang", bucket: "eastAsianFontFamily" },
  { id: "ar", text: "السَّلَامُ", language: "ar", script: "Arab", bucket: "complexScriptFontFamily" },
  { id: "ar-ligature", fontId: "ar", text: "لا", language: "ar", script: "Arab", bucket: "complexScriptFontFamily" },
  { id: "he", text: "שָׁלוֹם", language: "he", script: "Hebr", bucket: "complexScriptFontFamily" },
  { id: "latin", text: "a\u0301", language: "en", script: "Latn", bucket: "fontFamily" },
];

export function createRealFontFreshFixture(createPresentation, families) {
  const document = createPresentation({ size: { width: 720, height: 260 }, theme: { fonts: { body: families.latin, heading: families.latin } } });
  const map = { Latn: families.latin, Hans: families.sc, Hant: families.tc, Jpan: families.ja, Hang: families.ko, Arab: families.ar, Hebr: families.he, Grek: families.greek };
  const controls = [];
  for (const sample of freshFontSamples) for (const mode of ["direct", "map"]) {
    const name = `${mode}-${sample.id}`, fontId = sample.fontId ?? sample.id, target = families[fontId];
    const style = { fontFamily: families.latin, eastAsianFontFamily: families.sc, complexScriptFontFamily: families.ar, fontSize: 48, language: sample.language, color: "000000", ...(mode === "map" ? { scriptFontFamilies: map } : { [sample.bucket]: target }) };
    const slide = document.addSlide({ id: name });
    slide.addElement({ id: `label-${name}`, type: "text", name: "Label", box: { x: 20, y: 20, width: 680, height: 40 }, content: [{ runs: [{ text: name, style: { fontFamily: families.latin, fontSize: 16 } }] }] });
    slide.addElement({ id: `probe-${name}`, type: "text", name: "Probe", box: { x: 20, y: 90, width: 680, height: 110 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: sample.text, style }] }] });
    const slots = mode === "map" ? [target, target, target] : [style.fontFamily, style.eastAsianFontFamily, style.complexScriptFontFamily];
    controls.push({ ...sample, id: name, sampleId: sample.id, mode, fontId, targetFamily: target, sourceStyle: style, expectedSlots: slots, slideId: name, elementId: `probe-${name}` });
  }
  return { document, controls, sourceMap: map };
}
