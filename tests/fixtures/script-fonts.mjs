export const scriptText = "Latin中文العربية";
export const scriptStyle = { fontFamily: "Latin Face", eastAsianFontFamily: "Eastern Face", complexScriptFontFamily: "Complex Face", fontSize: 18, color: "123456" };
export function createScriptFontFixture(createPresentation) {
  const p = createPresentation({ size: { width: 600, height: 420 }, theme: { fonts: { heading: "Heading Face", body: "Body Face" } }, textStylePresets: { multi: { run: scriptStyle } } });
  const content = [{ runs: [{ text: scriptText, style: scriptStyle }] }];
  const s = p.addSlide({ id: "script-page", notes: content });
  s.addElement({ id: "native", name: "Script native", accessibility: { description: "Script native" }, type: "text", box: { x: 20, y: 20, width: 550, height: 55 }, frame: { margin: 0, wrap: false }, content });
  s.addElement({ id: "linked", name: "Script linked", type: "text", box: { x: 20, y: 90, width: 550, height: 55 }, frame: { margin: 0 }, content: [{ runs: [{ text: scriptText, style: scriptStyle, action: { type: "hyperlink", url: "https://example.invalid/script" } }] }] });
  s.addElement({ id: "shape", type: "shape", shape: "rect", box: { x: 20, y: 160, width: 550, height: 55 }, text: { textStylePreset: "multi", content: scriptText } });
  s.addElement({ id: "table", type: "table", box: { x: 20, y: 230, width: 550, height: 70 }, columns: [275, 275], style: { wholeTable: { text: { run: scriptStyle } } }, rows: [{ height: 70, cells: [{ content: scriptText }, { content: [{ runs: [{ text: scriptText, style: { eastAsianFontFamily: { theme: "heading" }, complexScriptFontFamily: null } }] }] }] }] });
  return p;
}
