export const mapText = "A中عא";
export const mapStyle = { fontFamily: "Default", language: "zh-TW", scriptFontFamilies: { Latn: "Latin Map", Hant: "Traditional Map", Arab: "Arabic Map", Hebr: "Hebrew Map" }, color: "123456", colorOpacity: .5, fontSize: 18 };
export function createScriptFontMapFixture(createPresentation) {
  const p = createPresentation({ theme: { fonts: { heading: "Heading", body: "Body" } } });
  p.defineSlideLayout({ id: "map-layout", name: "Map layout", placeholders: [{ key: "title", kind: "title", box: { x: 0, y: 0, width: 300, height: 50 }, textStyle: { run: mapStyle } }] });
  const content = [{ runs: [{ text: mapText, style: mapStyle }] }], box = { x: 0, y: 0, width: 400, height: 60 };
  const s = p.addSlide({ id: "map-slide", notes: content });
  s.addElement({ id: "map-native", type: "text", box, frame: { wrap: false }, content });
  s.addElement({ id: "map-linked", type: "text", box: { ...box, y: 80 }, content: [{ runs: [{ text: mapText, style: mapStyle, action: { type: "url", url: "https://example.invalid/map" } }] }] });
  s.addElement({ id: "map-shape", type: "shape", shape: "rect", box: { ...box, y: 160 }, text: { content } });
  s.addElement({ id: "map-table", type: "table", box: { ...box, y: 240 }, columns: [400], style: { wholeTable: { text: { run: mapStyle } } }, rows: [{ cells: [{ content: mapText }] }] });
  return p;
}
