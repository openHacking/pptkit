/** No source patching: exercise the empty-paragraph insertion style contract. */
export function notesInsertionStyles(createPresentation) {
  const cases = [
    { id: "regular", runs: [{ text: "", style: { fontFamily: "Georgia", fontSize: 14, language: "fr-FR", color: "123456" } }] },
    { id: "bold-italic", runs: [{ text: "", style: { fontFamily: "Arial", fontSize: 12, bold: true, italic: true, language: "en-US" } }] },
    { id: "script-map", runs: [{ text: "", style: { fontFamily: "Arial", fontSize: 18, language: "en-US", scriptFontFamilies: { Latn: "Georgia", Hant: "PingFang TC" } } }] },
    { id: "two-empty", runs: [{ text: "", style: { fontFamily: "Arial", fontSize: 24, bold: true } }, { text: "", style: { fontFamily: "Georgia", fontSize: 16, italic: true, language: "en-US" } }] },
  ];
  const document = createPresentation({ size: { width: 720, height: 260 }, theme: { fonts: { body: "Arial", heading: "Arial" } } });
  for (const c of cases) {
    const slide = document.addSlide({ id: c.id, notes: [{ runs: c.runs }] });
    slide.addElement({ type: "text", name: "Label", box: { x: 36, y: 60, width: 648, height: 100 }, content: [{ runs: [{ text: "Insertion " + c.id, style: { fontFamily: "Arial", fontSize: 32 } }] }] });
  }
  return { document, cases };
}
