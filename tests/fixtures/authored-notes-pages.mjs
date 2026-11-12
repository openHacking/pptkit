/** Public authoring only: no source XML patching or exporter internals. */
export function authoredNotesPages(createPresentation, size, name) {
  const document = createPresentation({ size, theme: { fonts: { body: "Arial", heading: "Arial" } } });
  const first = document.addSlide({ id: "first", notes: [{ runs: [{ text: `Fresh notes ${name} ONE`, style: { fontFamily: "Arial", fontSize: 18 } }] }] });
  first.addElement({ type: "text", name: "Identity", box: { x: size.width * .1, y: size.height * .2, width: size.width * .8, height: size.height * .3 }, content: [{ runs: [{ text: `Fresh ${name} ONE`, style: { fontFamily: "Arial", fontSize: 24 } }] }] });
  const copy = document.duplicateSlide(first.id);
  copy.setNotes([{ runs: [{ text: `Fresh notes ${name} COPY`, style: { fontFamily: "Arial", fontSize: 18 } }] }]);
  const label = structuredClone(copy.elements[0]);
  label.content[0].runs[0].text = `Fresh ${name} COPY`;
  copy.replaceElement(label.id, label);
  for (const [id, notes] of [["EMPTY", undefined], ["STYLED", [{ runs: [{ text: "", style: { fontSize: 12, bold: true } }] }]]]) {
    const slide = document.addSlide({ id, ...(notes ? { notes } : {}) });
    slide.addElement({ type: "text", name: "Identity", box: { x: size.width * .1, y: size.height * .2, width: size.width * .8, height: size.height * .3 }, content: [{ runs: [{ text: `Fresh ${name} ${id}`, style: { fontFamily: "Arial", fontSize: 24 } }] }] });
  }
  return document;
}
