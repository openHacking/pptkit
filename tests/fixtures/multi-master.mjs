export function multiMasterDocument(createPresentation) {
  const document = createPresentation({ id: 'masters', theme: { colors: { accent1: '112233' }, fonts: { body: 'Arial' } }, masters: [
    { id: 'red', name: 'Red & one', theme: { name: 'Red theme', colors: { accent1: 'CC2200', background1: 'F8E8E0' }, fonts: { body: 'Georgia' } }, background: { type: 'solid', color: { theme: 'background1' } } },
    { id: 'blue', name: 'Blue', theme: { name: 'Blue theme', colors: { accent1: '0044CC', background1: 'E0E8F8' }, fonts: { body: 'Times New Roman' } }, background: { type: 'solid', color: { theme: 'background1' } } },
    { id: 'unused', name: 'Unused', theme: { colors: { accent1: '007733' } } },
  ] });
  for (const [index, masterId] of ['red', 'blue', 'red', undefined].entries()) {
    const id = 'layout-' + index;
    document.defineSlideLayout({ id, name: id, ...(masterId ? { masterId } : {}), placeholders: [{ key: 'same', kind: 'body', box: { x: 10, y: 20, width: 300, height: 60 } }] });
    const slide = document.addSlide({ id: 'page-' + index, layoutId: id, notes: 'Independent notes' });
    slide.addElement({ id: 'text-' + index, type: 'text', placeholderKey: 'same', content: [{ runs: [{ text: 'Owned ' + index, style: { color: { theme: 'accent1' }, fontFamily: { theme: 'body' } } }] }] });
  }
  return document;
}
