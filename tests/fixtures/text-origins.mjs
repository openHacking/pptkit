export function textOriginsDocument(createPresentation) {
  const d = createPresentation({ id: 'text-origins', theme: { fonts: { body: 'Arial', heading: 'Arial' } }, textStylePresets: {
    'base.with.dot': { frame: { margin: 9, wrap: true, autoFit: { mode: 'shrink', fontScale: 0.7 } }, paragraph: { spaceBefore: 12 }, run: { bold: true, italic: true, underline: true, strike: true, fontSize: 24, eastAsianFontFamily: 'Arial', complexScriptFontFamily: 'Tahoma', scriptFontFamilies: { Hans: 'Arial', Arab: 'Tahoma' } } },
    'element"preset': { frame: { verticalAlign: 'middle' }, run: { italic: false, scriptFontFamilies: { Hans: 'Verdana' } } },
  } });
  const box = { x: 10, y: 20, width: 300, height: 80 };
  d.defineSlideLayout({ id: 'a', name: 'A', placeholders: [{ key: 'same', kind: 'body', box, textStylePreset: 'base.with.dot', textStyle: { frame: { margin: { left: 3 } }, run: { color: '123456', scriptFontFamilies: null, eastAsianFontFamily: null } } }], elements: [{ id: 'shared-text', type: 'text', placeholderKey: 'same', box: { x: 10, y: 300, width: 300, height: 60 }, content: 'Shared' }] });
  d.defineSlideLayout({ id: 'b', name: 'B', placeholders: [{ key: 'same', kind: 'body', box, textStyle: { run: { color: 'ABCDEF' } } }] });
  d.addSlide({ id: 'first', layoutId: 'a', elements: [
    { id: 'body', type: 'text', placeholderKey: 'same', textStylePreset: 'element"preset', frame: { margin: 0, wrap: false, autoFit: { mode: 'shrink' } }, content: [{ style: { bullet: { type: 'number' }, indent: 0, spaceBefore: 0 }, runs: [{ text: '', style: { bold: false, italic: false, underline: false, strike: false, colorOpacity: 0, complexScriptFontFamily: null, scriptFontFamilies: { Hans: null } } }, { text: 'Next' }, { text: ' Plain', style: { bold: false, italic: false, underline: false, strike: false } }] }] },
    { id: 'shape', type: 'shape', shape: 'rect', placeholderKey: 'same', box: { x: 10, y: 140, width: 300, height: 80 }, text: { content: 'Shape' } },
    { id: 'group', type: 'group', placeholderKey: 'same', box: { x: 350, y: 20, width: 300, height: 80 }, coordinateSize: { width: 300, height: 80 }, children: [{ id: 'child', type: 'text', box, frame: { wrap: false }, content: '' }] },
  ] });
  d.addSlide({ id: 'second', layoutId: 'b', elements: [{ id: 'other', type: 'text', placeholderKey: 'same', content: 'Other' }] });
  return d;
}
