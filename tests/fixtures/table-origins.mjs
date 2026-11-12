export function tableOriginsDocument(createPresentation) {
  const d = createPresentation({ id: 'table-origins', textStylePresets: {
    'cell."preset': { frame: { margin: { left: 0 }, verticalAlign: 'middle' }, paragraph: { spaceBefore: 0, bullet: { type: 'none' } }, run: { bold: false, eastAsianFontFamily: null, scriptFontFamilies: { Hans: 'Verdana' } } },
  }, masters: [{ id: 'red', name: 'Red', theme: { colors: { accent1: 'CC2200' }, fonts: { body: 'Georgia' } } }, { id: 'blue', name: 'Blue', theme: { colors: { accent1: '0044CC' }, fonts: { body: 'Verdana' } } }] });
  const style = {
    wholeTable: { fill: { type: 'solid', color: 'EEEEEE', opacity: 0.2 }, stroke: { paint: { type: 'solid', color: '111111', opacity: 0.4 }, width: 1, dash: 'dot' }, margin: 8, text: { paragraph: { align: 'right', spaceBefore: 8, bullet: { type: 'number', startAt: 7 } }, run: { fontFamily: { theme: 'body' }, color: { theme: 'accent1' }, fontSize: 13, bold: true, eastAsianFontFamily: 'Arial', scriptFontFamilies: { Hans: 'Arial', Arab: 'Tahoma' } } } },
    borders: { outer: { paint: { type: 'solid', color: 'FF0000' } }, innerHorizontal: { width: 3 }, innerVertical: { dash: 'dash' } },
    bandedColumns: { first: { text: { run: { fontSize: 14, italic: true } } }, second: { margin: { top: 2 }, text: { run: { italic: false } } } },
    bandedRows: { size: 1, first: { fill: { type: 'solid', color: 'CCCCCC' }, text: { run: { fontSize: 15, scriptFontFamilies: null } } }, second: { margin: { right: 2 }, text: { run: { underline: true } } } },
    firstColumn: { text: { paragraph: { align: 'center' } } }, lastColumn: { verticalAlign: 'bottom' },
    firstRow: { text: { paragraph: { bullet: { type: 'bullet' } } }, borders: { bottom: { dash: 'solid' } } }, lastRow: { fill: { type: 'solid', color: 'FFFFFF', opacity: 0 } },
    topLeft: { text: { run: { fontSize: 20 } } }, bottomRight: { text: { run: { fontSize: 22 } } },
  };
  const rows = [
    { height: 45, cells: [{ rowSpan: 2, colSpan: 2, textStylePreset: 'cell."preset', style: { fill: { type: 'none' }, margin: { right: 0 }, stroke: { width: 0 }, borders: { left: { dash: 'solid' } }, verticalAlign: 'bottom' }, content: [{ style: { indent: 0, bullet: { type: 'number' } }, runs: [{ text: '', style: { bold: false, colorOpacity: 0, scriptFontFamilies: { Arab: null } } }, { text: 'Merged', style: { italic: false, scriptFontFamilies: { Hans: null } } }] }] }, { content: 'Top' }, { content: [{ runs: [] }] }] },
    { height: 45, cells: [{ content: 'Right' }] },
    { height: 45, cells: [] },
    { height: 45, cells: [{ content: 'Bottom' }, { content: [] }, { content: 'Tail' }, { content: 'End' }] },
  ];
  for (const [index, masterId] of ['red', 'blue'].entries()) {
    d.defineSlideLayout({ id: 'layout-' + index, name: masterId, masterId });
    d.addSlide({ id: 'page-' + index, layoutId: 'layout-' + index, elements: [{ id: 'table-' + index, type: 'table', box: { x: 20, y: 20, width: 440, height: 180 }, columns: [110, 110, 110, 110], rows, style }] });
  }
  d.addSlide({ id: 'plain', elements: [{ id: 'group', type: 'group', box: { x: 20, y: 20, width: 440, height: 180 }, coordinateSize: { width: 440, height: 180 }, children: [{ id: 'plain-table', type: 'table', box: { x: 0, y: 0, width: 440, height: 180 }, columns: [220, 220], rows: [{ cells: [{ content: 'Preset', textStylePreset: 'cell."preset', style: { margin: { right: 0 }, stroke: { paint: { type: 'solid', color: '333333' }, width: 2 }, borders: { left: { dash: 'dash' } } } }, { content: [{ runs: [{ text: 'Local', style: { bold: false, eastAsianFontFamily: null, scriptFontFamilies: null } }] }] }] }] }] }] });
  return d;
}
