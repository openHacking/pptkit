export function chartNotesOriginsDocument(createPresentation) {
  const d = createPresentation({ id: 'chart-notes-origins', theme: { fonts: { body: 'Arial' } }, masters: [
    { id: 'red', name: 'Red', theme: { fonts: { body: 'Georgia' }, colors: { accent1: 'CC2200' } } },
    { id: 'blue', name: 'Blue', theme: { fonts: { body: 'Verdana' }, colors: { accent1: '0044CC' } } },
  ], textStylePresets: { unused: { run: { bold: true, fontSize: 48 } } } });
  for (const [index, masterId] of ['red', 'blue', 'red'].entries()) {
    d.defineSlideLayout({ id: 'layout-' + index, name: masterId + index, masterId, placeholders: [{ key: 'body', kind: 'body', box: { x: 10, y: 10, width: 300, height: 60 }, textStyle: { run: { bold: true, fontSize: 48 } } }] });
    const slide = d.addSlide({ id: 'page-' + index, layoutId: 'layout-' + index, notes: [{ style: { bullet: { type: 'number' }, indent: 0, spaceAfter: 0 }, runs: [
      { text: '', style: { bold: false, colorOpacity: 0, eastAsianFontFamily: null, scriptFontFamilies: null } },
      { text: 'Notes ' + index, style: { fontFamily: 'Tahoma', italic: true, fontSize: 11, scriptFontFamilies: { Hans: 'Arial', Arab: null } } },
    ] }] });
    for (const chartType of ['bar', 'line', 'pie']) slide.addElement({ id: chartType + index, type: 'chart', chartType, box: { x: 10 + ['bar', 'line', 'pie'].indexOf(chartType) * 310, y: 100, width: 300, height: 230 }, categories: ['A', 'B'], series: [{ name: 'Series', values: [2, 3] }], title: chartType,
      style: { textStyle: { fontFamily: { theme: 'body' }, color: { theme: 'accent1' }, fontSize: 14, bold: true, italic: true } },
      titleStyle: { bold: false, color: '#112233' }, legend: { visible: false, textStyle: { italic: false } },
      ...(chartType === 'pie' ? {} : { axes: { category: { labelStyle: { fontSize: 9, bold: false } }, value: { labelStyle: { fontFamily: 'Arial', color: '595959' } } } }),
    });
  }
  const s = d.addSlide({ id: 'empty-notes' });
  s.addElement({ id: 'default-chart', type: 'chart', chartType: 'pie', box: { x: 10, y: 10, width: 300, height: 220 }, categories: ['A'], series: [{ name: 'Only', values: [1] }] });
  s.addElement({ id: 'chart-group', type: 'group', box: { x: 350, y: 10, width: 300, height: 220 }, coordinateSize: { width: 300, height: 220 }, children: [{ id: 'nested-chart', type: 'chart', chartType: 'bar', box: { x: 0, y: 0, width: 300, height: 220 }, categories: ['A'], series: [{ name: 'Only', values: [1] }], titleStyle: { fontSize: 12, bold: true } }] });
  return d;
}
