import { createPresentation, validatePresentation } from '@pptkit/core';
import { generatePptx } from '@pptkit/pptx-exporter';
import { renderPresentationToSvg } from '@pptkit/svg-renderer';
import { writeFile } from 'node:fs/promises';
import { pathToFileURL } from 'node:url';

/** Build a fresh presentation. Inputs are illustrative; replace them with your own measured data. */
export function createDashboard() {
  const deck = createPresentation({ metadata: { title: 'Quarterly operations report' }, size: { width: 960, height: 540 } });
  const slide = deck.addSlide({ id: 'overview' });
  slide.addElement({ type: 'text', content: [{ runs: [{ text: 'Quarterly operations report', style: { fontSize: 30, bold: true } }] }], box: { x: 40, y: 24, width: 880, height: 50 } });
  slide.addElement({ type: 'text', content: [{ runs: [{ text: 'Orders by quarter', style: { fontSize: 20, bold: true } }] }], box: { x: 40, y: 92, width: 430, height: 30 } });
  slide.addElement({ type: 'chart', id: 'orders', chartType: 'bar', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Orders', values: [110, 125, 141, 158], color: '2457D6' }], box: { x: 40, y: 130, width: 430, height: 300 }, legend: { visible: false } });
  slide.addElement({ type: 'text', content: [{ runs: [{ text: 'Operating priorities', style: { fontSize: 20, bold: true } }] }], box: { x: 510, y: 92, width: 410, height: 30 } });
  slide.addElement({ type: 'table', id: 'priorities', box: { x: 510, y: 130, width: 410, height: 180 }, columns: [210, 200], rows: [
    { cells: [{ content: 'Priority' }, { content: 'Status' }] },
    { cells: [{ content: 'Delivery' }, { content: 'On track' }] },
    { cells: [{ content: 'Retention' }, { content: 'In progress' }] },
  ] });
  return deck;
}

/** Produce detached bytes and preview without writing or starting an application. */
export async function renderDashboard() {
  const deck = createDashboard();
  const diagnostics = validatePresentation(deck);
  if (diagnostics.some(item => item.severity === 'error')) throw new Error(JSON.stringify(diagnostics));
  const [pptx, preview] = await Promise.all([generatePptx(deck), renderPresentationToSvg(deck)]);
  return { diagnostics, pptx, preview };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const { diagnostics, pptx, preview } = await renderDashboard();
  await writeFile('report-dashboard.pptx', pptx.bytes, { flag: 'wx' });
  const firstSlide = preview.slides[0];
  if (!firstSlide) throw new Error('Expected a rendered slide.');
  await writeFile('report-dashboard.svg', firstSlide.svg, { flag: 'wx' });
  console.log(JSON.stringify({ diagnostics, exportWarnings: pptx.warnings, previewWarnings: preview.warnings }, null, 2));
}
