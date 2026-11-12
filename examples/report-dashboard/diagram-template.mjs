import { createPresentation, validatePresentation } from '@pptkit/core';
import { generatePptx } from '@pptkit/pptx-exporter';
import { renderPresentationToSvg } from '@pptkit/svg-renderer';
import { writeFile } from 'node:fs/promises';
import { pathToFileURL } from 'node:url';

/** Create a reusable layout and two pages sharing it. */
export function createDiagramTemplate() {
  const deck = createPresentation({ size: { width: 960, height: 540 }, theme: { colors: { accent1: '2457D6' } } });
  deck.defineSlideLayout({
    id: 'process-layout', name: 'Process',
    elements: [{ type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 12, height: 540 }, style: { fill: { type: 'solid', color: { theme: 'accent1' } } } }],
    placeholders: [{ key: 'title', kind: 'title', box: { x: 40, y: 30, width: 880, height: 54 }, textStyle: { run: { fontSize: 30, bold: true } } }],
  });
  const process = deck.addSlide({ id: 'process', layoutId: 'process-layout' });
  process.addElement({ type: 'text', placeholderKey: 'title', content: 'Delivery process' });
  process.addElement({ type: 'shape', id: 'plan', shape: 'rect', box: { x: 120, y: 210, width: 220, height: 100 }, text: { content: 'Plan' } });
  process.addElement({ type: 'shape', id: 'deliver', shape: 'rect', box: { x: 620, y: 210, width: 220, height: 100 }, text: { content: 'Deliver' } });
  process.addElement({ type: 'connector', id: 'flow', start: { elementId: 'plan', anchor: 'right' }, end: { elementId: 'deliver', anchor: 'left' }, style: { paint: { type: 'solid', color: '2457D6' }, width: 2, endArrow: 'triangle' } });
  const next = deck.addSlide({ id: 'next', layoutId: 'process-layout' });
  next.addElement({ type: 'text', placeholderKey: 'title', content: 'Next steps' });
  next.addElement({ type: 'text', content: 'Assign owners and review progress.', box: { x: 40, y: 140, width: 850, height: 60 } });
  return deck;
}

export async function renderDiagramTemplate() {
  const deck = createDiagramTemplate();
  const diagnostics = validatePresentation(deck);
  if (diagnostics.some(item => item.severity === 'error')) throw new Error(JSON.stringify(diagnostics));
  const [pptx, preview] = await Promise.all([generatePptx(deck, { connectorMode: 'native' }), renderPresentationToSvg(deck)]);
  return { diagnostics, pptx, preview };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const { diagnostics, pptx, preview } = await renderDiagramTemplate();
  await writeFile('diagram-template.pptx', pptx.bytes, { flag: 'wx' });
  console.log(JSON.stringify({ diagnostics, exportWarnings: pptx.warnings, previewWarnings: preview.warnings }, null, 2));
}
