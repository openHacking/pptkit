import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';

test('large gradient-warning batches preserve ordering without argument-count failure', async () => {
  const document = createPresentation();
  const stops = Array.from({ length: 256 }, (_, index) => ({ position: index * 1e-8, color: 'FF0000' }));
  const slide = document.addSlide({ id: 'slide' });
  for (let index = 0; index < 600; index++) slide.addElement({
    id: `shape-${index}`, type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 2, height: 2 },
    style: { fill: { type: 'linearGradient', angle: 0, stops } },
  });
  const result = await generatePptx(document);
  assert.equal(result.status, 'generated-with-warnings');
  assert.equal(result.warnings.length, 600 * 255);
  assert.equal(result.warnings[0].code, 'shape-gradient-stops-collapsed');
  assert.equal(result.warnings[0].elementId, 'shape-0');
  assert.equal(result.warnings.at(-1).elementId, 'shape-599');
  assert.ok(result.bytes.length > 0);
});

test('large layout diagnostic batches reach the export report in order', async () => {
  const document = createPresentation(); document.addSlide({ id: 'slide' });
  const missing = { type: 'align', container: { kind: 'slide', id: 'missing' }, elementIds: [], horizontal: 'start' };
  const count = 150000;
  const result = await generatePptx(document, { layout: { composition: Array(count).fill(missing) } });
  assert.equal(result.warnings.length, count);
  assert.equal(result.warnings[0].code, 'composition-container-missing');
  assert.equal(result.warnings[0].path, 'options.composition.0');
  assert.equal(result.warnings[count - 1].path, `options.composition.${count - 1}`);
  assert.equal(result.status, 'generated-with-warnings');
});
