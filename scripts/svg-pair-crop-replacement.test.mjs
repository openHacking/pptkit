import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createPresentation, planImageReplacement } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20" fill="red"/></svg>');
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAALElEQVR4nGNkYPjPQC5gIlsnw6hmkgET6VoQYFQziYCJVA3IYFQziYCiAAMAIJ4BJwS5UI4AAAAASUVORK5CYII=', 'base64');
const resolver = asset => asset.id === 'pair' ? { bytes: png, mimeType: 'image/png', svgPair: svg } : { bytes: Buffer.from(asset.source.value, 'base64'), mimeType: 'image/png' };

test('cropped SVG pair replacement isolates a shared picture and preserves the original on undo', async () => {
  const replacementBytes = await readFile('tests/fixtures/transparent-replacement.png');
  const deck = createPresentation({ size: { width: 400, height: 200 } });
  deck.registerAsset({ id: 'pair', kind: 'image', mimeType: 'image/svg+xml', source: { type: 'bytes', value: svg } });
  const replacement = deck.registerAsset({ id: 'replacement', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: replacementBytes }, width: 4, height: 2 });
  const slide = deck.addSlide({ id: 's' });
  const image = (id, x) => ({ id, type: 'image', assetId: 'pair', fit: 'crop', box: { x, y: 10, width: 180, height: 180 }, crop: { left: .1, top: .15, right: .2, bottom: .05 }, transform: { rotation: 30, flipH: true } });
  slide.addElement(image('target', 10)); slide.addElement(image('shared', 210));
  const plan = planImageReplacement(slide.elements[0], replacement, { fit: 'preserve' });
  slide.replaceElement('target', plan.replacement);
  const output = await generatePptx(deck, { resolveAsset: resolver });
  assert.ok(output.warnings.some(w => w.code === 'asset-svg-pair-unverified' && w.assetId === 'pair'));
  const parts = readGeneratedZip(output.bytes), xml = parts.get('ppt/slides/slide1.xml').toString(), rels = parts.get('ppt/slides/_rels/slide1.xml.rels').toString();
  assert.equal((xml.match(/<a:srcRect l="10000" t="15000" r="20000" b="5000"\/>/g) ?? []).length, 2);
  assert.equal((xml.match(/rot="1800000" flipH="1"/g) ?? []).length, 2);
  assert.equal((xml.match(/<asvg:svgBlip/g) ?? []).length, 1);
  assert.match(rels, /Target="\.\.\/media\/pptkit-svg-pairs\/1\.svg"/);
  assert.deepEqual(parts.get('ppt/media/pptkit-svg-pairs/1.svg'), svg);
  assert.deepEqual(parts.get('ppt/media/pair.png'), png);
  assert.deepEqual(parts.get('ppt/media/replacement.png'), replacementBytes);
  const session = await importPptx(output.bytes, { projection: 'content' });
  try {
    const images = session.document.slides[0].elements;
    assert.deepEqual(images.map(e => e.crop), [image('x', 10).crop, image('x', 10).crop]);
    assert.deepEqual(images.map(e => Buffer.from(session.document.assets.find(a => a.id === e.assetId).source.value, 'base64')), [replacementBytes, png]);
    assert.deepEqual(images.map(e => e.transform), [{ rotation: 30, flipH: true, flipV: false }, { rotation: 30, flipH: true, flipV: false }]);
    assert.ok(session.report.findings.some(f => f.operation === 'preview' && f.elementId === images[1].id && f.reason.includes('SVG image extension')));
    assert.ok(!session.report.findings.some(f => f.elementId === images[0].id && f.reason.includes('SVG image extension')));
  } finally { session.dispose(); }
  slide.replaceElement('target', plan.previous);
  assert.equal(slide.elements[0].assetId, 'pair'); assert.equal(slide.elements[1].assetId, 'pair');
  const undo = readGeneratedZip((await generatePptx(deck, { resolveAsset: resolver })).bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((undo.match(/<asvg:svgBlip/g) ?? []).length, 2);
});
