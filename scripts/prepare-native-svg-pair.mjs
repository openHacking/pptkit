import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAALElEQVR4nGNkYPjPQC5gIlsnw6hmkgET6VoQYFQziYCJVA3IYFQziYCiAAMAIJ4BJwS5UI4AAAAASUVORK5CYII=', 'base64');
const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20" fill="red"/></svg>');
const deck = createPresentation({ size: { width: 400, height: 200 } });
deck.registerAsset({ id: 'pair', kind: 'image', mimeType: 'image/svg+xml', source: { type: 'bytes', value: svg } });
deck.registerAsset({ id: 'control', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png } });
const slide = deck.addSlide({ id: 's' });
slide.addElement({ id: 'pair', name: 'SVG red / PNG blue', type: 'image', assetId: 'pair', box: { x: 10, y: 10, width: 180, height: 180 } });
slide.addElement({ id: 'control', name: 'PNG blue control', type: 'image', assetId: 'control', box: { x: 210, y: 10, width: 180, height: 180 } });
const result = await generatePptx(deck, { resolveAsset: asset => asset.id === 'pair' ? { bytes: png, mimeType: 'image/png', svgPair: svg } : { bytes: png, mimeType: 'image/png' } });
assert.deepEqual(result.warnings.map(w => w.code), ['asset-svg-pair-unverified']);
await mkdir('.tmp/native-svg-pair', { recursive: true });
await writeFile('.tmp/native-svg-pair/valid-source.pptx', result.bytes, { flag: 'wx' });
