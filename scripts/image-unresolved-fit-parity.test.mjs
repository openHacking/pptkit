import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { createPresentation } from '../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABQAAAAKCAIAAAA7N+mxAAAAGElEQVR4nGP8z0A+YKJAL8OoZhLBEA0wAHjzAROsNY7ZAAAAAElFTkSuQmCC', 'base64');

test('unresolved contain uses browser intrinsic aspect while PPTX reports stretch fallback', async t => {
  const deck = createPresentation({ size: { width: 300, height: 100 } });
  deck.registerAsset({ id: 'unknown', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png } });
  deck.registerAsset({ id: 'known', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png }, width: 20, height: 10 });
  const slide = deck.addSlide({ id: 's' });
  for (const [id, assetId, fit, x] of [['contain', 'unknown', 'contain', 0], ['cover', 'unknown', 'cover', 100], ['resolved', 'known', 'contain', 200]]) slide.addElement({ id, type: 'image', assetId, fit, box: { x, y: 0, width: 100, height: 100 } });
  const preview = await renderPresentationToSvg(deck);
  assert.match(preview.slides[0].svg, /preserveAspectRatio="xMidYMid meet"/);
  assert.match(preview.slides[0].svg, /preserveAspectRatio="xMidYMid slice"/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const pixels = await page.evaluate(async svg => {
    const image = new Image(), url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
    try { image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 300; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0); return [[50,10],[50,50],[150,10],[250,10],[250,50]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); }
    finally { URL.revokeObjectURL(url); }
  }, preview.slides[0].svg);
  assert.deepEqual(pixels, [[255,255,255,255],[255,0,0,255],[255,0,0,255],[255,255,255,255],[255,0,0,255]]);
  const result = await generatePptx(deck), warnings = result.warnings.filter(w => w.code === 'image-fit-unresolved');
  assert.deepEqual(warnings.map(w => [w.slideId, w.elementId, w.assetId]), [['s','contain','unknown'], ['s','cover','unknown']]);
  assert.ok(warnings.every(w => w.message.includes('stretches')));
  const xml = readGeneratedZip(result.bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/<a:stretch><a:fillRect\/><\/a:stretch>/g) ?? []).length, 3);
});

test('unresolved export warnings identify shared and nested owners and skip omitted media', async () => {
  const deck = createPresentation();
  deck.registerAsset({ id: 'image', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png } });
  const image = id => ({ id, type: 'image', assetId: 'image', fit: 'contain', box: { x: 0, y: 0, width: 20, height: 20 } });
  const group = id => ({ id, type: 'group', box: { x: 0, y: 0, width: 20, height: 20 }, coordinateSize: { width: 20, height: 20 }, children: [image(`${id}-image`)] });
  deck.defineSlideMaster({ id: 'm', name: 'Master', elements: [group('master')] });
  deck.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', elements: [group('layout')] });
  deck.addSlide({ id: 's', layoutId: 'l', elements: [group('slide')] });
  const result = await generatePptx(deck);
  const warnings = result.warnings.filter(w => w.code === 'image-fit-unresolved');
  assert.deepEqual(warnings.map(w => [w.elementId, w.slideId, w.layoutId, w.masterId, w.assetId]), [
    ['layout-image', undefined, 'l', undefined, 'image'], ['slide-image', 's', undefined, undefined, 'image'], ['master-image', undefined, undefined, 'm', 'image'],
  ]);
  const omitted = await generatePptx(deck, { resolveAsset: () => undefined });
  assert.equal(omitted.warnings.filter(w => w.code === 'image-fit-unresolved').length, 0);
  assert.equal(omitted.warnings.filter(w => w.code === 'image-omitted').length, 3);
});

test('fractional crop uses the same pixel region with or without declared image dimensions', async t => {
  const twoColors = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABQAAAAKCAIAAAA7N+mxAAAAHUlEQVR4nGP8z4APMDLgk2dioAAwjWomDQzRAAMAuhMCE3kFd/YAAAAASUVORK5CYII=', 'base64');
  const deck = createPresentation({ size: { width: 200, height: 100 } });
  for (const [id, dimensions] of [['unknown', {}], ['known', { width: 20, height: 10 }]]) deck.registerAsset({ id, kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: twoColors }, ...dimensions });
  const slide = deck.addSlide({ id: 's' });
  for (const [id, x] of [['unknown', 0], ['known', 100]]) slide.addElement({ id, type: 'image', assetId: id, fit: 'crop', crop: { left: .5 }, box: { x, y: 0, width: 100, height: 100 } });
  const result = await renderPresentationToSvg(deck);
  assert.ok(!result.warnings.some(w => w.code === 'image-crop-degraded'));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const pixels = await page.evaluate(async svg => {
    const image = new Image(), url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
    try { image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 200; canvas.height = 100; const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0); return [[20,50],[80,50],[120,50],[180,50]].map(([x,y]) => [...ctx.getImageData(x,y,1,1).data]); }
    finally { URL.revokeObjectURL(url); }
  }, result.slides[0].svg);
  assert.deepEqual(pixels, Array(4).fill([0,0,255,255]));
  const pptx = await generatePptx(deck), xml = readGeneratedZip(pptx.bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/<a:srcRect l="50000" t="0" r="0" b="0"\/>/g) ?? []).length, 2);
});

test('very narrow fractional crops are not clamped and unrepresentable SVG geometry rejects', async () => {
  const make = width => { const deck = createPresentation({ size: { width: 200, height: 100 } }); deck.registerAsset({ id: 'image', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png } }); deck.addSlide().addElement({ id: 'narrow', type: 'image', assetId: 'image', fit: 'crop', crop: { left: .99999 }, box: { x: 0, y: 0, width, height: 10 } }); return deck; };
  const svg = (await renderPresentationToSvg(make(100))).slides[0].svg;
  const match = svg.match(/<image[^>]* width="([^"]+)"/); assert.ok(match); assert.ok(Number(match[1]) > 9_999_999);
  await assert.rejects(renderPresentationToSvg(make(1e308)), /cannot be represented as finite SVG coordinates/);
});

test('PPTX warns when crop rounding removes a visible axis', async () => {
  const deck = createPresentation(); deck.registerAsset({ id: 'image', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png } });
  const slide = deck.addSlide({ id: 's' });
  for (const [id, crop] of [['safe', { left: .99999 }], ['horizontal', { left: .499996, right: .499996 }], ['vertical', { bottom: .999996 }]]) slide.addElement({ id, type: 'image', assetId: 'image', fit: 'crop', crop, box: { x: 0, y: 0, width: 10, height: 10 } });
  const result = await generatePptx(deck);
  assert.deepEqual(result.warnings.filter(w => w.code === 'image-crop-quantized-empty').map(w => [w.slideId,w.elementId,w.assetId]), [['s','horizontal','image'], ['s','vertical','image']]);
  const xml = readGeneratedZip(result.bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /l="50000" t="0" r="50000" b="0"/); assert.match(xml, /b="100000"/);
});

test('unrepresentable image aspect ratios remain unresolved and report the PPTX fallback', async () => {
  const deck=createPresentation();
  for(const [id,width,height] of [['overflow',1e308,1e-308],['rounded',1e20,1]]) deck.registerAsset({id,kind:'image',mimeType:'image/png',source:{type:'bytes',value:png},width,height});
  const slide=deck.addSlide({id:'s'});
  slide.addElement({id:'overflow',type:'image',assetId:'overflow',fit:'contain',box:{x:0,y:0,width:20,height:20}});
  slide.addElement({id:'rounded',type:'image',assetId:'rounded',fit:'cover',box:{x:30,y:0,width:20,height:20}});
  const result=await generatePptx(deck),warnings=result.warnings.filter(w=>w.code==='image-fit-unresolved');
  assert.deepEqual(warnings.map(w=>w.elementId),['overflow','rounded']);
  const xml=readGeneratedZip(result.bytes).get('ppt/slides/slide1.xml').toString();
  assert.doesNotMatch(xml,/NaN|Infinity|<a:srcRect/);
});
