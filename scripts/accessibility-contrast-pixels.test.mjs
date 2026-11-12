import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { createPresentation, inspectSlideTextContrast } from '../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

function fixture(kind, color, opacity, background, masterTheme) {
  const colors = { background1: background, text1: color };
  const deck = createPresentation({ size: { width: 320, height: 160 }, theme: { colors: masterTheme ? { background1: 'FFFFFF', text1: '000000' } : colors }, ...(masterTheme ? { masters: [{ id: 'master', name: 'Master', theme: { colors } }] } : {}) });
  if (masterTheme) deck.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master' });
  const box = { x: 20, y: 20, width: 280, height: 120 };
  const content = [{ runs: [{ text: 'MMMM', style: { fontFamily: 'Arial', fontSize: 72, bold: true, color: { theme: 'text1' }, colorOpacity: opacity } }] }];
  const style = { fill: { type: 'solid', color: { theme: 'background1' } }, stroke: { paint: { type: 'none' }, width: 0 } };
  const element = kind === 'text' ? { id: 'object', type: 'text', box, content } : kind === 'shape' ? { id: 'object', type: 'shape', shape: 'rect', box, text: { content }, style } : { id: 'object', type: 'table', box, columns: [280], rows: [{ height: 120, cells: [{ content, style }] }] };
  deck.addSlide({ id: 'page', ...(masterTheme ? { layoutId: 'layout' } : {}), background: { type: 'solid', color: { theme: 'background1' } }, elements: [element] });
  return deck;
}

function contrast(foreground, background) {
  const luminance = rgb => { const linear = rgb.map(value => { const c = value / 255; return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; }); return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]; };
  const a = luminance(foreground), b = luminance(background);
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

test('Chromium rendered text pixels agree with claimed solid-background contrast', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 320, height: 160 }, deviceScaleFactor: 1 });
  const requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const kind of ['text', 'shape', 'table']) for (const [color, opacity, background, masterTheme] of [['000000', 1, 'FFFFFF', false], ['000000', 0.5, 'FFFFFF', false], ['777777', 1, 'FFFFFF', false], ['C02070', 0.5, 'FFFFFF', false], ['FFFFFF', 1, '002244', false], ['FFFFFF', 0.5, '002244', false], ['FFFFFF', 0.5, '002244', true]]) {
    const deck = fixture(kind, color, opacity, background, masterTheme), observation = inspectSlideTextContrast(deck).findings[0];
    assert.notEqual(observation.status, 'indeterminate');
    const rendered = await renderPresentationToSvg(deck);
    await page.setContent(`<style>html,body{margin:0}</style>${rendered.slides[0].svg}`);
    await page.evaluate(() => document.fonts.ready);
    const screenshot = await page.screenshot();
    const minimum = await page.evaluate(async base64 => {
      const img = new Image(); img.src = `data:image/png;base64,${base64}`; await img.decode();
      const canvas = document.createElement('canvas'); canvas.width = img.width; canvas.height = img.height;
      const ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0);
      const pixels = ctx.getImageData(0, 0, img.width, img.height).data;
      let max = 1, rgb = [], count = 0;
      const linear = value => { const c = value / 255; return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; };
      const bgLum = 0.2126 * linear(pixels[0]) + 0.7152 * linear(pixels[1]) + 0.0722 * linear(pixels[2]);
      for (let i = 0; i < pixels.length; i += 4) {
        const lum = 0.2126 * linear(pixels[i]) + 0.7152 * linear(pixels[i + 1]) + 0.0722 * linear(pixels[i + 2]);
        const ratio = (Math.max(lum, bgLum) + 0.05) / (Math.min(lum, bgLum) + 0.05);
        if (ratio > 1.05) { if (ratio > max) { max = ratio; rgb = [pixels[i], pixels[i + 1], pixels[i + 2]]; } count++; }
      }
      return { rgb, count, background: Array.from(pixels.slice(0, 4)) };
    }, screenshot.toString('base64'));
    assert.ok(minimum.count > 100, `${kind} ${color}/${opacity} rendered visible glyphs`);
    const expectedBackground = [0, 2, 4].map(index => parseInt(background.slice(index, index + 2), 16));
    assert.deepEqual(minimum.background, [...expectedBackground, 255]);
    assert.ok(Math.abs(contrast(minimum.rgb, expectedBackground) - observation.ratio) < 0.15, `${kind} ${color}/${opacity}: pixel ${minimum.rgb}, expected ${observation.ratio}`);
  }
  assert.deepEqual(requests, []);
});

test('opaque full-slide shape text uses its fill rather than the underlying slide background', async t => {
  const deck = createPresentation({ size: { width: 320, height: 160 } });
  deck.addSlide({ background: { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: 'FFFFFF' }, { position: 1, color: 'FF0000' }] }, elements: [{ type: 'shape', id: 'cover', shape: 'rect', box: { x: 0, y: 0, width: 320, height: 160 }, style: { fill: { type: 'solid', color: '002244' }, stroke: { paint: { type: 'none' }, width: 0 } }, text: { content: [{ runs: [{ text: 'MMMM', style: { fontFamily: 'Arial', fontSize: 72, bold: true, color: 'FFFFFF' } }] }] } }] });
  const finding = inspectSlideTextContrast(deck).findings[0]; assert.equal(finding.status, 'meets-threshold');
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 320, height: 160 }, deviceScaleFactor: 1 });
  await page.setContent(`<style>html,body{margin:0}</style>${(await renderPresentationToSvg(deck)).slides[0].svg}`); await page.evaluate(() => document.fonts.ready);
  const screenshot = await page.screenshot(), result = await page.evaluate(async base64 => {
    const image = new Image(); image.src = `data:image/png;base64,${base64}`; await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
    const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0); const pixels = ctx.getImageData(0, 0, image.width, image.height).data;
    let whitePixels = 0; for (let i = 0; i < pixels.length; i += 4) if (pixels[i] === 255 && pixels[i + 1] === 255 && pixels[i + 2] === 255) whitePixels++;
    return { corner: Array.from(pixels.slice(0, 4)), whitePixels };
  }, screenshot.toString('base64'));
  assert.deepEqual(result.corner, [0, 34, 68, 255]); assert.ok(result.whitePixels > 100);
  assert.ok(Math.abs(contrast([255, 255, 255], result.corner.slice(0, 3)) - finding.ratio) < 1e-10);
  const xml = readGeneratedZip((await generatePptx(deck)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:srgbClr val="002244"/); assert.match(xml, /<a:ext cx="4064000" cy="2032000"\/>/);
});

test('contrast rejects a mathematical cover whose exported EMU geometry leaves a gap', async () => {
  const width = 320.63942679845786, x = -0.025010755222666936;
  const deck = createPresentation({ size: { width, height: 160 } });
  deck.addSlide({ background: { type: 'solid', color: 'FFFFFF' }, elements: [{ type: 'shape', shape: 'rect', box: { x, y: 0, width: width - x, height: 160 }, style: { fill: { type: 'solid', color: '002244' }, stroke: { paint: { type: 'none' }, width: 0 } }, text: { content: 'Text' } }] });
  assert.equal(inspectSlideTextContrast(deck).findings[0].reason, 'unsupported-text-owner');
  const parts = readGeneratedZip((await generatePptx(deck)).bytes), main = parts.get('ppt/presentation.xml').toString(), slide = parts.get('ppt/slides/slide1.xml').toString();
  const slideWidth = Number(main.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/)[1]);
  const shapeTransform = slide.match(/<p:spPr><a:xfrm><a:off x="(-?\d+)" y="0"\/><a:ext cx="(\d+)" cy="2032000"\/>/);
  assert.ok(shapeTransform); const origin = Number(shapeTransform[1]), extent = Number(shapeTransform[2]);
  assert.equal(origin + extent, slideWidth - 1);
});

test('translucent full-slide fill is composited in Chromium and serialized at native alpha precision', async t => {
  const make = (opacity, fontSize = 72) => {
    const deck = createPresentation({ size: { width: 320, height: 160 } });
    deck.addSlide({ background: { type: 'solid', color: 'FFFFFF' }, elements: [{ type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 320, height: 160 }, style: { fill: { type: 'solid', color: '000000', opacity }, stroke: { paint: { type: 'none' }, width: 0 } }, text: { content: [{ runs: [{ text: 'MMMM', style: { fontFamily: 'Arial', fontSize, bold: true, color: '000000' } }] }] } }] });
    return deck;
  };
  const deck = make(0.5), finding = inspectSlideTextContrast(deck).findings[0]; assert.equal(finding.status, 'meets-threshold');
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 320, height: 160 }, deviceScaleFactor: 1 });
  await page.setContent(`<style>html,body{margin:0}</style>${(await renderPresentationToSvg(deck)).slides[0].svg}`); await page.evaluate(() => document.fonts.ready);
  const imageBytes = await page.screenshot(), pixels = await page.evaluate(async base64 => {
    const image = new Image(); image.src = `data:image/png;base64,${base64}`; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0);
    const data = ctx.getImageData(0, 0, image.width, image.height).data; let black = 0;
    for (let i = 0; i < data.length; i += 4) if (data[i] === 0 && data[i + 1] === 0 && data[i + 2] === 0) black++;
    return { background: Array.from(data.slice(0, 3)), black };
  }, imageBytes.toString('base64'));
  assert.ok(pixels.black > 100); assert.ok(pixels.background.every(channel => channel >= 127 && channel <= 128));
  assert.ok(Math.abs(contrast([0, 0, 0], pixels.background) - finding.ratio) < 0.05);
  const borderline = make(0.5446693, 12), result = inspectSlideTextContrast(borderline).findings[0]; assert.equal(result.reason, 'output-opacity-dependent');
  const xml = readGeneratedZip((await generatePptx(borderline)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:solidFill><a:srgbClr val="000000"><a:alpha val="54467"\/><\/a:srgbClr><\/a:solidFill>/);
});

test('nominal font sizes that differ across SVG and PPTX remain indeterminate', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const [size, bold] of [[18, false], [14, true]]) {
    const deck = createPresentation({ size: { width: 320, height: 160 } });
    deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 250, height: 70 }, content: [{ runs: [{ text: 'Example', style: { color: '777777', fontSize: size, bold } }] }] }] });
    const finding = inspectSlideTextContrast(deck).findings[0];
    assert.equal(finding.status, 'indeterminate'); assert.equal(finding.reason, 'output-size-dependent');
    assert.equal(finding.nativeThreshold, 3); assert.equal(finding.nominalSvgThreshold, 4.5);
    await page.setContent((await renderPresentationToSvg(deck)).slides[0].svg);
    const actualSize = await page.locator('text tspan').first().evaluate(node => getComputedStyle(node).fontSize);
    assert.equal(actualSize, `${size}px`);
    const xml = readGeneratedZip((await generatePptx(deck)).bytes).get('ppt/slides/slide1.xml').toString();
    assert.match(xml, new RegExp(`\\bsz="${size * 100}"`));
  }
});

test('borderline text alpha is emitted at different nominal precision in SVG and PPTX', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage();
  const deck = createPresentation({ size: { width: 320, height: 160 } });
  deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 250, height: 70 }, content: [{ runs: [{ text: 'Example', style: { color: '000000', colorOpacity: 0.534681, fontSize: 12 } }] }] }] });
  const finding = inspectSlideTextContrast(deck).findings[0];
  assert.equal(finding.reason, 'output-opacity-dependent');
  await page.setContent((await renderPresentationToSvg(deck)).slides[0].svg);
  const actualAlpha = await page.locator('text tspan').first().evaluate(node => getComputedStyle(node).fillOpacity);
  assert.equal(actualAlpha, '0.534681');
  const xml = readGeneratedZip((await generatePptx(deck)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:alpha val="53468"\/>/);
});

test('contrast checks inherited artwork using the same explicit visibility gates as SVG', async () => {
  const box = { x: 0, y: 0, width: 200, height: 100 };
  const art = (id, color) => ({ id, type: 'shape', shape: 'rect', box, style: { fill: { type: 'solid', color } } });
  const deck = createPresentation({ size: { width: 320, height: 160 }, masters: [{ id: 'master', name: 'Master', elements: [art('master-art', 'FF0000')] }] });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', showMasterShapes: false, elements: [art('layout-art', '0000FF')] });
  const label = id => ({ id, type: 'text', box, content: [{ runs: [{ text: 'Black', style: { color: '000000', fontSize: 12 } }] }] });
  deck.addSlide({ id: 'visible', layoutId: 'layout', elements: [label('one')] });
  deck.addSlide({ id: 'hidden', layoutId: 'layout', showInheritedShapes: false, elements: [label('two')] });
  const findings = inspectSlideTextContrast(deck).findings;
  assert.equal(findings[0].reason, 'complex-scene'); assert.equal(findings[1].status, 'meets-threshold');
  const slides = (await renderPresentationToSvg(deck)).slides;
  assert.match(slides[0].svg, /data-pptkit-element-id="layout-art"/);
  assert.doesNotMatch(slides[0].svg, /data-pptkit-element-id="master-art"/);
  assert.doesNotMatch(slides[1].svg, /data-pptkit-element-id="(?:layout|master)-art"/);
  const parts = readGeneratedZip((await generatePptx(deck)).bytes);
  assert.match(parts.get('ppt/slides/slide2.xml').toString(), /<p:sld\b[^>]*\bshowMasterSp="0"/);
  assert.ok([...parts.entries()].some(([name, bytes]) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name) && /<p:sldLayout\b[^>]*\bshowMasterSp="0"/.test(bytes.toString())));
});

test('inherited and local translucent layers match Chromium foreground and background pixels', async t => {
  const full = { x: 0, y: 0, width: 320, height: 160 };
  const layer = (id, color, opacity = 1) => ({
    id, type: 'shape', shape: 'rect', box: full,
    style: { fill: { type: 'solid', color, opacity }, stroke: { paint: { type: 'none' } } },
  });
  const deck = createPresentation({ size: full, masters: [{
    id: 'm', name: 'Master', theme: { colors: { accent1: '000000' } }, elements: [layer('base', { theme: 'accent1' })],
  }] });
  deck.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', elements: [layer('middle', 'FFFFFF', 0.5)] });
  deck.addSlide({ layoutId: 'l', elements: [layer('overlay', '000000', 0.5), {
    type: 'text', box: { x: 20, y: 10, width: 270, height: 140 },
    content: [{ runs: [{ text: 'MMMM', style: { color: 'FFFFFF', fontFamily: 'Arial', fontSize: 68, bold: true } }] }],
  }] });
  const finding = inspectSlideTextContrast(deck).findings[0];
  assert.equal(finding.status, 'meets-threshold');
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 320, height: 160 }, deviceScaleFactor: 1 });
  await page.setContent(`<style>html,body{margin:0}</style>${(await renderPresentationToSvg(deck)).slides[0].svg}`);
  await page.evaluate(() => document.fonts.ready);
  const pixels = await page.evaluate(async png => {
    const image = new Image(); image.src = `data:image/png;base64,${png}`; await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
    const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0);
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    let white = 0;
    for (let i = 0; i < data.length; i += 4) if (data[i] === 255 && data[i + 1] === 255 && data[i + 2] === 255) white++;
    return { white, background: Array.from(data.slice(0, 3)) };
  }, (await page.screenshot()).toString('base64'));
  assert.ok(pixels.white > 100);
  assert.ok(pixels.background.every(channel => channel >= 63 && channel <= 65));
  assert.ok(Math.abs(finding.ratio - contrast([255, 255, 255], pixels.background)) < 0.15);
});

test('Chromium draws shared master text with the observed per-slide contrast', async t => {
  const deck = createPresentation({ size: { width: 320, height: 160 }, theme: { colors: { accent1: '000000' } }, masters: [{
    id: 'm', name: 'Master', theme: { colors: { accent1: 'FFFFFF' } }, elements: [{
      type: 'text', box: { x: 20, y: 15, width: 270, height: 120 },
      content: [{ runs: [{ text: 'MMMM', style: { color: { theme: 'accent1' }, fontFamily: 'Arial', fontSize: 65, bold: true } }] }],
    }],
  }] });
  deck.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm' });
  for (const color of ['002244', 'BBBBBB']) deck.addSlide({ layoutId: 'l', background: { type: 'solid', color } });
  const findings = inspectSlideTextContrast(deck).findings;
  assert.deepEqual(findings.map(f => f.status), ['meets-threshold', 'below-threshold']);
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 320, height: 160 }, deviceScaleFactor: 1 });
  for (const [index, slide] of (await renderPresentationToSvg(deck)).slides.entries()) {
    await page.setContent(`<style>html,body{margin:0}</style>${slide.svg}`);
    await page.evaluate(() => document.fonts.ready);
    const samples = await page.evaluate(async png => {
      const image = new Image(); image.src = `data:image/png;base64,${png}`; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
      const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0);
      const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
      let white = 0;
      for (let i = 0; i < data.length; i += 4) if (data[i] === 255 && data[i + 1] === 255 && data[i + 2] === 255) white++;
      return { white, background: Array.from(data.slice(0, 3)) };
    }, (await page.screenshot()).toString('base64'));
    assert.ok(samples.white > 100);
    assert.ok(Math.abs(findings[index].ratio - contrast([255, 255, 255], samples.background)) < 1e-10);
  }
});

test('translation-only groups preserve foreground text contrast and nominal glyph size', async t => {
  const size = { width: 320, height: 160 }, background = { type: 'solid', color: '002244' };
  const label = (id, x, y) => ({
    id, type: 'text', box: { x, y, width: 220, height: 105 },
    content: [{ runs: [{ text: 'MMM', style: { color: 'FFFFFF', fontFamily: 'Arial', fontSize: 60, bold: true } }] }],
  });
  const group = (id, x, y, children) => ({
    id, type: 'group', box: { x, y, width: 245, height: 110 }, coordinateSize: { width: 245, height: 110 }, children,
  });
  const deck = createPresentation({ size });
  deck.addSlide({ id: 'direct', background, elements: [label('direct-label', 30, 25)] });
  deck.addSlide({ id: 'nested', background, elements: [group('outer', 20, 15, [group('inner', 10, 10, [label('nested-label', 0, 0)])])] });
  const [direct, nested] = inspectSlideTextContrast(deck).findings;
  assert.equal(nested.status, 'meets-threshold');
  assert.equal(nested.ratio, contrast([255, 255, 255], [0, 34, 68]));
  assert.equal(nested.ratio, direct.ratio);
  assert.equal(nested.effectiveFontSize, direct.effectiveFontSize);
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage({ viewport: size, deviceScaleFactor: 1 });
  const renders = (await renderPresentationToSvg(deck)).slides;
  const screenshots = [];
  for (const render of renders) {
    await page.setContent(`<style>html,body{margin:0}</style>${render.svg}`);
    await page.evaluate(() => document.fonts.ready);
    assert.equal(await page.locator('text tspan').first().evaluate(node => getComputedStyle(node).fontSize), '60px');
    screenshots.push(await page.screenshot());
  }
  assert.deepEqual(screenshots[0], screenshots[1]);
  const xml = readGeneratedZip((await generatePptx(deck)).bytes).get('ppt/slides/slide2.xml').toString();
  assert.equal([...xml.matchAll(/<a:ext cx="3111500" cy="1397000"\/><a:chOff x="0" y="0"\/><a:chExt cx="3111500" cy="1397000"\/>/g)].length, 2);
});
