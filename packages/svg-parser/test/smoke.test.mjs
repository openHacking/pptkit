import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, validatePresentation } from '@pptkit/core';
import { importSimpleSvg, importSvgWithFallback, createSvgRasterFallback, SvgImportError } from '../dist/index.js';

const wrap = inner => `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">${inner}</svg>`;
const code = expected => error => error instanceof SvgImportError && error.code === expected;

test('explicit raster fallback carries caller identity and copied bytes without implying SVG fidelity', () => {
  const bytes = new Uint8Array([1, 2, 3]);
  const options = { sourceId: 'source:example', assetId: 'raster-1', bytes, mimeType: 'image/png', width: 150, height: 75 };
  const result = createSvgRasterFallback(options); bytes[0] = 9;
  assert.equal(result.status, 'raster-fallback'); assert.equal(result.sourceId, options.sourceId);
  assert.equal(result.issue.code, 'rasterized'); assert.deepEqual([...result.asset.source.value], [1, 2, 3]);
  assert.deepEqual(result.element, { type: 'image', assetId: 'raster-1', box: { x: 0, y: 0, width: 150, height: 75 }, fit: 'stretch' });
  for (const patch of [{ sourceId: '' }, { assetId: ' ' }, { bytes: new Uint8Array() }, { bytes: 'abc' }, { mimeType: 'image/svg+xml' }, { width: 0 }, { height: Infinity }, { width: NaN }, { height: 1_000_001 }]) {
    assert.throws(() => createSvgRasterFallback({ ...options, ...patch }), TypeError);
  }
  assert.throws(() => createSvgRasterFallback({ ...options, bytes: new Uint8Array(20_000_001) }), TypeError);
});

test('import outcome retains source identity, failure reasons and explicit fallback policy', () => {
  const bytes = new Uint8Array([1,2,3]);
  const fallback = {assetId: 'fallback', bytes, mimeType: 'image/png', width: 150, height: 75};
  const options = {sourceId: 'caller:source', fallback};
  const editable = importSvgWithFallback(wrap('<rect width="10" height="10"/>'), options);
  assert.equal(editable.status, 'editable'); assert.equal(editable.sourceId, options.sourceId); assert.equal('asset' in editable, false);
  const unsupportedSource = wrap('<rect width="10" height="10"/><image href="https://example.invalid/a"/>');
  const unsupported = importSvgWithFallback(unsupportedSource, {sourceId: options.sourceId});
  assert.equal(unsupported.status, 'unsupported'); assert.ok(unsupported.issues.length); assert.equal('elements' in unsupported, false);
  const raster = importSvgWithFallback(unsupportedSource, options); assert.equal(raster.status, 'raster-fallback');
  assert.equal(raster.reason.status, 'unsupported'); assert.deepEqual(raster.reason.issues, unsupported.issues); assert.equal(raster.issue.code, 'rasterized');
  bytes[0] = 9; assert.equal(raster.asset.source.value[0], 1);
  const malformed = importSvgWithFallback('<svg>', options); assert.equal(malformed.status, 'raster-fallback'); assert.equal(malformed.reason.status, 'parse-error'); assert.equal(malformed.reason.code, 'invalid-xml');
  const limited = importSvgWithFallback(' '.repeat(1_000_001), {sourceId: options.sourceId}); assert.equal(limited.status, 'parse-error'); assert.equal(limited.code, 'input-limit');
  assert.throws(() => importSvgWithFallback(null, options), code('invalid-input'));
  assert.throws(() => importSvgWithFallback(unsupportedSource, {...options, fallback: {...fallback, bytes: new Uint8Array()}}), TypeError);
  assert.throws(() => importSvgWithFallback(unsupportedSource, {sourceId: ' '}), TypeError);
  assert.equal(importSvgWithFallback(wrap('<rect width="10" height="10"/>'), {...options, fallback: {...fallback, bytes: new Uint8Array()}}).status, 'editable');
});

test('maps ordered editable shapes, pixel units and default/literal fills into Core', () => {
  const result = importSimpleSvg(wrap('<rect x="10" y="20" width="40" height="30" fill="#abc"/><circle cx="120" cy="40" r="20"/><ellipse cx="170" cy="60" rx="10" ry="20" fill="none"/>'));
  assert.equal(result.status, 'editable'); assert.equal(result.width, 150); assert.equal(result.height, 75);
  assert.deepEqual(result.elements.map(item => [item.shape, item.box, item.style.fill]), [
    ['rect', { x: 7.5, y: 15, width: 30, height: 22.5 }, { type: 'solid', color: 'AABBCC' }],
    ['ellipse', { x: 75, y: 15, width: 30, height: 30 }, { type: 'solid', color: '000000' }],
    ['ellipse', { x: 120, y: 30, width: 15, height: 30 }, { type: 'none' }],
  ]);
  const doc = createPresentation({ size: { width: result.width, height: result.height } }); doc.addSlide({ elements: [...result.elements] });
  assert.deepEqual(validatePresentation(doc), []);
  const shapes = normalizePresentation(doc).slides[0].elements;
  assert.deepEqual(shapes.map(shape => shape.box), result.elements.map(item => item.box));
  assert.ok(shapes.every(shape => shape.style.stroke.paint.type === 'none'));
});

test('basic named colors and bounded comma RGB parse for inherited fill and stroke', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill=" ReD " stroke="rgb(0, 128, 255)" stroke-width="2"><circle cx="20" cy="20" r="10"/><g fill="rgb(1, 2, 3)" stroke=" NoNe "><rect x="50" y="10" width="10" height="10"/></g></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.elements[0].style.fill, { type: 'solid', color: 'FF0000' });
  assert.deepEqual(result.elements[0].style.stroke, { paint: { type: 'solid', color: '0080FF' }, width: 1.5 });
  assert.deepEqual(result.elements[1].style.fill, { type: 'solid', color: '010203' });
  assert.equal(result.elements[1].style.stroke.paint.type, 'none');
  for (const color of ['rgb(256,0,0)', 'rgb(100%,0%,0%)', 'rgba(1,2,3,inherit)', 'rgba(1,2,3,1.1)', 'rgb(1,2,3 / .5)', 'currentColor', 'constructor', 'url(#x)']) {
    const rejected = importSimpleSvg(wrap(`<rect width="10" height="10" fill="${color}"/>`));
    assert.equal(rejected.status, 'unsupported', color); assert.equal('elements' in rejected, false);
  }
});

test('integer RGB functions support comma or space syntax and bounded alpha', () => {
  const colors = ['rgba(255,0,0,.5)', 'rgb(0 128 255 / 25%)', 'rgba(1 2 3 / 0)', 'rgb(10,20,30,1)', 'RGBA(4,5,6,50%)'];
  const source = `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50" fill-opacity=".5">${colors.map((color, i) => `<rect x="${i * 15}" y="5" width="10" height="10" fill="${color}"/>`).join('')}</svg>`;
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.elements.map(item => item.style.fill), [
    {type: 'solid', color: 'FF0000', opacity: .25}, {type: 'solid', color: '0080FF', opacity: .125},
    {type: 'solid', color: '010203', opacity: 0}, {type: 'solid', color: '0A141E', opacity: .5}, {type: 'solid', color: '040506', opacity: .25},
  ]);
});

test('hex alpha and transparent colors retain independent inherited fill and stroke opacity', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f008" stroke="#00ff0080" stroke-width="2"><circle cx="20" cy="20" r="10"/><g fill="transparent"><circle cx="50" cy="20" r="10" stroke="#abcd"/></g><rect x="80" y="10" width="10" height="10" stroke="none"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.elements[0].style.fill, {type: 'solid', color: 'FF0000', opacity: 136/255});
  assert.deepEqual(result.elements[0].style.stroke.paint, {type: 'solid', color: '00FF00', opacity: 128/255});
  assert.deepEqual(result.elements[1].style.fill, {type: 'solid', color: '000000', opacity: 0});
  assert.deepEqual(result.elements[1].style.stroke.paint, {type: 'solid', color: 'AABBCC', opacity: 221/255});
  assert.deepEqual(result.elements[2].style.fill, result.elements[0].style.fill);
  for (const color of ['#abcd0', '#123456789', '#123g', '#abcdffgg']) assert.equal(importSimpleSvg(wrap(`<rect width="10" height="10" fill="${color}"/>`)).status, 'unsupported', color);
});

test('fill-opacity and stroke-opacity inherit independently and multiply color alpha', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f008" fill-opacity="50%" stroke="#00ff0080" stroke-opacity=".8" stroke-width="2"><circle cx="20" cy="20" r="10"/><g fill-opacity=".25" stroke-opacity="25%"><circle cx="60" cy="20" r="10"/><circle cx="100" cy="20" r="10" fill-opacity="inherit" stroke-opacity=".5"/></g><circle cx="140" cy="20" r="10"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  const [a,b,c,d] = result.elements;
  assert.equal(a.style.fill.opacity, 136 / 255 * .5); assert.equal(a.style.stroke.paint.opacity, 128 / 255 * .8);
  assert.equal(b.style.fill.opacity, 136 / 255 * .25); assert.equal(b.style.stroke.paint.opacity, 128 / 255 * .25);
  assert.equal(c.style.fill.opacity, b.style.fill.opacity); assert.equal(c.style.stroke.paint.opacity, 128 / 255 * .5);
  assert.deepEqual(d.style, a.style);
  for (const value of ['-0.1', '101%', 'NaN', '1px', '']) {
    const bad = importSimpleSvg(wrap(`<rect width="10" height="10" fill-opacity="${value}"/>`)); assert.equal(bad.status, 'unsupported', value);
  }
  assert.equal(importSimpleSvg(wrap('<g opacity="0.5"><rect width="10" height="10"/></g>')).status, 'unsupported');
});

test('explicit inherited fill, stroke and stroke width respect parent and sibling scope', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f008" stroke="blue" stroke-width="4"><g fill="inherit" stroke="inherit" stroke-width="inherit"><circle cx="20" cy="20" r="10"/><g fill="none" stroke="red" stroke-width="2"><path d="M50 10 L70 20" fill="inherit" stroke="inherit" stroke-width="inherit" stroke-linecap="round" stroke-linejoin="round"/></g></g><circle cx="100" cy="20" r="10" fill="inherit" stroke="inherit" stroke-width="inherit"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.elements[0].style.fill, {type: 'solid', color: 'FF0000', opacity: 136 / 255});
  assert.deepEqual(result.elements[0].style.stroke, {paint: {type: 'solid', color: '0000FF'}, width: 3});
  assert.deepEqual(result.elements[1].style, {paint: {type: 'solid', color: 'FF0000'}, width: 1.5});
  assert.deepEqual(result.elements[2].style, result.elements[0].style);
  const defaults = importSimpleSvg('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" fill="inherit" stroke="inherit" stroke-width="inherit"><circle cx="20" cy="20" r="10"/></svg>');
  assert.equal(defaults.status, 'editable'); assert.deepEqual(defaults.elements[0].style.fill, {type: 'solid', color: '000000'}); assert.equal(defaults.elements[0].style.stroke.paint.type, 'none');
  const rejected = importSimpleSvg(wrap('<g stroke-width="calc(2px)"><circle cx="20" cy="20" r="5"/></g>')); assert.equal(rejected.status, 'unsupported');
});

test('unsupported features never return a partial editable result', () => {
  for (const source of [
    wrap('<rect width="20" height="20"/><path d="M0 0L1 1"/>'),
    wrap('<rect width="20" height="20" style="fill:red"/>'),
    wrap('<rect width="20" height="20" fill="url(https://example.invalid/x)"/>'),
    wrap('<image href="https://example.invalid/x" width="10" height="10"/>'),
    wrap('<script>alert(1)</script>'), wrap('<g opacity=".5"><rect width="10" height="10"/></g>'),
    wrap('<rect x="190" width="20" height="10"/>'), wrap('<circle cx="0" cy="0" r="1"/>'),
    wrap('<rect width="0" height="10"/>'),
    '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100" preserveAspectRatio="xWrongYMin slice"/>',
    '<svg width="100" height="100"><rect width="10" height="10"/></svg>',
  ]) {
    const result = importSimpleSvg(source); assert.equal(result.status, 'unsupported', source);
    assert.ok(result.issues.length); assert.equal('elements' in result, false);
  }
});

test('nested group transform and fill inheritance flatten with an explicit warning', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="#f00"><g fill="#0f0" transform="translate(40,20)"><g transform="scale(-1,2) translate(-30,0)"><rect x="5" y="5" width="10" height="10" transform="translate(2,1)"/><rect x="15" y="5" width="5" height="5" fill="none"/></g></g><rect x="2" y="2" width="4" height="4"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.issues.map(issue => issue.code), ['flattened-groups']);
  assert.deepEqual(result.elements[0].box, { x: 39.75, y: 24, width: 7.5, height: 15 });
  assert.deepEqual(result.elements.map(shape => shape.style.fill), [{ type: 'solid', color: '00FF00' }, { type: 'none' }, { type: 'solid', color: 'FF0000' }]);
  for (const bad of [wrap('<g fill="url(https://example.invalid/a)"><rect width="10" height="10"/></g>'), wrap('<g transform="rotate(2)"><rect width="10" height="10"/></g>'), wrap('<g><rect width="10" height="10"/><path d="M0 0"/></g>'), wrap('<rect width="10" height="10"><g/></rect>')]) {
    const rejected = importSimpleSvg(bad); assert.equal(rejected.status, 'unsupported'); assert.equal('elements' in rejected, false);
  }
});

test('ellipse strokes retain literal paint and uniform scaling and reject clipping or anisotropic strokes', () => {
  const result = importSimpleSvg(wrap('<g transform="translate(30,20) scale(2)"><ellipse cx="15" cy="15" rx="10" ry="5" fill="none" stroke="#abc" stroke-width="2"/></g>'));
  assert.equal(result.status, 'editable');
  assert.deepEqual(result.elements[0].box, { x: 30, y: 30, width: 30, height: 15 });
  assert.deepEqual(result.elements[0].style.stroke, { paint: { type: 'solid', color: 'AABBCC' }, width: 3 });
  const noStroke = importSimpleSvg(wrap('<circle cx="10" cy="10" r="10" stroke="#000" stroke-width="0"/>'));
  assert.equal(noStroke.status, 'editable'); assert.equal(noStroke.elements[0].style.stroke.paint.type, 'none');
  for (const source of [wrap('<circle cx="10" cy="10" r="10" stroke="#000"/>'), wrap('<ellipse cx="30" cy="30" rx="10" ry="5" stroke="#000" transform="scale(2,1)"/>'), wrap('<circle cx="30" cy="30" r="5" stroke="url(#a)"/>'), wrap('<circle cx="30" cy="30" r="5" stroke="#000" stroke-width="-1"/>'), wrap('<rect x="10" y="10" width="10" height="10" stroke="#000"/>')]) {
    const rejected = importSimpleSvg(source); assert.equal(rejected.status, 'unsupported', source); assert.equal('elements' in rejected, false);
  }
});

test('inherited stroke and width follow root, nested overrides and sibling scope', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" stroke="#f00" stroke-width="4"><g stroke="#0f0" stroke-width="2" transform="translate(20,0)"><circle cx="20" cy="20" r="8"/><g stroke="none"><rect x="40" y="10" width="10" height="10"/><circle cx="65" cy="20" r="8" stroke="#00f"/></g></g><circle cx="120" cy="20" r="8"/><rect x="150" y="10" width="10" height="10" stroke-width="0"/></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.elements.map(element => element.style.stroke), [
    { paint: { type: 'solid', color: '00FF00' }, width: 1.5 }, { paint: { type: 'none' } },
    { paint: { type: 'solid', color: '0000FF' }, width: 1.5 }, { paint: { type: 'solid', color: 'FF0000' }, width: 3 }, { paint: { type: 'none' } },
  ]);
  for (const source of [wrap('<g stroke="#f00"><rect x="10" y="10" width="10" height="10"/></g>'), wrap('<g stroke="url(#x)"><circle cx="20" cy="20" r="5"/></g>'), wrap('<g stroke-width="-2"><circle cx="20" cy="20" r="5"/></g>')]) {
    const rejected = importSimpleSvg(source); assert.equal(rejected.status, 'unsupported'); assert.equal('elements' in rejected, false);
  }
});

test('linear SVG paths map to bounded editable connectors with explicit closure warning', () => {
  const source = wrap('<g transform="translate(10,10) scale(2)" fill="none" stroke="red" stroke-width="2"><path d="M 10 10 l 10 0 v 10 H 10 z" stroke-linecap="round" stroke-linejoin="round"/></g>');
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.issues.map(issue => issue.code), ['flattened-groups', 'path-converted']);
  assert.deepEqual(result.elements[0], { type: 'connector', start: {x: 22.5, y: 22.5}, end: {x: 22.5, y: 22.5}, route: [{x: 37.5, y: 22.5}, {x: 37.5, y: 37.5}, {x: 22.5, y: 37.5}], pathMode: 'closed-round', style: {paint: {type: 'solid', color: 'FF0000'}, width: 3} });
  const doc = createPresentation({size: {width: result.width, height: result.height}}); doc.addSlide({elements: [...result.elements]});
  assert.deepEqual(validatePresentation(doc), []);
  const attrs = 'fill="none" stroke="red" stroke-linecap="round" stroke-linejoin="round"';
  for (const path of [
    `<path ${attrs} d="M0 0 L10 10"/>`, `<path ${attrs} d="M10 10 C20 20 30 30 40 40"/>`,
    `<path ${attrs} d="M10 10 L20 20 M30 30 L40 40"/>`,
    `<path ${attrs} d="M10 10 L10 10"/>`,
    `<path ${attrs} d="M10 10 L20 20 Z L30 30"/>`, `<path ${attrs} d="M10 10 L20 20" transform="scale(2,1)"/>`,
    '<path d="M10 10 L20 20" stroke="red" stroke-linecap="round" stroke-linejoin="round"/>',
    '<path d="M10 10 L20 20" fill="none" stroke="red"/>',
  ]) { const rejected = importSimpleSvg(wrap(path)); assert.equal(rejected.status, 'unsupported', path); assert.equal('elements' in rejected, false); }
});

test('implicit path coordinates repeat commands with SVG moveto and relative semantics', () => {
  const attrs = 'fill="none" stroke="red" stroke-linecap="round" stroke-linejoin="round"';
  const svg = data => wrap(`<path ${attrs} d="${data}"/>`);
  const implicit = importSimpleSvg(svg('m20 20 10 0 0 10 h-5-5 v-10 z'));
  const explicit = importSimpleSvg(svg('m20 20 l10 0 l0 10 h-5 h-5 v-10 z'));
  // Closing this path would repeat its current point and is deliberately rejected.
  assert.equal(implicit.status, 'unsupported'); assert.equal(explicit.status, 'unsupported');
  const a = importSimpleSvg(svg('m20 20 10 0 0 10 h-5-5 v-5 z'));
  const b = importSimpleSvg(svg('m20 20 l10 0 l0 10 h-5 h-5 v-5 z'));
  assert.equal(a.status, 'editable'); assert.deepEqual(a, b);
  for (const data of ['M20 20,', 'M20 20,,30 30', 'M20 20, L30 30', 'M20 20 30', 'M20 20 L30 30 Z,', 'M20 20 L30 30 40', 'M20 20 L30 30Z40 40']) {
    assert.equal(importSimpleSvg(svg(data)).status, 'unsupported', data);
  }
});

test('path line caps and joins inherit across root and groups with local overrides', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="none" stroke="red" stroke-linecap="round" stroke-linejoin="round"><path d="M10 10 L30 20"/><g stroke-linecap="butt"><path d="M50 10 L70 20" stroke-linecap="inherit"/></g></svg>';
  const rejected = importSimpleSvg(source); assert.equal(rejected.status, 'unsupported'); assert.equal('elements' in rejected, false);
  const accepted = importSimpleSvg(source.replace('stroke-linecap="inherit"', 'stroke-linecap="round"'));
  assert.equal(accepted.status, 'editable'); assert.equal(accepted.elements.length, 2); assert.ok(accepted.elements.every(item => item.pathMode === 'open-round'));
  const inherited = importSimpleSvg(source.replace('stroke-linecap="butt"', 'stroke-linecap="round"'));
  assert.equal(inherited.status, 'editable');
  for (const attribute of ['stroke-linecap="banana"', 'stroke-linejoin="miter-clip"', 'stroke-linecap="square"']) {
    const result = importSimpleSvg(source.replace('stroke-linecap="butt"', attribute)); assert.equal(result.status, 'unsupported', attribute);
  }
});

test('polyline and polygon map into open or closed connectors with inherited styles', () => {
  const source = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="none" stroke="blue" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="10,10 30,10 30,30"/><g transform="translate(70,0)" stroke-linecap="butt"><polygon points="10 10, 30 10, 30 30, 10 30"/></g></svg>';
  const result = importSimpleSvg(source); assert.equal(result.status, 'editable');
  assert.deepEqual(result.elements.map(item => item.pathMode), ['open-round', 'closed-round']);
  assert.deepEqual(result.elements[0].start, {x: 7.5, y: 7.5});
  assert.deepEqual(result.elements[1].route, [{x: 75, y: 7.5}, {x: 75, y: 22.5}, {x: 60, y: 22.5}]);
  assert.deepEqual(result.elements[1].end, result.elements[1].start);
  assert.deepEqual(result.issues.map(issue => issue.code), ['path-converted', 'flattened-groups']);
  for (const element of ['<polyline points="10,10 30,10"/>', '<polygon points="10,10 30,10 30,30" fill="red"/>', '<polygon points="10,10 L30,10 L30,30"/>', '<polygon points="10,10 30,10 30,30,"/>', '<polygon points="10,10 30,10"/>', '<polyline points="10,10 30,10" stroke-linejoin="bevel"/>']) {
    const source = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" fill="none" stroke="blue" stroke-linejoin="round">${element}</svg>`;
    const bad = importSimpleSvg(source); assert.equal(bad.status, 'unsupported', element); assert.equal('elements' in bad, false);
  }
  const closed = importSimpleSvg('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" fill="none" stroke="red" stroke-linejoin="round"><path d="M10 10 L40 10 L40 40 Z"/></svg>');
  assert.equal(closed.status, 'editable'); assert.equal(closed.elements[0].pathMode, 'closed-round');
});

test('viewBox mapping applies centered meet or nonuniform none and rejects clipping', () => {
  const svg = (aspect, shape, viewBox = '-10,-20,100,100') => `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="${viewBox}" ${aspect}>${shape}</svg>`;
  const rect = '<rect x="0" y="0" width="20" height="40" fill="#123456"/>';
  const meet = importSimpleSvg(svg('', rect)); assert.equal(meet.status, 'editable');
  assert.deepEqual(meet.elements[0].box, { x: 45, y: 15, width: 15, height: 30 });
  const none = importSimpleSvg(svg('preserveAspectRatio="none"', rect)); assert.equal(none.status, 'editable');
  assert.deepEqual(none.elements[0].box, { x: 15, y: 15, width: 30, height: 30 });
  const circle = importSimpleSvg(svg('preserveAspectRatio="none"', '<circle cx="40" cy="30" r="10"/>'));
  assert.equal(circle.status, 'editable'); assert.deepEqual(circle.elements[0].box, { x: 60, y: 30, width: 30, height: 15 });
  for (const source of [svg('', '<rect x="-11" y="0" width="20" height="10"/>'), svg('', rect, '0 0 0 100'), svg('', rect, '0,,0,100,100'), svg('', rect, '0 0 1e-320 100'), svg('', rect, '0 0 100px 100'), svg('preserveAspectRatio="xMidYMid slice"', '<rect x="0" y="-20" width="10" height="10"/>')]) {
    const result = importSimpleSvg(source); assert.equal(result.status, 'unsupported', source); assert.equal('elements' in result, false);
  }
});

test('all viewBox alignment modes map contained shapes and reject slice clipping', () => {
  const values = { Min: 0, Mid: 0.5, Max: 1 };
  for (const [xName, fx] of Object.entries(values)) for (const [yName, fy] of Object.entries(values)) for (const mode of ['meet', 'slice']) {
    const x = mode === 'meet' ? 50 : 25, y = mode === 'meet' ? 50 : 20 + 50 * fy;
    const input = `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100" viewBox="0 0 100 100" preserveAspectRatio="x${xName}Y${yName} ${mode}"><rect x="${x}" y="${y}" width="10" height="10"/></svg>`;
    const result = importSimpleSvg(input); assert.equal(result.status, 'editable', input);
    assert.deepEqual(result.elements[0].box, mode === 'meet' ? { x: (50 + 100 * fx) * .75, y: 37.5, width: 7.5, height: 7.5 } : { x: 37.5, y: 30, width: 15, height: 15 });
  }
});

test('ordered shape translation and scale preserve axis-aligned geometry, including reflection', () => {
  const sample = transform => importSimpleSvg(wrap(`<rect x="5" y="4" width="10" height="5" fill="#abc" transform="${transform}"/>`));
  const first = sample('translate(10,20) scale(2,3)'); assert.equal(first.status, 'editable');
  assert.deepEqual(first.elements[0].box, { x: 15, y: 24, width: 15, height: 11.25 });
  const second = sample('scale(2,3) translate(10,20)'); assert.equal(second.status, 'editable');
  assert.deepEqual(second.elements[0].box, { x: 22.5, y: 54, width: 15, height: 11.25 });
  const reflected = sample('scale(-1,1) translate(-40,0)'); assert.equal(reflected.status, 'editable');
  assert.deepEqual(reflected.elements[0].box, { x: 18.75, y: 3, width: 7.5, height: 3.75 });
  const reentered = importSimpleSvg(wrap('<rect x="-10" y="0" width="5" height="5" transform="translate(20)"/>'));
  assert.equal(reentered.status, 'editable'); assert.equal(reentered.elements[0].box.x, 7.5);
  for (const value of ['rotate(30)', 'skewX(20)', 'matrix(1,0,0,1,0,0)', 'scale(0)', 'translate(200)', 'scale(1e999)', 'translate(1px)', 'translate(1),', 'translate(1,,2)', '']) {
    const result = sample(value); assert.equal(result.status, 'unsupported', value); assert.equal('elements' in result, false);
  }
  assert.equal(sample('translate(1) '.repeat(65).trim()).status, 'unsupported');
});

test('parser rejects DTD, malformed input, excessive size, depth and attributes', () => {
  assert.throws(() => importSimpleSvg('<!DOCTYPE svg [<!ENTITY x "unsafe">]>' + wrap('<rect width="1" height="1"/>')), code('invalid-xml'));
  assert.throws(() => importSimpleSvg(wrap('<rect width="1" height="1">')), code('invalid-xml'));
  assert.throws(() => importSimpleSvg(wrap('&unknown;')), code('invalid-xml'));
  assert.throws(() => importSimpleSvg(' '.repeat(1_000_001)), code('input-limit'));
  assert.throws(() => importSimpleSvg(wrap('<g>'.repeat(65) + '</g>'.repeat(65))), code('input-limit'));
  assert.throws(() => importSimpleSvg(wrap(`<rect ${Array.from({length:65}, (_, i) => `a${i}="x"`).join(' ')} />`)), code('input-limit'));
  assert.throws(() => importSimpleSvg(null), code('invalid-input'));
});
