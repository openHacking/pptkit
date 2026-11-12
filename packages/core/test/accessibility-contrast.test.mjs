import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, inspectSlideTextContrast, TextContrastLimitError, normalizePresentation, PresentationValidationError } from '../dist/index.js';
const box = { x: 0, y: 0, width: 200, height: 100 };
const text = (id, content, extra = {}) => ({ id, type: 'text', box, content, ...extra });
const run = (value, style) => ({ text: value, style });

test('contrast compares sRGB theme colors and transparent foreground without rounding thresholds', () => {
  const deck = createPresentation({ theme: { colors: { background1: 'FFFFFF', text1: '000000' } } });
  deck.addSlide({ id: 's', background: { type: 'solid', color: { theme: 'background1' } }, elements: [text('t', [{ runs: [
    run('black', { color: { theme: 'text1' }, fontSize: 12 }),
    run('low', { color: '777777', fontSize: 12 }),
    run('large', { color: '777777', fontSize: 18 }),
    run('bold', { color: '777777', fontSize: 14, bold: true }),
    run('svg-large', { color: '777777', fontSize: 24 }),
    run('svg-bold', { color: '777777', fontSize: 56 / 3, bold: true }),
    run('alpha', { color: '000000', colorOpacity: 0.5, fontSize: 12 }),
    run(' ', { color: '000000' }),
  ] }])] });
  const before = JSON.stringify(normalizePresentation(deck));
  const result = inspectSlideTextContrast(deck), findings = result.findings;
  assert.equal(result.scope, 'slide-text-runs'); assert.equal(findings.length, 7);
  assert.equal(findings[0].ratio, 21); assert.equal(findings[0].threshold, 4.5); assert.equal(findings[0].status, 'meets-threshold');
  assert.equal(findings[1].status, 'below-threshold'); assert.ok(findings[1].ratio < 4.5);
  for (const finding of findings.slice(2, 4)) { assert.equal(finding.status, 'indeterminate'); assert.equal(finding.reason, 'output-size-dependent'); assert.equal(finding.nativeThreshold, 3); assert.equal(finding.nominalSvgThreshold, 4.5); }
  assert.equal(findings[4].threshold, 3); assert.equal(findings[4].status, 'meets-threshold');
  assert.equal(findings[5].threshold, 3); assert.equal(findings[5].status, 'meets-threshold');
  assert.ok(findings[6].ratio > 3.9 && findings[6].ratio < 4.1); assert.equal(findings[6].status, 'below-threshold');
  assert.equal(JSON.stringify(normalizePresentation(deck)), before);
  assert.ok(Object.isFrozen(result) && Object.isFrozen(findings) && Object.isFrozen(findings[0]));
});

test('unknown scenes remain indeterminate and retain owner and run paths', () => {
  const deck = createPresentation();
  deck.addSlide({ id: 'overlap', elements: [text('t1', 'A'), { type: 'shape', id: 'shape', shape: 'rect', box }] });
  deck.addSlide({ id: 'gradient', background: { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }] }, elements: [text('t2', 'B')] });
  deck.addSlide({ id: 'alpha', elements: [text('t3', 'C', { opacity: 0.5 })] });
  deck.addSlide({ id: 'nested', elements: [{ id: 'g', type: 'group', box, coordinateSize: { width: 100, height: 100 }, children: [text('t4', 'D')] }] });
  deck.addSlide({ id: 'shape-text', elements: [{ type: 'shape', id: 'st', shape: 'rect', box, style: { fill: { type: 'solid', color: '000000' } }, text: { content: 'E' } }] });
  deck.addSlide({ id: 'table', elements: [{ type: 'table', id: 'tbl', box, columns: [200], rows: [{ cells: [{ content: 'F', style: { fill: { type: 'solid', color: '000000' } } }] }] }] });
  deck.addSlide({ id: 'hidden', hidden: true, elements: [text('hidden-text', 'G')] });
  const rows = inspectSlideTextContrast(deck).findings;
  assert.deepEqual(rows.map(r => r.reason), ['complex-scene', 'non-solid-background', 'translucent-object', 'complex-scene', 'unsupported-text-owner', 'unsupported-text-owner']);
  assert.ok(rows.every(r => r.status === 'indeterminate' && !('ratio' in r)));
  assert.equal(rows[3].path, 'slides.3.elements.0.children.0.content.0.runs.0');
  assert.equal(rows[5].path, 'slides.5.elements.0.rows.0.cells.0.content.0.runs.0');
});

test('size scaling, inherited artwork and invalid documents are handled conservatively', () => {
  const deck = createPresentation();
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', elements: [{ type: 'shape', shape: 'rect', box }] });
  deck.addSlide({ id: 'inherited', layoutId: 'layout', elements: [text('t', 'A')] });
  deck.addSlide({ id: 'scale', elements: [text('scaled', [{ runs: [run('large before shrink', { fontSize: 18, color: '777777' })] }], { frame: { autoFit: { mode: 'shrink', fontScale: 0.5 } } })] });
  const result = inspectSlideTextContrast(deck);
  assert.equal(result.findings[0].reason, 'complex-scene');
  assert.equal(result.findings[1].effectiveFontSize, 9); assert.equal(result.findings[1].threshold, 4.5);
  const invalid = createPresentation(); invalid.addSlide({ elements: [text('bad', 'X', { box: { ...box, width: -1 } })] });
  assert.throws(() => inspectSlideTextContrast(invalid), PresentationValidationError);
});

test('event timelines and decorative groups do not produce false positive observations', () => {
  const deck = createPresentation();
  deck.addSlide({ id: 'timed', elements: [text('t', 'Timed')], eventTimeline: { stages: [{ id: 'stage', target: { pageId: 'timed', elementId: 't' }, trigger: { type: 'object-click', elementId: 't' }, timing: { kind: 'span', id: 'time', durationMs: 10, fill: 'hold' } }] } });
  deck.addSlide({ id: 'decorative', elements: [{ id: 'g', type: 'group', box, coordinateSize: { width: 200, height: 100 }, accessibility: { decorative: true }, children: [text('skipped', 'Text')] }] });
  const findings = inspectSlideTextContrast(deck).findings;
  assert.equal(findings.length, 1); assert.equal(findings[0].status, 'indeterminate'); assert.equal(findings[0].reason, 'dynamic-timeline');
});

test('threshold decisions account for native and SVG serialization rounding', () => {
  const deck = createPresentation();
  deck.addSlide({ elements: [text('rounding', [{ runs: [
    run('native-round', { color: '777777', fontSize: 17.999 }),
    run('svg-round', { color: '777777', fontSize: 18.6666, bold: true }),
    run('both-fail', { color: 'AAAAAA', fontSize: 18 }),
    run('both-pass', { color: '000000', fontSize: 18 }),
  ] }])] });
  const rows = inspectSlideTextContrast(deck).findings;
  assert.equal(rows[0].reason, 'output-size-dependent');
  assert.equal(rows[1].status, 'meets-threshold'); assert.equal(rows[1].nominalSvgThreshold, 3);
  assert.equal(rows[2].status, 'below-threshold');
  assert.equal(rows[3].status, 'meets-threshold'); assert.equal(rows[3].threshold, 4.5);
});

test('opacity serialization cannot turn a borderline output into a definite result', () => {
  const deck = createPresentation();
  deck.addSlide({ elements: [text('alpha-rounding', [{ runs: [
    run('svg-only-pass', { color: '000000', colorOpacity: 0.534681, fontSize: 12 }),
    run('native-only-pass', { color: '000000', colorOpacity: 0.4161684, fontSize: 24 }),
    run('both-pass', { color: '000000', colorOpacity: 0.6, fontSize: 12 }),
    run('both-fail', { color: '000000', colorOpacity: 0.4, fontSize: 12 }),
  ] }])] });
  const rows = inspectSlideTextContrast(deck).findings;
  assert.equal(rows[0].reason, 'output-opacity-dependent'); assert.ok(rows[0].ratio >= 4.5 && rows[0].nativeRatio < 4.5);
  assert.equal(rows[1].reason, 'output-opacity-dependent'); assert.ok(rows[1].ratio < 3 && rows[1].nativeRatio >= 3);
  assert.equal(rows[2].status, 'meets-threshold'); assert.equal(rows[3].status, 'below-threshold');
});

test('shape text is calculated only when shape and surrounding field are uniformly identical', () => {
  const deck = createPresentation();
  const shape = (id, fill, stroke = { paint: { type: 'none' }, width: 0 }, extra = {}) => ({ id, type: 'shape', shape: 'rect', box, style: { fill, stroke, ...extra }, text: { content: [{ runs: [run('X', { color: '000000', fontSize: 12 })] }] } });
  const background = { type: 'solid', color: 'FFFFFF' };
  deck.addSlide({ id: 'no-fill', background, elements: [shape('one', { type: 'none' })] });
  deck.addSlide({ id: 'same-fill', background, elements: [shape('two', { type: 'solid', color: 'FFFFFF', opacity: 0.5 })] });
  deck.addSlide({ id: 'different', background, elements: [shape('three', { type: 'solid', color: '000000' })] });
  deck.addSlide({ id: 'stroke', background, elements: [shape('four', { type: 'none' }, { paint: { type: 'solid', color: '000000' }, width: 2 })] });
  deck.addSlide({ id: 'shadow', background, elements: [shape('five', { type: 'none' }, undefined, { shadow: { color: '000000', blur: 1, distance: 1, direction: 0 } })] });
  const result = inspectSlideTextContrast(deck).findings;
  assert.deepEqual(result.map(f => f.status), ['meets-threshold', 'meets-threshold', 'indeterminate', 'indeterminate', 'indeterminate']);
  assert.equal(result[0].ratio, 21); assert.equal(result[1].ratio, 21);
  assert.ok(result.slice(2).every(f => f.reason === 'unsupported-text-owner'));
});

test('opaque full-slide rectangles supply a uniform text field independently of the slide background', () => {
  const deck = createPresentation({ size: { width: 320, height: 160 }, theme: { colors: { text1: 'FFFFFF', background1: '002244' } } });
  const full = { x: 0, y: 0, width: 320, height: 160 }, noStroke = { paint: { type: 'none' }, width: 0 };
  const makeShape = (id, changes = {}) => ({ id, type: 'shape', shape: 'rect', box: full, style: { fill: { type: 'solid', color: { theme: 'background1' } }, stroke: noStroke }, text: { content: [{ runs: [run('White', { color: { theme: 'text1' }, fontSize: 12 })] }] }, ...changes });
  deck.addSlide({ id: 'solid', background: { type: 'solid', color: 'FFFFFF' }, elements: [makeShape('s')] });
  deck.addSlide({ id: 'gradient', background: { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }] }, elements: [makeShape('g', { transform: { flipH: true, flipV: true } })] });
  const white = { type: 'solid', color: 'FFFFFF' };
  deck.addSlide({ id: 'partial', background: white, elements: [makeShape('p', { box: { ...full, width: 319 } })] });
  deck.addSlide({ id: 'rotated', background: white, elements: [makeShape('r', { transform: { rotation: 1 } })] });
  deck.addSlide({ id: 'ellipse', background: white, elements: [makeShape('e', { shape: 'ellipse' })] });
  deck.addSlide({ id: 'fill-alpha', background: white, elements: [makeShape('f', { box: { ...full, width: 319 }, style: { fill: { type: 'solid', color: '002244', opacity: 0.5 }, stroke: noStroke } })] });
  deck.addSlide({ id: 'object-alpha', elements: [makeShape('a', { opacity: 0.5 })] });
  deck.addSlide({ id: 'overlap', elements: [makeShape('o'), text('other', 'Other')] });
  const before = JSON.stringify(normalizePresentation(deck)), findings = inspectSlideTextContrast(deck).findings;
  assert.equal(JSON.stringify(normalizePresentation(deck)), before);
  assert.equal(findings[0].status, 'meets-threshold'); assert.equal(findings[1].status, 'meets-threshold');
  assert.ok(Math.abs(findings[0].ratio - findings[1].ratio) < 1e-12); assert.ok(findings[0].ratio > 15);
  for (const finding of findings.slice(2, 6)) assert.equal(finding.reason, 'unsupported-text-owner');
  assert.equal(findings[6].reason, 'translucent-object'); assert.equal(findings[7].reason, 'complex-scene');
});

test('full-slide translucent fill composites separately at SVG and PPTX precision', () => {
  const deck = createPresentation({ size: { width: 320, height: 160 } });
  const shape = (id, opacity) => ({ id, type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 320, height: 160 }, style: { fill: { type: 'solid', color: '000000', opacity }, stroke: { paint: { type: 'none' }, width: 0 } }, text: { content: [{ runs: [run('Text', { color: '000000', fontSize: 12 })] }] } });
  const white = { type: 'solid', color: 'FFFFFF' }, gradient = { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: 'FFFFFF' }, { position: 1, color: 'FF0000' }] };
  deck.addSlide({ id: 'half', background: white, elements: [shape('half-shape', 0.5)] });
  deck.addSlide({ id: 'borderline', background: white, elements: [shape('borderline-shape', 0.5446693)] });
  deck.addSlide({ id: 'unknown', background: gradient, elements: [shape('unknown-shape', 0.5)] });
  const findings = inspectSlideTextContrast(deck).findings;
  assert.equal(findings[0].status, 'meets-threshold'); assert.ok(Math.abs(findings[0].ratio - 5.2808228) < 1e-6); assert.equal(findings[0].ratio, findings[0].nativeRatio);
  assert.equal(findings[1].status, 'indeterminate'); assert.equal(findings[1].reason, 'output-opacity-dependent'); assert.ok(findings[1].ratio >= 4.5 && findings[1].nativeRatio < 4.5);
  assert.equal(findings[2].reason, 'unsupported-text-owner');
});

test('slide-covering shapes require coverage before and after independent EMU rounding', () => {
  const width = 320.63942679845786, x = -0.025010755222666936;
  const deck = createPresentation({ size: { width, height: 160 } });
  const shape = (id, shapeBox) => ({ id, type: 'shape', shape: 'rect', box: shapeBox, style: { fill: { type: 'solid', color: '002244' }, stroke: { paint: { type: 'none' }, width: 0 } }, text: { content: [{ runs: [run('White', { color: 'FFFFFF', fontSize: 12 })] }] } });
  const white = { type: 'solid', color: 'FFFFFF' };
  deck.addSlide({ id: 'oversized', background: white, elements: [shape('oversized-shape', { x: -10, y: -10, width: width + 20, height: 180 })] });
  deck.addSlide({ id: 'emu-gap', background: white, elements: [shape('gap-shape', { x, y: 0, width: width - x, height: 160 })] });
  deck.addSlide({ id: 'emu-covered', background: white, elements: [shape('covered-shape', { x, y: 0, width: width - x + 1 / 12700, height: 160 })] });
  deck.addSlide({ id: 'actual-gap', background: white, elements: [shape('actual-gap-shape', { x: 0.00001, y: 0, width, height: 160 })] });
  const rows = inspectSlideTextContrast(deck).findings;
  assert.deepEqual(rows.map(row => row.status), ['meets-threshold', 'indeterminate', 'meets-threshold', 'indeterminate']);
  assert.equal(rows[1].reason, 'unsupported-text-owner'); assert.equal(rows[3].reason, 'unsupported-text-owner');
  assert.equal(Math.round(x * 12700) + Math.round((width - x) * 12700), Math.round(width * 12700) - 1);
});

test('shared layout lookups retain distinct master themes and artwork decisions', () => {
  const deck = createPresentation({ masters: [
    { id: 'dark', name: 'Dark', theme: { colors: { background1: '000000', text1: 'FFFFFF' } } },
    { id: 'light', name: 'Light', theme: { colors: { background1: 'FFFFFF', text1: '000000' } }, elements: [{ type: 'shape', id: 'art', shape: 'rect', box }] },
  ] });
  deck.defineSlideLayout({ id: 'dark-layout', name: 'Dark', masterId: 'dark' });
  deck.defineSlideLayout({ id: 'light-layout', name: 'Light', masterId: 'light' });
  for (let index = 0; index < 6; index++) {
    deck.addSlide({ id: `s${index}`, layoutId: index % 2 ? 'light-layout' : 'dark-layout', background: { type: 'solid', color: { theme: 'background1' } }, elements: [text(`t${index}`, [{ runs: [run('A', { color: { theme: 'text1' }, fontSize: 12 })] }])] });
  }
  const rows = inspectSlideTextContrast(deck).findings;
  assert.deepEqual(rows.map(row => row.status), ['meets-threshold', 'indeterminate', 'meets-threshold', 'indeterminate', 'meets-threshold', 'indeterminate']);
  assert.ok(rows.filter(row => row.status === 'meets-threshold').every(row => row.ratio === 21));
  assert.ok(rows.filter(row => row.status === 'indeterminate').every(row => row.reason === 'complex-scene'));
});

test('inherited artwork visibility gates contrast without hiding the effective background', () => {
  const artwork = id => ({ id, type: 'shape', shape: 'rect', box, style: { fill: { type: 'solid', color: 'FF0000' } } });
  const deck = createPresentation({ masters: [{ id: 'master', name: 'Master', background: { type: 'solid', color: '000000' }, elements: [artwork('master-art')] }] });
  deck.defineSlideLayout({ id: 'master-visible', name: 'Master visible', masterId: 'master' });
  deck.defineSlideLayout({ id: 'master-hidden', name: 'Master hidden', masterId: 'master', showMasterShapes: false });
  deck.defineSlideLayout({ id: 'layout-visible', name: 'Layout visible', masterId: 'master', showMasterShapes: false, elements: [artwork('layout-art')] });
  const label = id => text(id, [{ runs: [run('white', { color: 'FFFFFF', fontSize: 12 })] }]);
  deck.addSlide({ id: 'all-visible', layoutId: 'master-visible', elements: [label('a')] });
  deck.addSlide({ id: 'master-only-hidden', layoutId: 'master-hidden', elements: [label('b')] });
  deck.addSlide({ id: 'layout-still-visible', layoutId: 'layout-visible', elements: [label('c')] });
  deck.addSlide({ id: 'all-inherited-hidden', layoutId: 'layout-visible', showInheritedShapes: false, elements: [label('d')] });
  deck.addSlide({ id: 'slide-explicit-true', layoutId: 'master-hidden', showInheritedShapes: true, elements: [label('e')] });
  const rows = inspectSlideTextContrast(deck).findings;
  assert.deepEqual(rows.map(row => row.status), ['indeterminate', 'meets-threshold', 'indeterminate', 'meets-threshold', 'meets-threshold']);
  assert.ok(rows.filter(row => row.status === 'indeterminate').every(row => row.reason === 'complex-scene'));
  assert.ok(rows.filter(row => row.status === 'meets-threshold').every(row => row.ratio === 21));
});

test('table text is calculated only when every fill and border preserves the background field', () => {
  const deck = createPresentation();
  const invisible = { paint: { type: 'none' }, width: 0 };
  const cell = (content, extra = {}) => ({ content: [{ runs: [run(content, { color: '000000', fontSize: 12 })] }], style: { fill: { type: 'solid', color: { theme: 'background1' } }, stroke: invisible, ...extra } });
  const table = (id, rows, style) => ({ type: 'table', id, box, columns: [100, 100], rows, ...(style ? { style } : {}) });
  const background = { type: 'solid', color: { theme: 'background1' } };
  deck.addSlide({ id: 'uniform', background, elements: [table('uniform-table', [{ cells: [{ ...cell('A'), rowSpan: 2 }, cell('B')] }, { cells: [cell('C')] }])] });
  deck.addSlide({ id: 'mixed-fill', background, elements: [table('mixed', [{ cells: [cell('A'), cell('B', { fill: { type: 'solid', color: '000000' } })] }])] });
  deck.addSlide({ id: 'edge', background, elements: [table('bordered', [{ cells: [cell('A', { borders: { right: { paint: { type: 'solid', color: '000000' }, width: 1 } } }), cell('B')] }])] });
  deck.addSlide({ id: 'region', background, elements: [table('regional', [{ cells: [{ content: 'A' }, { content: 'B' }] }], { wholeTable: { fill: { type: 'solid', color: 'FFFFFF' }, stroke: invisible }, borders: { innerVertical: { paint: { type: 'solid', color: '000000' }, width: 1 } } })] });
  deck.addSlide({ id: 'transparent', background, elements: [table('transparent-table', [{ cells: [cell('A', { fill: { type: 'solid', color: '000000', opacity: 0 } }), cell('B')] }])] });
  const rows = inspectSlideTextContrast(deck).findings;
  assert.deepEqual(rows.filter(row => row.slideId === 'uniform').map(row => row.ratio), [21, 21, 21]);
  assert.ok(rows.filter(row => row.slideId !== 'uniform').every(row => row.status === 'indeterminate' && row.reason === 'unsupported-text-owner'));
  assert.equal(rows[2].path, 'slides.0.elements.0.rows.1.cells.0.content.0.runs.0');
});

test('ordered full-slide master, layout and local fills supply a uniform field for foreground text', () => {
  const fill = (id, color, opacity = 1, extra = {}) => ({
    id, type: 'shape', shape: 'rect', box, style: { fill: { type: 'solid', color, opacity }, stroke: { paint: { type: 'none' } } }, ...extra,
  });
  const deck = createPresentation({ size: { width: 200, height: 100 }, masters: [{
    id: 'master', name: 'Master', theme: { colors: { accent1: '000000' } }, elements: [fill('base', { theme: 'accent1' })],
  }] });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [fill('middle', 'FFFFFF', 0.5)] });
  const gradient = { type: 'linearGradient', angle: 0, stops: [{ position: 0, color: 'FF0000' }, { position: 1, color: '00FF00' }] };
  const content = [{ runs: [run('text', { color: 'FFFFFF', fontSize: 12 })] }];
  deck.addSlide({ id: 'known', layoutId: 'layout', background: gradient, elements: [fill('top', '000000', 0.5), text('front', content)] });
  deck.addSlide({ id: 'unknown', layoutId: 'layout', showInheritedShapes: false, background: gradient, elements: [fill('top-alt', '000000', 0.5), text('front-alt', content)] });
  deck.addSlide({ id: 'above', elements: [text('behind', content), fill('cover-above', '000000')] });
  deck.addSlide({ id: 'partial', elements: [fill('partial-cover', '000000', 1, { box: { ...box, width: 199.9 } }), text('front-partial', content)] });
  deck.addSlide({ id: 'obj-alpha', elements: [fill('alpha-cover', '000000', 1, { opacity: 0.5 }), text('front-alpha', content)] });
  const [known, unknown, above, partial, objectAlpha] = inspectSlideTextContrast(deck).findings;
  const l = ((0.25 + 0.055) / 1.055) ** 2.4;
  assert.equal(known.status, 'meets-threshold');
  assert.ok(Math.abs(known.ratio - (1.05 / (l + 0.05))) < 1e-12);
  assert.equal(known.ratio, known.nativeRatio);
  assert.deepEqual([unknown.reason, above.reason, partial.reason, objectAlpha.reason], ['non-solid-background', 'complex-scene', 'complex-scene', 'complex-scene']);
  assert.ok([unknown, above, partial, objectAlpha].every(finding => !('ratio' in finding)));
});

test('shared inherited text has per-slide findings with definition identity and visibility', () => {
  const black = [{ runs: [run('Inherited text', { color: { theme: 'accent1' }, fontSize: 12 })] }];
  const deck = createPresentation({ theme: { colors: { accent1: 'FFFFFF' } }, masters: [{
    id: 'master', name: 'Master', theme: { colors: { accent1: '000000' } }, elements: [text('master-text', black)],
  }] });
  deck.defineSlideLayout({ id: 'master-only', name: 'Master only', masterId: 'master' });
  deck.defineSlideLayout({ id: 'layout-only', name: 'Layout only', masterId: 'master', showMasterShapes: false,
    elements: [{ id: 'layout-table', type: 'table', box, columns: [200], rows: [{ cells: [{ content: black }] }] }] });
  deck.defineSlideLayout({ id: 'combined', name: 'Combined', masterId: 'master', elements: [text('layout-text', black)] });
  const add = (id, layoutId, color, extra = {}) => deck.addSlide({ id, layoutId, background: { type: 'solid', color }, ...extra });
  add('white', 'master-only', 'FFFFFF');
  add('black', 'master-only', '000000');
  add('layout', 'layout-only', 'FFFFFF');
  add('complex', 'combined', 'FFFFFF');
  add('suppressed', 'master-only', 'FFFFFF', { showInheritedShapes: false });
  add('hidden', 'master-only', 'FFFFFF', { hidden: true });
  const result = inspectSlideTextContrast(deck);
  assert.deepEqual(result.findings.map(f => [f.slideId, f.elementId, f.source.kind, f.source.id, f.status, f.ratio]), [
    ['white', 'master-text', 'master', 'master', 'meets-threshold', 21],
    ['black', 'master-text', 'master', 'master', 'below-threshold', 1],
    ['layout', 'layout-table', 'layout', 'layout-only', 'meets-threshold', 21],
    ['complex', 'master-text', 'master', 'master', 'indeterminate', undefined],
    ['complex', 'layout-text', 'layout', 'combined', 'indeterminate', undefined],
  ]);
  assert.match(result.findings[0].path, /^masters\.0\.elements\.0\.content\./);
  assert.match(result.findings[2].path, /^layouts\.\d+\.elements\.0\.rows\.0\.cells\.0\.content\./);
  assert.ok(Object.isFrozen(result.findings[0].source));
  assert.deepEqual(inspectSlideTextContrast(deck, { maxFindings: 5 }), result);
  assert.throws(() => inspectSlideTextContrast(deck, { maxFindings: 4 }), error => error instanceof TextContrastLimitError && error.maxFindings === 4);
});

test('contrast reaches a foreground text object through translation-only singleton groups', () => {
  const deck = createPresentation({ size: { width: 200, height: 100 } });
  const group = (id, children, extra = {}) => ({
    id, type: 'group', box: { ...box, x: 12, y: 18 }, coordinateSize: { width: box.width, height: box.height }, children, ...extra,
  });
  const white = [{ runs: [run('Foreground', { color: 'FFFFFF', fontSize: 12 })] }];
  const background = { type: 'solid', color: '000000' };
  const label = id => text(id, white);
  deck.addSlide({ id: 'positive', background, elements: [group('outer', [group('inner', [label('known'), { ...label('hidden'), hidden: true }])])] });
  for (const [id, extra] of [
    ['scale', { coordinateSize: { width: 100, height: 100 } }],
    ['rotate', { transform: { rotation: 15 } }],
    ['flip', { transform: { flipV: true } }],
    ['alpha', { opacity: 0.5 }],
  ]) deck.addSlide({ id, background, elements: [group(`g-${id}`, [label(`text-${id}`)], extra)] });
  deck.addSlide({ id: 'sibling', background, elements: [group('siblings', [
    label('with-sibling'), { type: 'shape', shape: 'rect', box, accessibility: { decorative: true } },
  ])] });
  deck.addSlide({ id: 'shifted-cover', background, elements: [group('shifted', [
    { type: 'shape', shape: 'rect', box, style: { fill: { type: 'solid', color: 'FFFFFF' }, stroke: { paint: { type: 'none' } } }, text: { content: white } },
  ])] });
  const [known, ...unknown] = inspectSlideTextContrast(deck).findings;
  assert.equal(known.status, 'meets-threshold');
  assert.equal(known.ratio, 21);
  assert.equal(known.effectiveFontSize, 12);
  assert.equal(known.path, 'slides.0.elements.0.children.0.children.0.content.0.runs.0');
  assert.deepEqual(unknown.map(f => f.reason), Array(5).fill('complex-scene').concat('unsupported-text-owner'));
  assert.ok(unknown.every(f => !('ratio' in f)));
});
