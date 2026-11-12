import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation } from '../packages/core/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('PowerPoint rotated target retains local connection identity and resolved endpoint', async () => {
  const source = await readFile('tests/fixtures/native-connector-rotate-after.pptx');
  const session = await importPptx(source, { projection: 'content' });
  try {
    const [line, left, right] = session.document.slides[0].elements;
    assert.deepEqual(line.start, { elementId: left.id, anchor: 'right' });
    assert.deepEqual(line.end, { elementId: right.id, anchor: 'left' });
    assert.ok(Math.abs(right.transform.rotation - 2727718 / 60000) < 1e-10);
    const resolved = resolveLayout(session.document).slides[0].elements[0];
    assert.ok(Math.abs(resolved.end.x - 7904433 / 12700) < .001);
    assert.ok(Math.abs(resolved.end.y - 2623072 / 12700) < .001);
    const output = await generatePptx(session.document, { connectorMode: 'native' });
    assert.ok(!output.warnings.some(w => w.code === 'native-connector-fallback'));
    const xml = readGeneratedZip(output.bytes).get('ppt/slides/slide1.xml').toString();
    assert.match(xml, /<a:endCxn id="\d+" idx="1"\/>/);
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready');
    assert.ok(Buffer.from((await session.save(plan.plan)).bytes).equals(source));
  } finally { session.dispose(); }
});

test('PowerPoint rerouted attached elbow imports and reexports with target bindings', async () => {
  const source = await readFile('tests/fixtures/native-elbow-move-after.pptx');
  let bytes = source;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content' });
    try {
      const [line, left, right] = session.document.slides[0].elements;
      assert.equal(line.type, 'connector');
      assert.deepEqual(line.start, { elementId: left.id, anchor: 'right' });
      assert.deepEqual(line.end, { elementId: right.id, anchor: 'left' });
      assert.equal(line.route.length, 2);
      assert.ok(Math.abs(line.route[0].x - 465.6521653543307) < .001);
      assert.ok(Math.abs(line.route[1].y - 298.4348031496063) < .001);
      const noOp = await session.prepareSave({ mode: 'no-op' });
      assert.equal(noOp.status, 'ready'); assert.ok(Buffer.from((await session.save(noOp.plan)).bytes).equals(Buffer.from(bytes)));
      const output = await generatePptx(session.document, { connectorMode: 'native' });
      assert.ok(!output.warnings.some(w => w.code === 'native-connector-fallback'));
      assert.match(readGeneratedZip(output.bytes).get('ppt/slides/slide1.xml').toString(), /prst="bentConnector3"/);
      bytes = output.bytes;
    } finally { session.dispose(); }
  }
});

test('native elbow preset geometry matches SVG routes and survives three semantic import/export rounds', async () => {
  for (const [x, y] of [[100, 80], [-100, 80], [100, -80], [-100, -80]]) {
    const deck = createPresentation({ size: { width: 600, height: 400 } });
    const start = { x: 200, y: 180 }, end = { x: 200 + x, y: 180 + y }, bend = { x: end.x, y: start.y };
    deck.addSlide().addElement({ id: 'elbow', type: 'connector', start, end, route: [bend] });
    const native = await generatePptx(deck, { connectorMode: 'native' }), slide = readGeneratedZip(native.bytes).get('ppt/slides/slide1.xml').toString();
    const connector = slide.match(/<p:cxnSp>.*?<\/p:cxnSp>/)[0];
    const match = connector.match(/<a:xfrm([^>]*)><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
    assert.ok(match); assert.match(connector, /prst="bentConnector2"/);
    const [left, top, width, height] = match.slice(2).map(value => Number(value) / 12700);
    const flipH = match[1].includes('flipH="1"'), flipV = match[1].includes('flipV="1"');
    const presetPoints = [[0, 0], [1, 0], [1, 1]].map(([px, py]) => ({ x: left + (flipH ? 1 - px : px) * width, y: top + (flipV ? 1 - py : py) * height }));
    assert.deepEqual(presetPoints, [start, bend, end]);
    const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
    assert.ok(svg.includes(`points="${start.x},${start.y} ${bend.x},${bend.y} ${end.x},${end.y}"`));
    let bytes = native.bytes;
    for (let round = 0; round < 3; round++) {
      const imported = await importPptx(bytes, { projection: 'content' });
      try {
        const [element] = imported.document.slides[0].elements; assert.equal(element.type, 'connector');
        assert.deepEqual(element.start, start); assert.deepEqual(element.route, [bend]); assert.deepEqual(element.end, end);
        const plan = await imported.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual((await imported.save(plan.plan)).bytes, bytes);
        const next = await generatePptx(imported.document, { connectorMode: 'native' }); assert.equal(next.warnings.filter(item => item.code === 'native-connector-fallback').length, 0);
        assert.match(readGeneratedZip(next.bytes).get('ppt/slides/slide1.xml').toString(), /prst="bentConnector2"/); bytes = next.bytes;
      } finally { imported.dispose(); }
    }
  }
});

test('attached native elbows retain target identities over three rounds and diagnose movement without silent rerouting', async () => {
  const deck = createPresentation({ size: { width: 600, height: 400 } }), slide = deck.addSlide();
  slide.addElement({ id: 'edge', name: 'Edge', type: 'connector', start: { elementId: 'left', anchor: 'right' }, end: { elementId: 'right', anchor: 'left' }, route: [{ x: 220, y: 90 }] });
  slide.addElement({ id: 'left', name: 'Left', type: 'shape', shape: 'rect', box: { x: 40, y: 80, width: 40, height: 20 } });
  slide.addElement({ id: 'right', name: 'Right', type: 'shape', shape: 'rect', box: { x: 220, y: 180, width: 40, height: 20 } });
  let bytes = (await generatePptx(deck, { connectorMode: 'native' })).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content' });
    try {
      const page = session.document.slides[0], [line, left, right] = page.elements;
      assert.deepEqual(line.start, { elementId: left.id, anchor: 'right' }); assert.deepEqual(line.end, { elementId: right.id, anchor: 'left' });
      assert.deepEqual(line.route, [{ x: 220, y: 90 }]);
      const before = await session.prepareSave({ mode: 'no-op' }); assert.equal(before.status, 'ready'); assert.deepEqual((await session.save(before.plan)).bytes, bytes);
      const next = await generatePptx(session.document, { connectorMode: 'native' }), xml = readGeneratedZip(next.bytes).get('ppt/slides/slide1.xml').toString();
      assert.equal(next.warnings.filter(warning => warning.code === 'native-connector-fallback').length, 0);
      const ids = Object.fromEntries([...xml.matchAll(/<p:cNvPr id="(\d+)" name="(Left|Right)"/g)].map(match => [match[2], match[1]]));
      assert.ok(xml.includes(`<a:stCxn id="${ids.Left}" idx="3"/>`)); assert.ok(xml.includes(`<a:endCxn id="${ids.Right}" idx="1"/>`));
      assert.match(xml, /prst="bentConnector2"/); bytes = next.bytes;
      if (round === 2) {
        page.replaceElement(right.id, { ...right, box: { ...right.box, x: right.box.x + 20 } });
        const resolved = resolveLayout(session.document).slides[0].elements.find(item => item.type === 'connector');
        assert.deepEqual(resolved.end, { x: 240, y: 190 }); assert.deepEqual(resolved.route, [{ x: 220, y: 90 }]);
        const changed = await generatePptx(session.document, { connectorMode: 'native' });
        assert.ok(changed.warnings.some(warning => warning.code === 'native-connector-fallback'));
        assert.match(readGeneratedZip(changed.bytes).get('ppt/slides/slide1.xml').toString(), /<a:custGeom>/);
        assert.equal((await session.prepareSave({ mode: 'no-op' })).status, 'blocked');
        assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      }
    } finally { session.dispose(); }
  }
});

test('double elbows preserve adjustable waypoint geometry over three fresh rounds', async () => {
  for (const [dx, dy] of [[100, 80], [-100, 80], [100, -80], [-100, -80]]) for (const fraction of [0.25, 0.5, 0.75]) {
    const deck = createPresentation({ size: { width: 600, height: 400 } }), start = { x: 200, y: 180 }, end = { x: 200 + dx, y: 180 + dy }, middle = start.x + dx * fraction;
    const route = [{ x: middle, y: start.y }, { x: middle, y: end.y }];
    deck.addSlide().addElement({ id: 'double', type: 'connector', start, end, route });
    const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
    assert.ok(svg.includes(`points="${start.x},${start.y} ${middle},${start.y} ${middle},${end.y} ${end.x},${end.y}"`));
    let bytes = (await generatePptx(deck, { connectorMode: 'native' })).bytes;
    for (let round = 0; round < 3; round++) {
      const xml = readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();
      assert.match(xml, new RegExp(`<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val ${fraction * 100000}"/>`));
      const session = await importPptx(bytes, { projection: 'content' });
      try {
        const [line] = session.document.slides[0].elements; assert.deepEqual(line.start, start); assert.deepEqual(line.route, route); assert.deepEqual(line.end, end);
        const noOp = await session.prepareSave({ mode: 'no-op' }); assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
        const next = await generatePptx(session.document, { connectorMode: 'native' }); assert.ok(!next.warnings.some(item => item.code === 'native-connector-fallback')); bytes = next.bytes;
      } finally { session.dispose(); }
    }
  }
});

test('browser content import and native re-export preserve an elbow identically to Node', async t => {
  const deck = createPresentation(), slide = deck.addSlide(); slide.addElement({ id: 'elbow', type: 'connector', start: { x: 180, y: 160 }, end: { x: 50, y: 80 }, route: [{ x: 50, y: 160 }] });
  slide.addElement({ id: 'double', type: 'connector', start: { x: 200, y: 160 }, end: { x: 100, y: 80 }, route: [{ x: 175, y: 160 }, { x: 175, y: 80 }] });
  const source = (await generatePptx(deck, { connectorMode: 'native' })).bytes, node = await importPptx(source, { projection: 'content' });
  const expected = { routes: node.document.slides[0].elements.map(item => item.route), bytes: Array.from((await generatePptx(node.document, { connectorMode: 'native' })).bytes) }; node.dispose();
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter', 'pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async data => {
    const session = await PPTKit.importPptx(new Uint8Array(data), { projection: 'content' });
    try { return { routes: session.document.slides[0].elements.map(item => item.route), bytes: Array.from((await PPTKit.generatePptx(session.document, { connectorMode: 'native' })).bytes) }; } finally { session.dispose(); }
  }, Array.from(source));
  assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
});

test('attached double elbows retain bindings and waypoint geometry through fresh exports', async () => {
  const deck = createPresentation({ size: { width: 600, height: 400 } }), slide = deck.addSlide();
  slide.addElement({ id: 'edge', type: 'connector', start: { elementId: 'left', anchor: 'right' }, end: { elementId: 'right', anchor: 'left' }, route: [{ x: 150, y: 90 }, { x: 150, y: 190 }] });
  slide.addElement({ id: 'left', name: 'Left', type: 'shape', shape: 'rect', box: { x: 40, y: 80, width: 40, height: 20 } });
  slide.addElement({ id: 'right', name: 'Right', type: 'shape', shape: 'rect', box: { x: 220, y: 180, width: 40, height: 20 } });
  let bytes = (await generatePptx(deck, { connectorMode: 'native' })).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content' });
    try {
      const [line, left, right] = session.document.slides[0].elements;
      assert.deepEqual(line.start, { elementId: left.id, anchor: 'right' });
      assert.deepEqual(line.end, { elementId: right.id, anchor: 'left' });
      assert.deepEqual(line.route, [{ x: 150, y: 90 }, { x: 150, y: 190 }]);
      const next = await generatePptx(session.document, { connectorMode: 'native' });
      assert.ok(!next.warnings.some(warning => warning.code === 'native-connector-fallback'));
      const xml = readGeneratedZip(next.bytes).get('ppt/slides/slide1.xml').toString();
      assert.match(xml, /prst="bentConnector3"/);
      assert.match(xml, /<a:stCxn id="\d+" idx="3"\/>/);
      assert.match(xml, /<a:endCxn id="\d+" idx="1"\/>/);
      bytes = next.bytes;
    } finally { session.dispose(); }
  }
});
