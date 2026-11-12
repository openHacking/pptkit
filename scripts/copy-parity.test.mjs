import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { inspectOpcPackage } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

function fixture() {
  const p = createPresentation({ size: { width: 400, height: 200 } });
  const s = p.addSlide({ id: "source", notes: [{ runs: [{ text: "Again", action: { type: "slide", slideId: "source" } }] }] });
  s.addElement({ type: "connector", id: "edge", start: { elementId: "left", anchor: "right" }, end: { elementId: "right", anchor: "left" }, route: [{ x: 150, y: 70 }] });
  s.addElement({ type: "shape", id: "left", shape: "rect", name: "Left", box: { x: 20, y: 30, width: 40, height: 20 }, action: { type: "slide", slideId: "source" } });
  s.addElement({ type: "shape", id: "right", shape: "rect", name: "Right", box: { x: 220, y: 30, width: 40, height: 20 } });
  return { p, s };
}

test("cross-document slide copy survives source removal through layout, SVG and PPTX relationships", async () => {
  const { p: source } = fixture(), destination = createPresentation({ size: { width: 400, height: 200 } });
  const { slide } = destination.copySlideFrom(source, "source", { layoutId: null, theme: "destination", geometry: "preserve", customData: "preserve", assets: { deduplicate: "content", external: "reject" } });
  source.removeSlide("source");
  assert.deepEqual(validatePresentation(destination), []);
  const edge = resolveLayout(destination).slides[0].elements[0];
  assert.deepEqual(edge.start, { x: 60, y: 40 }); assert.deepEqual(edge.end, { x: 220, y: 40 });
  assert.equal((await renderPresentationToSvg(destination)).slides[0].slideId, slide.id);
  const opc = await inspectOpcPackage((await generatePptx(destination)).bytes);
  assert.ok(opc.relationships.some(r => r.source === "/ppt/slides/slide1.xml" && r.targetPart === "/ppt/slides/slide1.xml"));
  assert.ok(opc.relationships.some(r => r.source === "/ppt/notesSlides/notesSlide1.xml" && r.targetPart === "/ppt/slides/slide1.xml"));
});

test('fit explicit geometry is shared by Layout, SVG and PPTX after source removal', async () => {
  const { p: source } = fixture(), target = createPresentation({ size: { width: 200, height: 200 } });
  target.copySlideFrom(source, 'source', { layoutId: null, theme: 'destination', geometry: 'fit-explicit-geometry', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
  source.removeSlide('source');
  const edge = resolveLayout(target).slides[0].elements[0];
  assert.deepEqual(edge.start, { x: 30, y: 70 }); assert.deepEqual(edge.end, { x: 110, y: 70 });
  assert.deepEqual(edge.route, [{ x: 75, y: 85 }]);
  assert.match((await renderPresentationToSvg(target)).slides[0].svg, /30,70 75,85 110,70/);
  const xml = readGeneratedZip((await generatePptx(target)).bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:path w="1016000" h="190500"/);
  assert.match(xml, /<a:pt x="571500" y="190500"/);
});

test('fit geometry exposes measured overflow and reports missing measurement', () => {
  const source = createPresentation({ size: { width: 400, height: 200 } }), target = createPresentation({ size: { width: 200, height: 200 } });
  source.addSlide({ id: 'source', elements: [{ id: 'text', type: 'text', content: 'Measured', box: { x: 20, y: 10, width: 160, height: 40 }, frame: { margin: 0 } }] });
  const copied = target.copySlideFrom(source, 'source', { layoutId: null, theme: 'destination', geometry: 'fit-explicit-geometry', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
  const textMeasurement = { measure: () => ({ bounds: { x: 0, y: 0, width: 70, height: 25 } }) }, diagnostics = { textOverflow: true };
  assert.equal(resolveLayout(source, { textMeasurement, diagnostics }).analysis.diagnostics.some(item => item.code === 'text-overflow'), false);
  assert.ok(resolveLayout(target, { textMeasurement, diagnostics }).analysis.diagnostics.some(item => item.code === 'text-overflow' && item.elementId === copied.elementIds.text));
  assert.ok(resolveLayout(target, { diagnostics }).analysis.diagnostics.some(item => item.code === 'text-measurement-unavailable' && item.elementId === copied.elementIds.text));
});

test('cross-document copied image bytes and page correspondence survive three fresh PPTX imports', async () => {
  const { png } = await import('../packages/pptx-parser/test/image-fixture.mjs');
  const { importPptx } = await import('../packages/pptx-parser/dist/index.js');
  const source = createPresentation(), target = createPresentation();
  for (const id of ['a', 'b']) source.registerAsset({ id, kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 });
  source.addSlide({ id: 'origin', notes: 'Copied notes', elements: [
    { id: 'text', type: 'text', box: { x: 20, y: 20, width: 180, height: 30 }, content: 'Copied page' },
    { id: 'a-image', type: 'image', assetId: 'a', box: { x: 20, y: 70, width: 100, height: 50 } },
    { id: 'b-image', type: 'image', assetId: 'b', box: { x: 140, y: 70, width: 100, height: 50 } },
  ] });
  const result = target.copySlideFrom(source, 'origin', { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
  assert.equal(result.assets[0].destinationId, result.assets[1].destinationId);
  source.removeSlide('origin');
  let bytes = (await generatePptx(target)).bytes;
  for (let round = 0; round < 3; round++) {
    const archive = readGeneratedZip(bytes), slide = archive.get('ppt/slides/slide1.xml').toString(), notes = archive.get('ppt/notesSlides/notesSlide1.xml').toString();
    assert.match(slide, /Copied page/); assert.match(notes, /Copied notes/);
    const media = [...archive.entries()].filter(([name]) => name.startsWith('ppt/media/'));
    assert.equal(media.length, 1); assert.deepEqual(new Uint8Array(media[0][1]), new Uint8Array(png));
    const opc = await inspectOpcPackage(bytes);
    assert.ok(opc.relationships.some(r => r.source === '/ppt/notesSlides/notesSlide1.xml' && r.targetPart === '/ppt/slides/slide1.xml'));
    const session = await importPptx(bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try {
      assert.equal(session.report.pictureSources.length, 2);
      assert.equal(session.report.pictureSources[0].mediaPart, session.report.pictureSources[1].mediaPart);
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
});

test('copied layout artwork and placeholders feed destination SVG and PPTX', async () => {
  const { png } = await import('../packages/pptx-parser/test/image-fixture.mjs');
  const source = createPresentation(), target = createPresentation();
  source.registerAsset({ id: 'art', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png }, width: 8, height: 4 });
  source.defineSlideLayout({ id: 'source-layout', name: 'Source', placeholders: [{ key: 'body', kind: 'body', box: { x: 40, y: 40, width: 300, height: 50 } }], elements: [{ id: 'picture', type: 'image', assetId: 'art', box: { x: 40, y: 120, width: 200, height: 100 } }] });
  target.copySlideLayoutFrom(source, 'source-layout', { id: 'destination-layout', masterId: null, theme: 'destination', assets: { deduplicate: 'content', external: 'reject' } });
  target.addSlide({ layoutId: 'destination-layout', elements: [{ id: 'body', type: 'text', placeholderKey: 'body', content: 'Destination content' }] });
  source.replaceAsset('art', { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  assert.deepEqual(validatePresentation(target), []);
  const rendered = await renderPresentationToSvg(target); assert.match(rendered.slides[0].svg, /Destination content/); assert.match(rendered.slides[0].svg, /data:image\/png;base64/);
  const generated = await generatePptx(target), archive = readGeneratedZip(generated.bytes);
  assert.equal([...archive.keys()].filter(name => name.startsWith('ppt/media/')).length, 1);
  assert.ok([...archive.entries()].some(([name, data]) => name.startsWith('ppt/slideLayouts/slideLayout') && name.endsWith('.xml') && data.toString().includes('<p:pic>')));
  const opc = await inspectOpcPackage(generated.bytes);
  assert.ok(opc.relationships.some(r => r.source === '/ppt/slides/slide1.xml' && r.targetPart?.startsWith('/ppt/slideLayouts/')));
});

test('copied master and layout remain connected through destination SVG and PPTX', async () => {
  const { png } = await import('../packages/pptx-parser/test/image-fixture.mjs');
  const source = createPresentation({ theme: { colors: { accent2: '245678' } } }), target = createPresentation({ theme: { colors: { accent2: 'FF00FF' } } });
  source.registerAsset({ id: 'art', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: png }, width: 8, height: 4 });
  source.defineSlideMaster({ id: 'source-master', name: 'Master', theme: { colors: { accent1: 'FF0000' } }, elements: [{ id: 'picture', type: 'image', assetId: 'art', box: { x: 40, y: 120, width: 200, height: 100 } }, { id: 'accent', type: 'shape', shape: 'rect', box: { x: 300, y: 120, width: 50, height: 50 }, style: { fill: { type: 'solid', color: { theme: 'accent2' } } } }] });
  source.defineSlideLayout({ id: 'source-layout', name: 'Layout', masterId: 'source-master', placeholders: [{ key: 'body', kind: 'body', box: { x: 40, y: 40, width: 300, height: 50 } }] });
  const assets = { deduplicate: 'content', external: 'reject' };
  target.copySlideMasterFrom(source, 'source-master', { id: 'target-master', theme: 'source-effective', assets });
  target.copySlideLayoutFrom(source, 'source-layout', { id: 'target-layout', masterId: 'target-master', theme: 'destination', assets });
  target.addSlide({ layoutId: 'target-layout', elements: [{ id: 'body', type: 'text', placeholderKey: 'body', content: 'Master control' }] });
  assert.deepEqual(validatePresentation(target), []);
  const svg = (await renderPresentationToSvg(target)).slides[0].svg;
  assert.match(svg, /data:image\/png;base64/); assert.match(svg, /#245678/i);
  const archive = readGeneratedZip((await generatePptx(target)).bytes);
  assert.ok([...archive.entries()].some(([name, data]) => name.startsWith('ppt/slideMasters/slideMaster') && name.endsWith('.xml') && data.toString().includes('<p:pic>')));
  assert.equal([...archive.keys()].filter(name => name.startsWith('ppt/media/')).length, 1);
  assert.ok([...archive.entries()].some(([name, data]) => name.startsWith('ppt/theme/') && name.endsWith('.xml') && data.toString().includes('245678')));
});

test('all cross-document master theme policies reach SVG and generated PPTX themes', async () => {
  const source = createPresentation({ theme: { colors: { accent1: '112233', accent2: '445566' } } });
  source.defineSlideMaster({ id: 'm', name: 'Master', theme: { colors: { accent1: 'ABCDEF' } }, elements: [
    { id: 'first', type: 'shape', shape: 'rect', box: { x: 10, y: 10, width: 40, height: 40 }, style: { fill: { type: 'solid', color: { theme: 'accent1' } } } },
    { id: 'second', type: 'shape', shape: 'rect', box: { x: 60, y: 10, width: 40, height: 40 }, style: { fill: { type: 'solid', color: { theme: 'accent2' } } } },
  ] });
  source.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm' });
  const assets = { deduplicate: 'content', external: 'reject' };
  for (const [theme, colors] of [
    ['destination', ['010203', 'A0B0C0']],
    ['source-declarations', ['ABCDEF', 'A0B0C0']],
    ['source-effective', ['ABCDEF', '445566']],
  ]) {
    const target = createPresentation({ theme: { colors: { accent1: '010203', accent2: 'A0B0C0' } } });
    target.copySlideMasterFrom(source, 'm', { id: 'copied-master', theme, assets });
    target.copySlideLayoutFrom(source, 'l', { id: 'copied-layout', masterId: 'copied-master', theme: 'destination', assets });
    target.addSlide({ layoutId: 'copied-layout' });
    const svg = (await renderPresentationToSvg(target)).slides[0].svg.toUpperCase();
    for (const color of colors) assert.ok(svg.includes('#' + color), `${theme} SVG lacks ${color}`);
    const archive = readGeneratedZip((await generatePptx(target)).bytes);
    const themes = [...archive.entries()].filter(([name]) => /^ppt\/theme\/theme\d+\.xml$/.test(name)).map(([, data]) => data.toString().toUpperCase());
    assert.ok(themes.some(xml => colors.every(color => xml.includes(`VAL="${color}"`))), `${theme} PPTX lacks expected theme colors`);
  }
});

test('reused copied layout stays shared by two exported slides', async () => {
  const source = createPresentation(), target = createPresentation();
  source.defineSlideLayout({ id: 'layout', name: 'Shared', elements: [{ id: 'decoration', type: 'shape', shape: 'rect', box: { x: 10, y: 10, width: 40, height: 40 }, style: { fill: { type: 'solid', color: '123456' } } }] });
  const assets = { deduplicate: 'content', external: 'reject' };
  const first = target.copySlideLayoutFrom(source, 'layout', { id: 'first', masterId: null, theme: 'destination', assets });
  const second = target.copySlideLayoutFrom(source, 'layout', { id: 'second', masterId: null, theme: 'destination', assets, reuseEquivalent: true });
  assert.equal(second.layout.id, first.layout.id);
  target.addSlide({ layoutId: first.layout.id });
  target.addSlide({ layoutId: second.layout.id });
  const svg = (await renderPresentationToSvg(target)).slides;
  assert.equal(svg.length, 2);
  assert.ok(svg.every(slide => slide.svg.toUpperCase().includes('#123456')));
  const opc = await inspectOpcPackage((await generatePptx(target)).bytes);
  const links = [1, 2].map(i => opc.relationships.filter(r => r.source === `/ppt/slides/slide${i}.xml` && r.type.endsWith('/slideLayout')));
  assert.ok(links.every(items => items.length === 1));
  assert.equal(links[0][0].targetPart, links[1][0].targetPart);
  target.replaceSlideLayout(first.layout.id, { ...first.layout, elements: first.layout.elements.map(element => ({ ...element, style: { fill: { type: 'solid', color: 'ABCDEF' } } })) });
  const edited = (await renderPresentationToSvg(target)).slides;
  assert.ok(edited.every(slide => slide.svg.toUpperCase().includes('#ABCDEF')));
  assert.equal(source.layouts[0].elements[0].style.fill.color, '123456');
});

test('reused copied master and layout keep their shared PPTX ownership chain', async () => {
  const source = createPresentation(), target = createPresentation(), assets = { deduplicate: 'content', external: 'reject' };
  source.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ id: 'shape', type: 'shape', shape: 'rect', box: { x: 5, y: 5, width: 40, height: 40 }, style: { fill: { type: 'solid', color: '123456' } } }] });
  source.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm' });
  const first = target.copySlideMasterFrom(source, 'm', { id: 'first', theme: 'source-effective', assets });
  const second = target.copySlideMasterFrom(source, 'm', { id: 'second', theme: 'source-effective', assets, reuseEquivalent: true });
  assert.equal(second.master.id, first.master.id);
  const layout = target.copySlideLayoutFrom(source, 'l', { id: 'layout', masterId: second.master.id, theme: 'destination', assets });
  target.addSlide({ layoutId: layout.layout.id }); target.addSlide({ layoutId: layout.layout.id });
  const slides = (await renderPresentationToSvg(target)).slides;
  assert.ok(slides.every(slide => slide.svg.toUpperCase().includes('#123456')));
  const opc = await inspectOpcPackage((await generatePptx(target)).bytes);
  const layoutLinks = [1, 2].map(i => opc.relationships.filter(r => r.source === `/ppt/slides/slide${i}.xml` && r.type.endsWith('/slideLayout')));
  assert.ok(layoutLinks.every(items => items.length === 1));
  assert.equal(layoutLinks[0][0].targetPart, layoutLinks[1][0].targetPart);
  const masterLinks = opc.relationships.filter(r => r.source === layoutLinks[0][0].targetPart && r.type.endsWith('/slideMaster'));
  assert.equal(masterLinks.length, 1);
});

test('editing a reused master updates its distinct destination layouts without changing the source', async () => {
  const source = createPresentation(), target = createPresentation(), assets = { deduplicate: 'content', external: 'reject' };
  source.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ id: 'shape', type: 'shape', shape: 'rect', box: { x: 5, y: 5, width: 40, height: 40 }, style: { fill: { type: 'solid', color: '123456' } } }] });
  source.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm' });
  const first = target.copySlideMasterFrom(source, 'm', { id: 'master', theme: 'destination', assets });
  const second = target.copySlideMasterFrom(source, 'm', { id: 'unused', theme: 'destination', assets, reuseEquivalent: true });
  assert.equal(second.master, first.master);
  const a = target.copySlideLayoutFrom(source, 'l', { id: 'layout-a', masterId: first.master.id, theme: 'destination', assets });
  const b = target.copySlideLayoutFrom(source, 'l', { id: 'layout-b', masterId: second.master.id, theme: 'destination', assets });
  target.addSlide({ layoutId: a.layout.id }); target.addSlide({ layoutId: b.layout.id });
  target.replaceSlideMaster(first.master.id, { ...first.master, elements: first.master.elements.map(element => ({ ...element, style: { fill: { type: 'solid', color: 'ABCDEF' } } })) });
  const slides = (await renderPresentationToSvg(target)).slides;
  assert.ok(slides.every(slide => slide.svg.toUpperCase().includes('#ABCDEF')));
  assert.equal(source.masters[0].elements[0].style.fill.color, '123456');
  const opc = await inspectOpcPackage((await generatePptx(target)).bytes);
  const layouts = [1, 2].map(i => opc.relationships.find(r => r.source === `/ppt/slides/slide${i}.xml` && r.type.endsWith('/slideLayout'))?.targetPart);
  assert.ok(layouts[0] && layouts[1] && layouts[0] !== layouts[1]);
  const masters = layouts.map(layout => opc.relationships.filter(r => r.source === layout && r.type.endsWith('/slideMaster')));
  assert.ok(masters.every(items => items.length === 1));
  assert.equal(masters[0][0].targetPart, masters[1][0].targetPart);
});

test('copied external connector mappings resolve inherited master and layout geometry', async () => {
  const source = createPresentation(), target = createPresentation(), assets = { deduplicate: 'content', external: 'reject' };
  const box = { x: 10, y: 20, width: 40, height: 20 };
  source.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ id: 'master-shape', type: 'shape', shape: 'rect', box }] });
  source.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', elements: [{ id: 'layout-edge', type: 'connector', start: { elementId: 'master-shape', anchor: 'right' }, end: { x: 200, y: 30 } }, { id: 'layout-shape', type: 'shape', shape: 'rect', box: { x: 100, y: 40, width: 40, height: 20 } }] });
  source.addSlide({ id: 's', layoutId: 'l', elements: [{ id: 'slide-edge', type: 'connector', start: { elementId: 'layout-shape', anchor: 'right' }, end: { x: 200, y: 50 } }] });
  const master = target.copySlideMasterFrom(source, 'm', { id: 'm2', theme: 'source-effective', assets });
  const layout = target.copySlideLayoutFrom(source, 'l', { id: 'l2', masterId: 'm2', theme: 'destination', assets, connectorTargets: { 'master-shape': master.elementIds['master-shape'] } });
  target.copySlideFrom(source, 's', { layoutId: 'l2', theme: 'destination', geometry: 'preserve', customData: 'preserve', assets, connectorTargets: { 'layout-shape': layout.elementIds['layout-shape'] } });
  const resolved = resolveLayout(target);
  assert.deepEqual(resolved.layouts.find(item => item.id === 'l2').elements[0].start, { x: 50, y: 30 });
  assert.deepEqual(resolved.slides[0].elements[0].start, { x: 140, y: 50 });
  assert.ok((await generatePptx(target)).bytes.length > 0);
});

test('cross-scope connectors are rejected by layout, SVG and PPTX rather than rendered at the origin', async () => {
  const document = createPresentation(), box = { x: 0, y: 0, width: 100, height: 100 };
  document.addSlide({ elements: [
    { id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 100 }, children: [{ id: 'nested', type: 'shape', shape: 'rect', box }] },
    { id: 'edge', type: 'connector', start: { elementId: 'nested' }, end: { x: 200, y: 20 } },
  ] });
  assert.throws(() => resolveLayout(document), /validation/);
  await assert.rejects(renderPresentationToSvg(document), /validation/);
  await assert.rejects(generatePptx(document), /validation/);
});

test('batch-copied page and notes links resolve to the copied pages in PPTX', async () => {
  const source = createPresentation(), target = createPresentation(), box = { x: 20, y: 20, width: 200, height: 50 };
  source.addSlide({ id: 'a', notes: [{ runs: [{ text: 'Go B', action: { type: 'slide', slideId: 'b' } }] }], elements: [{ id: 'a-text', type: 'text', content: 'A', box, action: { type: 'slide', slideId: 'b' } }] });
  source.addSlide({ id: 'b', elements: [{ id: 'b-text', type: 'text', content: 'B', box, action: { type: 'slide', slideId: 'a' } }] });
  target.copySlidesFrom(source, ['b', 'a'], { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
  source.removeSlide('a'); source.removeSlide('b');
  const bytes = (await generatePptx(target)).bytes, opc = await inspectOpcPackage(bytes);
  const link = (from, to) => opc.relationships.some(r => r.source === from && r.type.endsWith('/slide') && r.targetPart === to);
  assert.ok(link('/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml'));
  assert.ok(link('/ppt/slides/slide2.xml', '/ppt/slides/slide1.xml'));
  assert.ok(link('/ppt/notesSlides/notesSlide2.xml', '/ppt/slides/slide1.xml'));
  assert.deepEqual((await renderPresentationToSvg(target)).slides.map(slide => slide.slideId), ['slide-1', 'slide-2']);
});

test('per-page batch mappings use different target layouts and preserve mutual links in output', async () => {
  const source = createPresentation({ textStylePresets: { old: { run: { fontSize: 18 } } } }), target = createPresentation({ textStylePresets: { first: { run: { fontSize: 20 } }, second: { run: { fontSize: 24 } } } });
  const box = { x: 20, y: 20, width: 200, height: 50 };
  source.defineSlideLayout({ id: 'source-layout', name: 'Source', placeholders: [{ key: 'title', kind: 'title', index: 0, box }] });
  for (const [id, other] of [['a', 'b'], ['b', 'a']]) source.addSlide({ id, layoutId: 'source-layout', elements: [{ id: `${id}-text`, type: 'text', box, placeholderKey: 'title', textStylePreset: 'old', content: id.toUpperCase(), action: { type: 'slide', slideId: other } }] });
  for (const id of ['a', 'b']) target.defineSlideLayout({ id: `layout-${id}`, name: id, placeholders: [{ key: `title-${id}`, kind: 'title', index: 0, box }] });
  target.copySlidesFrom(source, ['a', 'b'], { layoutId: null, layouts: { a: 'layout-a', b: 'layout-b' }, theme: 'destination', geometry: 'preserve', customData: 'preserve', placeholderRoles: 'require-match-and-index', placeholderMatching: 'explicit-index', assets: { deduplicate: 'content', external: 'reject' }, perSlide: { a: { textPresets: { old: 'first' } }, b: { textPresets: { old: 'second' } } } });
  const bytes = (await generatePptx(target)).bytes, archive = readGeneratedZip(bytes), opc = await inspectOpcPackage(bytes);
  assert.match(archive.get('ppt/slides/slide1.xml').toString(), /sz="2000"/);
  assert.match(archive.get('ppt/slides/slide2.xml').toString(), /sz="2400"/);
  for (const number of [1, 2]) {
    const part = `/ppt/slides/slide${number}.xml`;
    const layout = opc.relationships.find(r => r.source === part && r.type.endsWith('/slideLayout'))?.targetPart;
    assert.ok(layout);
    assert.match(archive.get(part.slice(1)).toString(), /<p:ph type="title" idx="0"\/>/);
    assert.match(archive.get(layout.slice(1)).toString(), /<p:ph type="title" idx="0"\/>/);
  }
  assert.ok(opc.relationships.some(r => r.source === '/ppt/slides/slide1.xml' && r.targetPart === '/ppt/slides/slide2.xml'));
  assert.ok(opc.relationships.some(r => r.source === '/ppt/slides/slide2.xml' && r.targetPart === '/ppt/slides/slide1.xml'));
  assert.deepEqual((await renderPresentationToSvg(target)).slides.map(slide => slide.slideId), ['slide-1', 'slide-2']);
});

test("copied slide connector geometry stays correct after source removal through Layout, SVG and PPTX", async () => {
  const { p, s } = fixture(), original = resolveLayout(p).slides[0].elements[0];
  const copy = p.duplicateSlide(s.id); p.removeSlide(s.id);
  assert.deepEqual(validatePresentation(p), []);
  const resolved = resolveLayout(p), edge = resolved.slides[0].elements[0];
  assert.deepEqual(edge.start, { x: 60, y: 40 }); assert.deepEqual(edge.end, { x: 220, y: 40 }); assert.deepEqual(edge.route, [{ x: 150, y: 70 }]);
  assert.deepEqual(edge.box, original.box);
  const before = normalizePresentation(p), rendered = await renderPresentationToSvg(p), generated = await generatePptx(p);
  assert.deepEqual(generated.warnings, []); assert.ok(rendered.slides[0].svg.includes("60,40 150,70 220,40"));
  const entries = readGeneratedZip(generated.bytes), slideXml = entries.get("ppt/slides/slide1.xml").toString();
  assert.match(slideXml, /<a:path w="2032000" h="381000"/);
  assert.match(slideXml, /<a:pt x="1143000" y="381000"/);
  const opc = await inspectOpcPackage(generated.bytes);
  assert.ok(opc.relationships.some(r => r.source === "/ppt/slides/slide1.xml" && r.targetPart === "/ppt/slides/slide1.xml"));
  const notesXml = entries.get("ppt/notesSlides/notesSlide1.xml").toString();
  const noteActionId = /<a:hlinkClick r:id="([^"]+)"/.exec(notesXml)?.[1];
  assert.ok(noteActionId, "notes text must emit an action, not merely its ordinary slide backlink");
  assert.ok(opc.relationships.some(r => r.source === "/ppt/notesSlides/notesSlide1.xml" && r.id === noteActionId && r.targetPart === "/ppt/slides/slide1.xml"));
  assert.equal(rendered.slides[0].slideId, copy.id); assert.deepEqual(normalizePresentation(p), before);
});

test("page reordering resolves copied self-links independently from the original page", async () => {
  const { p, s } = fixture(), copy = p.duplicateSlide(s.id, 0);
  const opc = await inspectOpcPackage((await generatePptx(p)).bytes);
  const links = opc.relationships.filter(r => r.type.endsWith("/slide") && r.source.startsWith("/ppt/slides/"));
  assert.deepEqual(links.map(r => [r.source, r.targetPart]), [["/ppt/slides/slide1.xml", "/ppt/slides/slide1.xml"], ["/ppt/slides/slide2.xml", "/ppt/slides/slide2.xml"]]);
  assert.equal(copy.elements[1].action.slideId, copy.id); assert.equal(s.elements[1].action.slideId, s.id);
});

test("copied group geometry resolves only to its own descendants after original group removal", async () => {
  const p = createPresentation(), s = p.addSlide();
  const original = s.addElement({ type: "group", id: "group", box: { x: 100, y: 50, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, children: [
    { type: "shape", id: "child", shape: "rect", box: { x: 10, y: 20, width: 40, height: 20 } },
    { type: "connector", id: "edge", start: { elementId: "child", anchor: "right" }, end: { x: 180, y: 80 } },
  ] });
  const copy = s.duplicateElement(original.id); s.removeElement(original.id);
  assert.equal(copy.children[1].start.elementId, copy.children[0].id);
  const edge = resolveLayout(p).slides[0].elements[0].children[1];
  assert.deepEqual(edge.start, { x: 50, y: 30 }); assert.deepEqual(edge.end, { x: 180, y: 80 });
  assert.deepEqual((await generatePptx(p)).warnings, []);
  assert.ok((await renderPresentationToSvg(p)).slides[0].svg.includes("50,30 180,80"));
});

test('atomic asset batch deduplication survives preview, export and content import', async () => {
  const { png } = await import('../packages/pptx-parser/test/image-fixture.mjs');
  const { importPptx } = await import('../packages/pptx-parser/dist/index.js');
  const p = createPresentation({ id: 'batch', size: { width: 400, height: 200 } });
  const input = { kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 };
  const [first, second] = p.registerAssets([{ asset: input, options: { deduplicate: 'content' } }, { asset: input, options: { deduplicate: 'content' } }]); assert.equal(first.id, second.id);
  const slide = p.addSlide({ id: 'page', elements: [{ id: 'left', type: 'image', assetId: first.id, box: { x: 10, y: 10, width: 100, height: 50 } }, { id: 'right', type: 'image', assetId: second.id, box: { x: 200, y: 10, width: 100, height: 50 } }] });
  assert.throws(() => p.registerAssets([{ asset: { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } } }, { asset: { id: first.id, ...input, width: 9 } }]));
  assert.equal(p.assets.length, 1); assert.equal(slide.elements.length, 2);
  const preview = await renderPresentationToSvg(p); assert.equal((preview.slides[0].svg.match(/data:image\/png;base64,/g) ?? []).length, 2);
  const output = await generatePptx(p), archive = readGeneratedZip(output.bytes); assert.equal([...archive.keys()].filter(name => name.startsWith('ppt/media/')).length, 1);
  const reopened = await importPptx(output.bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try { assert.equal(reopened.report.pictureSources.length, 2); assert.equal(reopened.report.pictureSources[0].mediaPart, reopened.report.pictureSources[1].mediaPart); } finally { reopened.dispose(); }
});

test('cross-document asset mapping feeds destination preview and editable export without changing source', async () => {
  const { png, validGif } = await import('../packages/pptx-parser/test/image-fixture.mjs');
  const { importPptx } = await import('../packages/pptx-parser/dist/index.js');
  const source = createPresentation({ id: 'origin' }), target = createPresentation({ id: 'destination' });
  source.registerAssets([{ asset: { id: 'first', kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 } }, { asset: { id: 'second', kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 } }]);
  const sourceBefore = normalizePresentation(source), mappings = target.copyAssetsFrom(source, ['first', 'second'], { deduplicate: 'content', external: 'reject' });
  assert.equal(mappings[0].destinationId, mappings[1].destinationId); assert.deepEqual(normalizePresentation(source), sourceBefore);
  target.addSlide({ id: 'target-page', elements: mappings.map((item, index) => ({ id: `img-${index}`, type: 'image', assetId: item.destinationId, box: { x: 20 + index * 120, y: 20, width: 100, height: 50 } })) });
  source.replaceAsset('first', { kind: 'image', source: { type: 'bytes', value: validGif }, mimeType: 'image/gif', width: 1, height: 1 });
  const svg = await renderPresentationToSvg(target); assert.equal((svg.slides[0].svg.match(/data:image\/png;base64,/g) ?? []).length, 2);
  const generated = await generatePptx(target), archive = readGeneratedZip(generated.bytes);
  assert.equal([...archive.keys()].filter(name => name.startsWith('ppt/media/')).length, 1);
  const reopened = await importPptx(generated.bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try { assert.equal(reopened.report.pictureSources.length, 2); assert.equal(reopened.report.pictureSources[0].mediaPart, reopened.report.pictureSources[1].mediaPart); } finally { reopened.dispose(); }
});
