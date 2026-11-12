import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { importPptx } from '../dist/index.js';
import { importPptxFile } from '../dist/node.js';
import { readXml } from '../dist/xml/read-xml.js';
import { imageFixture, encodeImages, picture } from './image-fixture.mjs';
import { relationships } from './opc-fixture.mjs';

test('picture source inventory binds shared media to distinct raw MCE-selected blips in both dialects', async () => {
  for (const strict of [false, true]) {
    const selected = picture({ id: 2 }), fixture = imageFixture(selected + picture({ id: 3 }), strict);
    fixture.parts['deck/z.xml'] = fixture.parts['deck/z.xml']
      .replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
      .replace(selected, `<mc:AlternateContent><mc:Choice Requires="x">${picture({ id: 99 })}</mc:Choice><mc:Fallback>${selected}</mc:Fallback></mc:AlternateContent>`);
    const bytes = encodeImages(fixture), session = await importPptx(bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try {
      const sources = session.report.pictureSources;
      assert.equal(sources.length, 2); assert.ok(Object.isFrozen(sources));
      assert.notEqual(sources[0].elementId, sources[1].elementId);
      assert.notEqual(sources[0].sourceElementIndex, sources[1].sourceElementIndex);
      assert.equal(sources[0].assetId, sources[1].assetId);
      assert.ok(sources.every(source => source.mediaPart === '/media/raster.bin' && source.relationshipId === 'image' && source.source.part === '/deck/z.xml' && Object.isFrozen(source)));
      const raw = await readXml(Buffer.from(fixture.parts['deck/z.xml']), {}, true), stack = [raw.root], nodes = [];
      while (stack.length) { const node = stack.pop(); nodes.push(node); for (let i = node.children.length - 1; i >= 0; i--) if (node.children[i].kind === 'element') stack.push(node.children[i]); }
      for (const source of sources) {
        const blip = nodes[source.sourceElementIndex - 1];
        assert.equal(blip.local, 'blip');
        assert.equal(blip.attributes.find(attr => attr.local === 'embed' && attr.namespaceUri === fixture.R).value, 'image');
      }
      const hidden = nodes.findIndex(node => node.local === 'cNvPr' && node.attributes.some(attr => attr.local === 'id' && attr.value === '99'));
      assert.ok(hidden >= 0 && sources[0].sourceElementIndex > hidden + 1);
      const saved = await session.prepareSave({ mode: 'no-op' }); assert.equal(saved.status, 'ready');
      assert.deepEqual(Buffer.from((await session.save(saved.plan)).bytes), bytes);
    } finally { session.dispose(); }
    const ordinary = await importPptx(bytes, { projection: 'content' });
    try { assert.equal(ordinary.report.pictureSources, undefined); } finally { ordinary.dispose(); }
  }
});

test('picture profile is explicit and Node file import snapshots it before I/O', async t => {
  const bytes = encodeImages();
  await assert.rejects(importPptx(bytes, { projection: 'structure', pictureSourceProfile: 'embedded-v1' }), TypeError);
  await assert.rejects(importPptx(bytes, { projection: 'content', pictureSourceProfile: 'bad' }), TypeError);
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-pictures-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, 'source.pptx'); await writeFile(file, bytes);
  const options = { projection: 'content', pictureSourceProfile: 'embedded-v1' }, pending = importPptxFile(file, options);
  options.pictureSourceProfile = 'bad';
  const imported = await pending;
  try { assert.equal(imported.session.report.pictureSources.length, 1); } finally { imported.session.dispose(); }
});

test('nested slide pictures are inventoried while unsupported pictures stay omitted', async () => {
  const group = children => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>${children}</p:grpSp>`;
  const fixture = imageFixture(group(picture({ id: 11 })) + picture({ id: 12, preset: 'ellipse' }));
  const session = await importPptx(encodeImages(fixture), { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try {
    assert.equal(session.report.pictureSources.length, 1);
    assert.equal(session.report.pictureSources[0].source.sourceObjectId, '11');
    const parent = session.document.slides[0].elements.find(item => item.type === 'group');
    assert.equal(parent.children[0].id, session.report.pictureSources[0].elementId);
  } finally { session.dispose(); }
  await assert.rejects(importPptx(encodeImages(fixture), { projection: 'content', pictureSourceProfile: 'embedded-v1', sessionLimits: { maxReportFindings: 1 } }), error => error.code === 'import-limit-exceeded');
});

test('picture inventory identifies raw SVG representations and external links without fetching them', async () => {
  for (const strict of [false, true]) {
    const extension = '<a:extLst><a:ext uri="svg"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="svg"/></a:ext></a:extLst>';
    const fixture = imageFixture(picture({ id: 2, link: 'outside', effects: extension }) + picture({ id: 3 }), strict);
    fixture.parts['media/image.svg'] = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>'); fixture.mediaTypes['media/image.svg'] = 'image/svg+xml';
    fixture.parts['deck/_rels/z.xml.rels'] = relationships([fixture.rel('layout', 'slideLayout', 'layout.xml'), fixture.rel('image', 'image', '../media/raster.bin'), fixture.rel('svg', 'image', '../media/image.svg'), { ...fixture.rel('outside', 'image', 'https://example.invalid/image.png'), TargetMode: 'External' }]);
    const session = await importPptx(encodeImages(fixture), { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try {
      const [first, second] = session.report.pictureSources;
      assert.equal(first.hasSvgExtension, true); assert.equal(first.hasExternalLink, true);
      assert.equal(second.hasSvgExtension, false); assert.equal(second.hasExternalLink, false);
    } finally { session.dispose(); }
  }
});

test('picture inventory sees SVG in an inactive raw MCE branch', async () => {
  const extension = '<a:extLst><a:ext uri="svg"><mc:AlternateContent><mc:Choice Requires="x"><v:svgBlip xmlns:v="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="svg"/></mc:Choice><mc:Fallback/></mc:AlternateContent></a:ext></a:extLst>';
  const fixture = imageFixture(picture({ effects: extension }));
  fixture.parts['deck/z.xml'] = fixture.parts['deck/z.xml'].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ');
  fixture.parts['media/image.svg'] = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>'); fixture.mediaTypes['media/image.svg'] = 'image/svg+xml';
  fixture.parts['deck/_rels/z.xml.rels'] = relationships([fixture.rel('layout', 'slideLayout', 'layout.xml'), fixture.rel('image', 'image', '../media/raster.bin'), fixture.rel('svg', 'image', '../media/image.svg')]);
  const session = await importPptx(encodeImages(fixture), { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try { assert.equal(session.report.pictureSources[0].hasSvgExtension, true); } finally { session.dispose(); }
});
