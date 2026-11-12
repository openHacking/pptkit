import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { DEFAULT_SESSION_LIMITS } from '../dist/import/import-pptx.js';
import { preparePictureCandidate } from '../dist/import/prepare-picture-candidate.js';
import { verifyPictureProjection } from '../dist/import/verify-picture-projection.js';
import { readArchive } from '../dist/archive/read-archive.js';
import { normalizePresentation } from '@pptkit/core';
import { imageFixture, encodeImages, picture, png, validGif as gif } from './image-fixture.mjs';
import { relationships } from './opc-fixture.mjs';

const limits = { ...DEFAULT_SESSION_LIMITS }, checkpoint = async () => {};
const build = (source, binding, bytes = gif, mime = 'image/gif', overrides = {}) => preparePictureCandidate(source, binding, bytes, mime, {}, { ...limits, ...overrides }, checkpoint);

test('picture candidate builder isolates one shared source across three rounds in both dialects', async () => {
  for (const strict of [false, true]) {
    let source = encodeImages(imageFixture(picture({ id: 2, crop: '<a:srcRect l="12500"/>' }) + picture({ id: 3 }), strict));
    for (let round = 0; round < 3; round++) {
      const session = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
      try {
        const [target, sibling] = session.report.pictureSources, before = await readArchive(source, {}, true);
        const candidate = await build(source, target, round % 2 ? png : gif, round % 2 ? 'image/png' : 'image/gif');
        const after = await readArchive(candidate.bytes, {}, true);
        for (const [name, bytes] of before.parts) if (!['deck/z.xml', 'deck/_rels/z.xml.rels', '[Content_Types].xml'].includes(name)) assert.deepEqual(after.parts.get(name), bytes, name);
        const reopened = await importPptx(candidate.bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
        try {
          const [changed, untouched] = reopened.report.pictureSources;
          assert.equal(changed.mediaPart, candidate.part); assert.equal(changed.relationshipId, candidate.relationshipId);
          assert.equal(untouched.mediaPart, sibling.mediaPart); assert.equal(reopened.document.slides[0].elements[0].crop.left, 0.125);
        } finally { reopened.dispose(); }
        source = candidate.bytes;
      } finally { session.dispose(); }
    }
  }
});

test('picture candidate captures inputs and avoids existing part and relationship identities', async () => {
  const fixture = imageFixture(); fixture.parts['ppt/media/PPTKIT-IMAGE-1.gif'] = gif; fixture.mediaTypes['ppt/media/PPTKIT-IMAGE-1.gif'] = 'image/gif';
  fixture.parts['deck/_rels/z.xml.rels'] = relationships([fixture.rel('layout', 'slideLayout', 'layout.xml'), fixture.rel('image', 'image', '../media/raster.bin'), fixture.rel('pptkitImage2', 'image', '../media/raster.bin')]);
  const source = encodeImages(fixture), session = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try {
    const input = Buffer.from(source), replacement = Buffer.from(gif), binding = { ...session.report.pictureSources[0], source: { ...session.report.pictureSources[0].source } };
    const pending = build(input, binding, replacement); input.fill(0); replacement.fill(0); binding.relationshipId = 'stale'; binding.source.part = '/wrong.xml';
    const candidate = await pending;
    assert.equal(candidate.part, '/ppt/media/pptkit-image-3.gif'); assert.equal(candidate.relationshipId, 'pptkitImage3');
    const reopened = await importPptx(candidate.bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try { assert.equal(reopened.report.pictureSources[0].mediaPart, candidate.part); } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test('picture candidate builder rejects stale bindings, opaque effects and resource limits without changing source', async () => {
  const fixture = imageFixture(picture({ id: 2 }) + picture({ id: 3, effects: '<a:extLst/>' })), source = encodeImages(fixture), original = Buffer.from(source);
  const session = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try {
    const [valid, opaque] = session.report.pictureSources, code = expected => error => error.code === expected;
    for (const changed of [{ relationshipId: 'stale' }, { mediaPart: '/missing.png' }, { sourceElementIndex: 1 }]) await assert.rejects(build(source, { ...valid, ...changed }), code('save-candidate-invalid'));
    await assert.rejects(build(source, opaque), code('save-candidate-invalid'));
    await assert.rejects(build(source, valid, png, 'image/gif'), code('save-candidate-invalid'));
    await assert.rejects(build(source, valid, gif, 'image/gif', { maxImportedAssetBytes: 1 }), code('save-limit-exceeded'));
    await assert.rejects(build(source, valid, gif, 'image/gif', { maxSaveAnalysisWork: 1 }), code('save-limit-exceeded'));
    await assert.rejects(build(source, valid, gif, 'image/gif', { maxSaveValidationBytes: 1 }), code('save-limit-exceeded'));
    const controller = new AbortController(); controller.abort();
    await assert.rejects(preparePictureCandidate(source, valid, gif, 'image/gif', { signal: controller.signal }, limits, checkpoint), code('operation-aborted'));
    const during = new AbortController();
    await assert.rejects(preparePictureCandidate(source, valid, gif, 'image/gif', { signal: during.signal }, limits, async () => { during.abort(); }), code('operation-aborted'));
    assert.deepEqual(source, original);
    assert.ok((await build(source, valid)).bytes.length);
  } finally { session.dispose(); }
});

test('picture candidate builder blocks external links and incomplete raw relationship closure', async () => {
  const external = imageFixture(picture({ link: 'outside' }));
  external.parts['deck/_rels/z.xml.rels'] = relationships([external.rel('layout', 'slideLayout', 'layout.xml'), external.rel('image', 'image', '../media/raster.bin'), { ...external.rel('outside', 'image', 'https://example.invalid/image.png'), TargetMode: 'External' }]);
  const missing = imageFixture();
  missing.parts['deck/_rels/z.xml.rels'] = missing.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>', `<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="${missing.R}/image" Target="../media/missing.png"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>`);
  for (const fixture of [external, missing]) {
    const source = encodeImages(fixture), session = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try { await assert.rejects(build(source, session.report.pictureSources[0]), error => error.code === 'save-candidate-invalid'); } finally { session.dispose(); }
  }
});

test('projection verifier rejects changed sibling geometry, source identities and unused candidate assets', async () => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), before = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try {
    const target = before.report.pictureSources[0], candidate = await build(source, target), after = await importPptx(candidate.bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try {
      const left = normalizePresentation(before.document), right = normalizePresentation(after.document), originalSources = before.report.pictureSources, resultSources = after.report.pictureSources;
      const verify = (doc, sources = resultSources) => verifyPictureProjection(left, doc, originalSources, sources, target, candidate.part, candidate.relationshipId, limits.maxSnapshotValues);
      verify(right);
      const wrongBox = structuredClone(right); wrongBox.slides[0].elements[1].box.x += 1; assert.throws(() => verify(wrongBox), error => error.code === 'save-candidate-invalid');
      const wrongSource = structuredClone(resultSources); wrongSource[1].mediaPart = candidate.part; assert.throws(() => verify(right, wrongSource), error => error.code === 'save-candidate-invalid');
      const extra = structuredClone(right); extra.assets.push({ ...structuredClone(extra.assets[0]), id: 'unmapped' }); assert.throws(() => verify(extra), error => error.code === 'save-candidate-invalid');
      const forged = structuredClone(target); forged.sourceElementIndex += 1;
      assert.throws(() => verifyPictureProjection(left, right, originalSources, resultSources, forged, candidate.part, candidate.relationshipId, limits.maxSnapshotValues), error => error.code === 'save-candidate-invalid');
    } finally { after.dispose(); }
  } finally { before.dispose(); }
});

test('picture candidate verification handles nested groups and reordered imported asset IDs', async () => {
  const group = children => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>${children}</p:grpSp>`;
  const fixture = imageFixture(group(picture({ id: 11 })) + picture({ id: 12, embed: 'other' }));
  fixture.parts['media/other.gif'] = gif; fixture.mediaTypes['media/other.gif'] = 'image/gif';
  fixture.parts['deck/_rels/z.xml.rels'] = relationships([fixture.rel('layout', 'slideLayout', 'layout.xml'), fixture.rel('image', 'image', '../media/raster.bin'), fixture.rel('other', 'image', '../media/other.gif')]);
  const source = encodeImages(fixture), before = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
  try {
    const selected = before.report.pictureSources.find(item => item.source.sourceObjectId === '11');
    const candidate = await build(source, selected), reopened = await importPptx(candidate.bytes, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try {
      const [changed, sibling] = reopened.report.pictureSources;
      assert.equal(changed.mediaPart, candidate.part); assert.equal(sibling.mediaPart, '/media/other.gif');
      assert.equal(reopened.document.slides[0].elements[0].children[0].id, changed.elementId);
    } finally { reopened.dispose(); }
  } finally { before.dispose(); }
});
