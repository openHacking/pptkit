import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { readArchive } from '../dist/archive/read-archive.js';
import { imageFixture, encodeImages, picture, png, validGif as gif } from './image-fixture.mjs';
import { relationships } from './opc-fixture.mjs';
import { importPptxFile } from '../dist/node.js';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';

const profile = { projection: 'content', pictureEditSaveProfile: 'replace-v1' };
function edit(session, id, bytes, mimeType, width, height) {
  const slide = session.document.slides[0], picture = slide.elements.find(element => element.type === 'image');
  session.document.registerAsset({ id, kind: 'image', source: { type: 'bytes', value: bytes }, mimeType, width, height });
  slide.replaceElement(picture.id, { ...picture, assetId: id });
  return { slide, picture };
}

test('public picture save isolates shared media and preserves crop over three rounds in both dialects', async () => {
  for (const strict of [false, true]) {
    let source = encodeImages(imageFixture(picture({ id: 2, crop: '<a:srcRect l="12500"/>' }) + picture({ id: 3 }), strict));
    for (let round = 0; round < 3; round++) {
      const session = await importPptx(source, profile), original = await readArchive(source, {}, true);
      try {
        const bytes = round % 2 ? png : gif, mime = round % 2 ? 'image/png' : 'image/gif';
        edit(session, `new-${round}`, bytes, mime, round % 2 ? 8 : 1, round % 2 ? 4 : 1);
        assert.equal((await session.prepareSave({ mode: 'no-op' })).status, 'blocked');
        const prepared = await session.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
        assert.equal(prepared.plan.report.assurance, 'scoped-preservation'); assert.deepEqual(prepared.plan.report.changes.map(change => change.kind), ['asset', 'asset']);
        const saved = await session.save(prepared.plan), output = await readArchive(saved.bytes, {}, true);
        for (const [name, payload] of original.parts) if (!['deck/z.xml', 'deck/_rels/z.xml.rels', '[Content_Types].xml'].includes(name)) assert.deepEqual(output.parts.get(name), payload, name);
        const reopened = await importPptx(saved.bytes, profile);
        try {
          const [changed, sibling] = reopened.report.pictureSources;
          assert.notEqual(changed.mediaPart, sibling.mediaPart); assert.equal(sibling.mediaPart, '/media/raster.bin');
          assert.equal(reopened.document.slides[0].elements[0].crop.left, 0.125);
          assert.deepEqual(output.parts.get(changed.mediaPart.slice(1)), new Uint8Array(bytes));
        } finally { reopened.dispose(); }
        source = saved.bytes;
      } finally { session.dispose(); }
    }
  }
});

test('picture save blocks unsupported state, source hazards and stale plans without mutating source', async () => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), original = Buffer.from(source), session = await importPptx(source, profile);
  try {
    const { slide, picture } = edit(session, 'new', gif, 'image/gif', 1, 1);
    const changed = slide.elements.find(element => element.id === picture.id);
    slide.replaceElement(picture.id, { ...changed, box: { ...changed.box, x: changed.box.x + 1 } });
    assert.equal((await session.prepareSave({ mode: 'edit' })).report.diagnostics[0].code, 'save-picture-edit-unsupported');
    slide.replaceElement(picture.id, changed);
    const ready = await session.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready');
    slide.replaceElement(picture.id, { ...changed, box: { ...changed.box, x: changed.box.x + 1 } });
    await assert.rejects(session.save(ready.plan), error => error.code === 'save-snapshot-stale');
    slide.replaceElement(picture.id, changed); assert.ok((await session.save(ready.plan)).bytes.length);
    ready.plan.dispose(); await assert.rejects(session.save(ready.plan), error => error.code === 'plan-disposed');
    assert.deepEqual(source, original);
  } finally { session.dispose(); }
  const fixture = imageFixture(picture({ effects: '<a:extLst/>' })), blocked = await importPptx(encodeImages(fixture), profile);
  try { edit(blocked, 'new', gif, 'image/gif', 1, 1); assert.equal((await blocked.prepareSave({ mode: 'edit' })).report.diagnostics[0].code, 'save-picture-candidate-unsupported'); } finally { blocked.dispose(); }
  const external = imageFixture(picture({ link: 'outside' }));
  external.parts['deck/_rels/z.xml.rels'] = relationships([external.rel('layout', 'slideLayout', 'layout.xml'), external.rel('image', 'image', '../media/raster.bin'), { ...external.rel('outside', 'image', 'https://example.invalid/a'), TargetMode: 'External' }]);
  const linked = await importPptx(encodeImages(external), profile);
  try { edit(linked, 'new', gif, 'image/gif', 1, 1); assert.equal((await linked.prepareSave({ mode: 'edit' })).report.diagnostics[0].code, 'save-picture-edit-unsupported'); } finally { linked.dispose(); }
});

test('picture save profile is opt-in and rejects incompatible modes and profiles', async () => {
  const source = encodeImages();
  for (const options of [{ projection: 'structure', pictureEditSaveProfile: 'replace-v1' }, { projection: 'content', pictureEditSaveProfile: 'other' }, { ...profile, editSaveProfile: 'text-runs-v1' }, { ...profile, masterArtwork: 'pictures-v1' }, { ...profile, chartEditSaveProfile: 'existing-values-v1' }]) await assert.rejects(importPptx(source, options), TypeError);
  const session = await importPptx(source, { projection: 'content' });
  try { edit(session, 'new', gif, 'image/gif', 1, 1); assert.equal((await session.prepareSave({ mode: 'edit' })).report.diagnostics[0].code, 'save-mode-unsupported'); } finally { session.dispose(); }
});

test('picture save rejects metadata mismatch, multiple changes, budget exhaustion and cancellation', async () => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 })));
  const wrong = await importPptx(source, profile);
  try { edit(wrong, 'new', gif, 'image/gif', 8, 4); assert.equal((await wrong.prepareSave({ mode: 'edit' })).report.diagnostics[0].code, 'save-picture-edit-unsupported'); } finally { wrong.dispose(); }
  const mixed = await importPptx(source, profile);
  try {
    edit(mixed, 'new', gif, 'image/gif', 1, 1);
    const slide = mixed.document.slides[0], sibling = slide.elements[1]; slide.replaceElement(sibling.id, { ...sibling, assetId: 'new' });
    assert.equal((await mixed.prepareSave({ mode: 'edit' })).report.diagnostics[0].code, 'save-picture-edit-unsupported');
  } finally { mixed.dispose(); }
  const limited = await importPptx(source, { ...profile, sessionLimits: { maxSaveValidationBytes: 1 } });
  try {
    edit(limited, 'new', gif, 'image/gif', 1, 1);
    await assert.rejects(limited.prepareSave({ mode: 'edit' }), error => error.code === 'save-limit-exceeded');
    const controller = new AbortController(); controller.abort();
    await assert.rejects(limited.prepareSave({ mode: 'edit', signal: controller.signal }), error => error.code === 'operation-aborted');
  } finally { limited.dispose(); }
});

test('Node file import snapshots picture save profile before asynchronous I/O', async t => {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-picture-save-')); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, 'source.pptx'); await writeFile(file, encodeImages());
  const options = { ...profile }, pending = importPptxFile(file, options); options.pictureEditSaveProfile = 'invalid';
  const imported = await pending;
  try { edit(imported.session, 'new', gif, 'image/gif', 1, 1); assert.equal((await imported.session.prepareSave({ mode: 'edit' })).status, 'ready'); } finally { imported.session.dispose(); }
});

test('picture save rejects signature relationships and missing targets in inactive branches', async () => {
  for (const type of ['urn:resource', 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature', 'http://schemas.microsoft.com/office/2006/relationships/vbaProject']) {
    const fixture = imageFixture();
    fixture.parts['deck/_rels/z.xml.rels'] = fixture.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>', `<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="${type}" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>`);
    const session = await importPptx(encodeImages(fixture), profile);
    try { edit(session, 'new', gif, 'image/gif', 1, 1); const result = await session.prepareSave({ mode: 'edit' }); assert.equal(result.status, 'blocked'); assert.equal(result.report.diagnostics[0].code, type === 'urn:resource' ? 'save-picture-candidate-unsupported' : 'save-picture-source-guard'); } finally { session.dispose(); }
  }
});

test('picture save preserves a nested group and its shared-image sibling', async () => {
  const group = children => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/><a:chOff x="0" y="0"/><a:chExt cx="127000" cy="127000"/></a:xfrm></p:grpSpPr>${children}</p:grpSp>`;
  const source = encodeImages(imageFixture(group(picture({ id: 11 }) + picture({ id: 12 }))));
  const session = await importPptx(source, profile);
  try {
    const slide = session.document.slides[0], parent = slide.elements[0], child = parent.children[0];
    session.document.registerAsset({ id: 'new', kind: 'image', source: { type: 'bytes', value: gif }, mimeType: 'image/gif', width: 1, height: 1 });
    slide.replaceElement(parent.id, { ...parent, children: [{ ...child, assetId: 'new' }, parent.children[1]] });
    const prepared = await session.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
    const reopened = await importPptx((await session.save(prepared.plan)).bytes, profile);
    try { assert.notEqual(reopened.report.pictureSources[0].mediaPart, reopened.report.pictureSources[1].mediaPart); } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test('picture save releases reservations on in-flight cancellation and session disposal', async () => {
  const source = encodeImages(), session = await importPptx(source, profile);
  try {
    edit(session, 'new', gif, 'image/gif', 1, 1);
    const controller = new AbortController(), pending = session.prepareSave({ mode: 'edit', signal: controller.signal }); controller.abort();
    await assert.rejects(pending, error => error.code === 'operation-aborted');
    const retry = await session.prepareSave({ mode: 'edit' }); assert.equal(retry.status, 'ready'); retry.plan.dispose();
    const disposed = session.prepareSave({ mode: 'edit' }); session.dispose();
    await assert.rejects(disposed, error => error.code === 'session-disposed');
  } finally { session.dispose(); }
});

test('picture replacement and starting number form an atomic candidate across three rounds in both dialects', async () => {
  const options = { ...profile, firstSlideNumberSaveProfile: 'attribute-v2' };
  for (const strict of [false, true]) {
    let source = encodeImages(imageFixture(picture({ id: 2, crop: '<a:srcRect l="12500"/>' }) + picture({ id: 3 }), strict)), previous;
    for (const [round, value] of [17, 0, undefined].entries()) {
      const s = await importPptx(source, options), original = await readArchive(source, {}, true);
      try {
        assert.equal(s.document.firstSlideNumber, previous); s.document.setFirstSlideNumber(value);
        const data = round % 2 ? png : gif, mime = round % 2 ? 'image/png' : 'image/gif'; edit(s, `new-${round}`, data, mime, round % 2 ? 8 : 1, round % 2 ? 4 : 1);
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 3);
        const next = (await s.save(p.plan)).bytes, after = await readArchive(next, {}, true);
        for (const [name, bytes] of original.parts) if (!['deck/main.data', 'deck/z.xml', 'deck/_rels/z.xml.rels', '[Content_Types].xml'].includes(name)) assert.deepEqual(after.parts.get(name), bytes, name);
        const oldMain = Buffer.from(original.parts.get('deck/main.data')).toString(), newMain = Buffer.from(after.parts.get('deck/main.data')).toString();
        if (previous === undefined) assert.equal(newMain.replace(` firstSlideNum="${value}"`, ''), oldMain);
        else if (value === undefined) assert.equal(newMain, oldMain.replace(`firstSlideNum="${previous}"`, ''));
        else assert.equal(newMain, oldMain.replace(`firstSlideNum="${previous}"`, `firstSlideNum="${value}"`));
        const reopened = await importPptx(next, options); try { const [target, sibling] = reopened.report.pictureSources; assert.equal(reopened.document.firstSlideNumber, value); assert.notEqual(target.mediaPart, sibling.mediaPart); assert.equal(sibling.mediaPart, '/media/raster.bin'); assert.deepEqual(after.parts.get(target.mediaPart.slice(1)), new Uint8Array(data)); assert.equal(reopened.document.slides[0].elements[0].crop.left, 0.125); } finally { reopened.dispose(); }
        previous = value; source = next;
      } finally { s.dispose(); }
    }
  }
});

test('combined profile also saves number-only edits without changing pictures across three rounds', async () => {
  const options = { ...profile, firstSlideNumberSaveProfile: 'attribute-v2' };
  for (const strict of [false, true]) {
    let source = encodeImages(imageFixture(picture({ id: 2, crop: '<a:srcRect l="12500"/>' }) + picture({ id: 3 }), strict));
    for (const value of [17, 0, undefined]) {
      const s = await importPptx(source, options), before = await readArchive(source, {}, true);
      try {
        s.document.setFirstSlideNumber(value);
        const slide = s.document.slides[0], pic = slide.elements[0];
        slide.replaceElement(pic.id, { ...pic, box: { ...pic.box, x: pic.box.x + 1 } });
        assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked');
        slide.replaceElement(pic.id, pic);
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report));
        assert.equal(p.plan.report.changes.length, 1);
        const output = (await s.save(p.plan)).bytes, after = await readArchive(output, {}, true);
        assert.deepEqual([...after.parts.keys()].sort(), [...before.parts.keys()].sort());
        for (const [name, bytes] of before.parts) if (name !== 'deck/main.data') assert.deepEqual(after.parts.get(name), bytes, name);
        const reopened = await importPptx(output, options);
        try { assert.equal(reopened.document.firstSlideNumber, value); assert.equal(reopened.report.pictureSources.length, 2); } finally { reopened.dispose(); }
        source = output;
      } finally { s.dispose(); }
    }
  }
});

test('combined picture and number rejects illegal edits, budgets, cancellation and stale plans', async () => {
  const bytes = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), options = { ...profile, firstSlideNumberSaveProfile: 'attribute-v2' };
  const s = await importPptx(bytes, options);
  try {
    const { slide, picture: original } = edit(s, 'new', gif, 'image/gif', 1, 1); s.document.setFirstSlideNumber(17);
    const changed = slide.elements[0]; slide.replaceElement(changed.id, { ...changed, box: { ...changed.box, x: changed.box.x + 1 } }); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); slide.replaceElement(original.id, changed);
    const controller = new AbortController(); controller.abort(); await assert.rejects(s.prepareSave({ mode: 'edit', signal: controller.signal }), { code: 'operation-aborted' });
    const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); s.document.setFirstSlideNumber(18); await assert.rejects(s.save(p.plan), { code: 'save-snapshot-stale' });
  } finally { s.dispose(); }
  for (const sessionLimits of [{ maxSaveChanges: 2 }, { maxSaveValidationBytes: 1 }, { maxSaveAnalysisWork: 2 }, { maxPreparedOutputBytes: 100 }]) {
    const limited = await importPptx(bytes, { ...options, sessionLimits }); try { edit(limited, 'new', gif, 'image/gif', 1, 1); limited.document.setFirstSlideNumber(17); await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
  }
});

test('combined picture and number stages share the output budget even when each separate plan fits', async () => {
  const source = encodeImages(), number = await importPptx(source, { projection: 'content', firstSlideNumberSaveProfile: 'attribute-v2' }); let numberSize;
  try { number.document.setFirstSlideNumber(17); const p = await number.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready'); numberSize = (await number.save(p.plan)).bytes.length; } finally { number.dispose(); }
  const picture = await importPptx(source, profile); let pictureSize;
  try { edit(picture, 'new', gif, 'image/gif', 1, 1); const p = await picture.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready'); pictureSize = (await picture.save(p.plan)).bytes.length; } finally { picture.dispose(); }
  const cap = Math.max(numberSize, pictureSize) + 256; assert.ok(cap < numberSize + pictureSize);
  const standalone = await importPptx(source, { ...profile, sessionLimits: { maxPreparedOutputBytes: cap } });
  try { edit(standalone, 'new', gif, 'image/gif', 1, 1); assert.equal((await standalone.prepareSave({ mode: 'edit' })).status, 'ready'); } finally { standalone.dispose(); }
  const both = await importPptx(source, { ...profile, firstSlideNumberSaveProfile: 'attribute-v2', sessionLimits: { maxPreparedOutputBytes: cap } });
  try { edit(both, 'new', gif, 'image/gif', 1, 1); both.document.setFirstSlideNumber(17); await assert.rejects(both.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { both.dispose(); }
});

test('failure after staging numbering publishes no partial output and corrected picture can retry', async () => {
  const source = encodeImages(), s = await importPptx(source, { ...profile, firstSlideNumberSaveProfile: 'attribute-v2' });
  try {
    edit(s, 'new', Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'), 'image/gif', 1, 1); s.document.setFirstSlideNumber(17);
    const blocked = await s.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.plan, undefined); assert.equal(blocked.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    s.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: gif }, mimeType: 'image/gif', width: 1, height: 1 });
    const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready'); const reopened = await importPptx((await s.save(ready.plan)).bytes, { ...profile, firstSlideNumberSaveProfile: 'attribute-v2' });
    try { assert.equal(reopened.document.firstSlideNumber, 17); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});
