import test from 'node:test';
import assert from 'node:assert/strict';
import { planPropertyCreation } from '../dist/preservation/plan-property-creation.js';
import { appendPropertyDeclarations } from '../dist/preservation/append-property-declarations.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';

const plan = (bytes, dialect = 'transitional') => planPropertyCreation(bytes, 'extended', dialect, {}, 1024 * 1024, 128);

test('extended creation plans collision-free declarations in both dialects', async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict), dialect = strict ? 'strict' : 'transitional';
    f.parts['pptkit-extended-properties-1.xml'] = '<unused/>';
    f.rootRelations.push({ Id: 'pptkitExtendedProperties2', Type: 'urn:example', Target: 'deck/main.data' });
    const p = await plan(encodeMetadata(f), dialect);
    assert.equal(p.part, '/pptkit-extended-properties-3.xml');
    assert.equal(p.relationshipId, 'pptkitExtendedProperties3');
    assert.match(Buffer.from(p.relationships).toString(), strict ? /relationships\/extendedProperties/ : /relationships\/extended-properties/);
    assert.match(Buffer.from(p.contentTypes).toString(), /application\/vnd.openxmlformats-officedocument.extended-properties\+xml/);
  }
});

test('extended creation refuses either dialect of existing relationships and predeclared MIME', async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict); f.add('extended', '');
    await assert.rejects(plan(encodeMetadata(f), strict ? 'transitional' : 'strict'), { code: 'xml-edit-conflict' });
    const orphan = metadataFixture(strict); orphan.parts['inventory/unused.data'] = '<unused/>'; orphan.metadataTypes['inventory/unused.data'] = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
    await assert.rejects(plan(encodeMetadata(orphan), strict ? 'strict' : 'transitional'), { code: 'xml-edit-conflict' });
  }
});

test('extended declarations merge with independently planned core and custom declarations', async () => {
  for (const strict of [false, true]) {
    const dialect = strict ? 'strict' : 'transitional', bytes = encodeMetadata(metadataFixture(strict));
    const core = await planPropertyCreation(bytes, 'core', dialect, {}, 1024 * 1024, 128);
    const extended = await plan(bytes, dialect);
    const custom = await planPropertyCreation(bytes, 'custom', dialect, {}, 1024 * 1024, 128);
    const merged = await appendPropertyDeclarations(core.relationships, core.contentTypes, { kind: 'extended', part: extended.part, relationshipId: extended.relationshipId, dialect }, { maxOutputBytes: 1024 * 1024 });
    const all = await appendPropertyDeclarations(merged.relationships, merged.contentTypes, { kind: 'custom', part: custom.part, relationshipId: custom.relationshipId, dialect }, { maxOutputBytes: 1024 * 1024 });
    for (const id of [core.relationshipId, extended.relationshipId, custom.relationshipId]) assert.match(Buffer.from(all.relationships).toString(), new RegExp(`Id="${id}"`));
    for (const part of [core.part, extended.part, custom.part]) assert.ok(Buffer.from(all.contentTypes).toString().includes(`PartName="${part}"`));
    await assert.rejects(appendPropertyDeclarations(all.relationships, all.contentTypes, { kind: 'extended', part: extended.part, relationshipId: extended.relationshipId, dialect }, { maxOutputBytes: 1024 * 1024 }), { code: 'xml-edit-conflict' });
  }
});
