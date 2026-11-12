import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { tmpdir } from 'node:os';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { zip } from '../packages/pptx-parser/test/zip-fixture.mjs';

const replacement = 'tests/fixtures/transparent-replacement.png';
const python = (script, ...args) => execFileSync('python3', [`scripts/research/${script}.py`, ...args], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
const makeZip = parts => zip([...parts].map(([name, body]) => ({ name, body }))).bytes;

test('picture and number oracles reject altered media, references, crop, geometry and slide order', async t => {
  const parent = await mkdtemp(path.join(tmpdir(), 'pptkit-picture-number-')); t.after(() => rm(parent, { recursive: true, force: true }));
  const dir = path.join(parent, 'sample'); execFileSync(process.execPath, ['scripts/prepare-picture-number-native.mjs', dir]);
  const source = path.join(dir, 'source.pptx'), candidate = path.join(dir, 'candidate.pptx');
  const verify = (s, c) => python('verify-picture-number', s, c, replacement);
  const native = (c, pos) => python('verify-picture-number-native', candidate, c, String(pos));
  assert.equal(JSON.parse(verify(source, candidate)).starting_number, 17);
  assert.equal(JSON.parse(native(candidate, 0)).picture_page, 1);
  assert.throws(() => native(candidate, 1));
  const original = readGeneratedZip(await readFile(candidate));
  const changes = [
    ['ppt/presentation.xml', s => s.replace('firstSlideNum="17"', 'firstSlideNum="18"')],
    ['ppt/slides/slide1.xml', s => s.replace('r:embed="pptkitImage1"', 'r:embed="rId4"')],
    ['ppt/slides/slide1.xml', s => s.replace('r:embed="rId4"', 'r:embed="pptkitImage1"')],
    ['ppt/slides/slide1.xml', s => s.replace('l="12500"', 'l="10000"')],
    ['ppt/slides/slide1.xml', s => s.replace('x="381000"', 'x="381001"')],
    ['ppt/slides/slide2.xml', s => s.replace('Control page', 'wrong page')],
    ['ppt/slides/_rels/slide1.xml.rels', s => s.replace('../../ppt/media/pptkit-image-1.png', '../media/shared.png')],
  ];
  for (const [index, [name, edit]] of changes.entries()) {
    const modified = new Map(original), before = modified.get(name).toString(), after = edit(before); assert.notEqual(after, before, name); modified.set(name, Buffer.from(after));
    const target = path.join(parent, `bad-${index}.pptx`); await writeFile(target, makeZip(modified));
    assert.throws(() => verify(source, target)); assert.throws(() => native(target, 0));
  }
  for (const name of ['ppt/media/shared.png', 'ppt/media/pptkit-image-1.png']) {
    const modified = new Map(original); modified.set(name, Buffer.from('incorrect media'));
    const target = path.join(parent, `bad-${path.basename(name)}`); await writeFile(target, makeZip(modified));
    assert.throws(() => verify(source, target)); assert.throws(() => native(target, 0));
  }
  const old = readGeneratedZip(await readFile(source));
  const oldRels = old.get('ppt/slides/_rels/slide1.xml.rels').toString();
  const badRels = oldRels.replace('Target="../media/shared.png"', 'Target="../media/other.png"'); assert.notEqual(oldRels, badRels);
  old.set('ppt/slides/_rels/slide1.xml.rels', Buffer.from(badRels));
  const forgedSource = path.join(parent, 'bad-source.pptx'); await writeFile(forgedSource, makeZip(old));
  assert.throws(() => verify(forgedSource, candidate));
  const forgedCandidateParts = new Map(original), candidateRels = forgedCandidateParts.get('ppt/slides/_rels/slide1.xml.rels').toString();
  const forgedRels = candidateRels.replace('Target="../media/shared.png"', 'Target="../media/other.png"'); assert.notEqual(forgedRels, candidateRels);
  forgedCandidateParts.set('ppt/slides/_rels/slide1.xml.rels', Buffer.from(forgedRels));
  const forgedCandidate = path.join(parent, 'bad-source-and-candidate.pptx'); await writeFile(forgedCandidate, makeZip(forgedCandidateParts));
  assert.throws(() => verify(forgedSource, forgedCandidate));
});
