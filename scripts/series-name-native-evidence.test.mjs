import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

test('retained PowerPoint series-name reopen inputs match their hashes and independent source changes', async () => {
  const hashes = ['028a10b4a11c1a644568f026627774ecadebf473c8463b4a02c0071571a4c359', 'b24156b7f05f4db9507c2b6f879cd04fcbf1deb183cf1b93faa67824668554df', '6838385c792b73ec55d5c9c9a01a1c3363a49f9c9af85608e925d9b215d72484'];
  const names = ['Native revised', 'Native second', 'Native third'];
  let before = 'tests/fixtures/chart-label-source.pptx';
  for (const [index, expected] of names.entries()) {
    const after = `tests/fixtures/chart-series-name-native-round-${index + 1}.pptx`, bytes = await readFile(after);
    assert.equal(createHash('sha256').update(bytes).digest('hex'), hashes[index]);
    const result = spawnSync('python3', ['scripts/research/verify-chart-series-name-save.py', before, after, JSON.stringify(expected)], { encoding: 'utf8' });
    assert.equal(result.status, 0, result.stderr); before = after;
  }
});

test('retained PowerPoint resave contains agreeing chart and workbook series names', async () => {
  const path = 'tests/fixtures/chart-series-name-powerpoint-resave.pptx', bytes = await readFile(path);
  assert.equal(createHash('sha256').update(bytes).digest('hex'), '8fe55fb2dc2d4b6a25b852dc46b6ecbf34562cd6036a0a8c91e254de7041414b');
  const result = spawnSync('python3', ['scripts/research/verify-powerpoint-series-resave.py', path], { encoding: 'utf8' }); assert.equal(result.status, 0, result.stderr);
});
