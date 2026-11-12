import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, open, rename, rm, stat, symlink, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { readInput } from '../dist/commands/read-file.js';

async function fixture(t) {
  const dir = await mkdtemp(path.join(tmpdir(), 'pptkit-capture-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  await writeFile(path.join(dir, 'input'), 'original');
  return dir;
}

test('capture detects equal-length path replacement before and after reading', async t => {
  for (const replacementAt of [1, 2]) {
    const dir = await fixture(t); let calls = 0;
    const io = { openFile: open, statTarget: async target => {
      if (++calls === replacementAt) {
        await rename(target, `${target}.old`);
        await writeFile(target, 'replaced');
      }
      return stat(target, { bigint: true });
    } };
    await assert.rejects(readInput('input', dir, 100, undefined, io), error => error.code === 'file-changed');
  }
});

test('capture detects disappearance and symlink retargeting during reading', async t => {
  for (const operation of ['disappear', 'retarget']) {
    const dir = await fixture(t);
    await writeFile(path.join(dir, 'other'), 'replaced');
    await symlink('input', path.join(dir, 'link'));
    let calls = 0;
    const io = { openFile: open, statTarget: async target => {
      if (++calls === 2) {
        await unlink(target);
        if (operation === 'retarget') await symlink('other', target);
      }
      return stat(target, { bigint: true });
    } };
    await assert.rejects(readInput('link', dir, 100, undefined, io), error => error.code === 'file-changed');
  }
});

test('stable symlink reports its argument path and captured bytes', async t => {
  const dir = await fixture(t);
  await symlink('input', path.join(dir, 'link'));
  const result = await readInput('link', dir, 100);
  assert.equal(result.input.path, 'link');
  assert.equal(Buffer.from(result.bytes).toString(), 'original');
  assert.equal(result.input.byteLength, 8);
});
