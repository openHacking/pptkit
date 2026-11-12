import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import test from 'node:test';

function unregistered(scripts, files) {
  const reachable = new Set(), pending = ['test'];
  while (pending.length) {
    const name = pending.pop();
    if (reachable.has(name)) continue;
    assert.equal(typeof scripts[name], 'string', `missing script ${name}`);
    reachable.add(name);
    for (const match of scripts[name].matchAll(/\bpnpm\s+(?:run\s+)?(test:[\w-]+)/g)) pending.push(match[1]);
  }
  const commands = [...reachable].map(name => scripts[name]).join(' ');
  return files.filter(name => name.endsWith('-parity.test.mjs') && !commands.split(/\s+/).includes(`scripts/${name}`));
}

test('every root cross-package parity test is reachable from the standard test command', async () => {
  const { scripts } = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
  assert.deepEqual(unregistered(scripts, await readdir(new URL('.', import.meta.url))), [], 'register each cross-package parity test in a suite reachable from pnpm test');
});

test('registration audit rejects unreachable and partially matching script names', () => {
  const scripts = { test: 'pnpm run test:included', 'test:included': 'node --test scripts/one-parity.test.mjs', 'test:unused': 'node --test scripts/two-parity.test.mjs' };
  assert.deepEqual(unregistered(scripts, ['one-parity.test.mjs', 'two-parity.test.mjs', 'ne-parity.test.mjs', 'optional.test.mjs']), ['two-parity.test.mjs', 'ne-parity.test.mjs']);
});
