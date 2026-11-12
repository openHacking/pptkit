import assert from 'node:assert/strict';
import test from 'node:test';
import { splitAtGraphemeBoundary } from '../dist/preservation/split-text-boundary.js';

test('split preflight accepts only interior grapheme boundaries without changing text', async () => {
  const clusters = ['A', 'e\u0301', '👨‍👩‍👧‍👦', '🇨🇳', 'Z'], text = clusters.join('');
  let offset = 0;
  for (let i = 0; i < clusters.length - 1; i++) {
    offset += clusters[i].length;
    const split = await splitAtGraphemeBoundary(text, offset);
    assert.equal(split.left, clusters.slice(0, i + 1).join(''));
    assert.equal(split.right, clusters.slice(i + 1).join(''));
    assert.ok(Object.isFrozen(split));
  }
  for (const invalid of [0, text.length, 2, 4, 5, 6]) {
    await assert.rejects(splitAtGraphemeBoundary(text, invalid), { code: 'split-invalid' });
  }
  for (const invalid of ['a\ud800b', 'a\udc00b']) await assert.rejects(splitAtGraphemeBoundary(invalid, 1), { code: 'split-invalid' });
});

test('split preflight bounds work, supports cancellation, and fails closed without segmentation', async () => {
  await assert.rejects(splitAtGraphemeBoundary('abc', 1, { maxCharacters: 2 }), { code: 'split-limit-exceeded' });
  assert.deepEqual(await splitAtGraphemeBoundary('abc', 1, { maxCharacters: 3, maxGraphemes: 2 }), { left: 'a', right: 'bc' });
  await assert.rejects(splitAtGraphemeBoundary('abc', 1, { maxGraphemes: 1 }), { code: 'split-limit-exceeded' });
  await assert.rejects(splitAtGraphemeBoundary('abc', 1, { maxGraphemes: 0 }), RangeError);
  const stopped = new AbortController(); stopped.abort();
  await assert.rejects(splitAtGraphemeBoundary('abc', 1, { signal: stopped.signal }), { code: 'operation-aborted' });
  const during = new AbortController(), pending = splitAtGraphemeBoundary('a'.repeat(20000), 19999, { signal: during.signal });
  setTimeout(() => during.abort(), 0);
  await assert.rejects(pending, { code: 'operation-aborted' });
  const descriptor = Object.getOwnPropertyDescriptor(Intl, 'Segmenter');
  try {
    Object.defineProperty(Intl, 'Segmenter', { configurable: true, value: undefined });
    await assert.rejects(splitAtGraphemeBoundary('abc', 1), { code: 'split-unavailable' });
  } finally {
    if (descriptor) Object.defineProperty(Intl, 'Segmenter', descriptor);
    else delete Intl.Segmenter;
  }
});
