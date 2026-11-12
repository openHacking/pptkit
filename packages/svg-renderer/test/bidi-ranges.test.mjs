import test from 'node:test';
import assert from 'node:assert/strict';
import { needsBidiParagraphLayout } from '../dist/rendering/bidi-ranges.js';

test('Unicode bidi routing detects strong RTL, Arabic numbers and directional controls', () => {
  for (const text of ['שלום', 'مرحبا', 'ܫܠܡ', 'ދިވެހި', 'ߒߞߏ', 'ࠔ', 'ࡀ', '𞤀', '𐴁', '١٢٣', '\u061c', '\u200e', '\u200f', '\u202a', '\u202e', '\u2066', '\u2069']) {
    assert.equal(needsBidiParagraphLayout(text), true, JSON.stringify(text));
    assert.equal(needsBidiParagraphLayout(`Latin ${text} 123`), true, JSON.stringify(text));
  }
  for (const text of ['', 'English 123', '中文 Ελληνικά', 'a\u0301', '👩🏽‍💻', '\ud800', '\udfff']) {
    assert.equal(needsBidiParagraphLayout(text), false, JSON.stringify(text));
  }
});
