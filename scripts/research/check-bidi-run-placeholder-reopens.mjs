import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';

const files = process.argv.slice(2);
assert.ok(files.length > 0, 'Pass the generated and/or PowerPoint-saved PPTX paths');
for (const path of files) {
  const bytes = await readFile(path);
  const session = await importPptx(bytes, { projection: 'content' });
  try {
    assert.equal(session.document.slides.length, 2);
    for (const [i, slide] of session.document.slides.entries()) {
      const runs = slide.elements.map(e => e.content?.flatMap(p => p.runs) ?? []);
      const text = r => r.map(item => item.text).join('');
      assert.equal(runs.length, 2);
      assert.equal(text(runs[0]), i === 0 ? 'PAGE ONE' : 'PAGE TWO');
      const expected = [['ABC ', false], ['مرحبا - 123', i === 0], [' XYZ', false], [' שלום - 456', i !== 0]];
      const expand = pairs => pairs.flatMap(([t, rtl]) => Array.from(t, c => [c, rtl]));
      assert.deepEqual(expand(runs[1].map(r => [r.text, r.style.rightToLeft])), expand(expected));
    }
    const defaults = session.report.textStyleInheritance.properties.filter(p => p.property === 'run.rightToLeft' && p.level === 'layout-placeholder');
    assert.deepEqual(new Set(defaults.map(p => p.value)), new Set([true, false]));
    console.log(path + ': both pages, effective run directions, and layout provenance verified');
  } finally { session.dispose(); }
}
