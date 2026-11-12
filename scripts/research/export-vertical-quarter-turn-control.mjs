import { readFile, writeFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

const folder = process.argv[2] ?? '.tmp/vertical-writing';
const profiles = {
  quarter: { pages: [0, 1, 2], output: 'quarter-turn-fresh.pptx' },
  eaVert: { pages: [3], output: 'east-asian-fresh.pptx' },
  remaining: { pages: [4, 5, 6], output: 'stacked-mongolian-fresh.pptx' },
};
const profile = process.argv[3] ?? 'quarter';
if (!Object.hasOwn(profiles, profile)) throw new Error('Unknown writing mode profile: ' + profile);
const { pages, output } = profiles[profile];
const session = await importPptx(await readFile(folder + '/vertical-matrix.pptx'), { projection: 'content', textFontProfile: 'script-families-v1' });
try {
  const deck = session.document;
  for (const [i, page] of deck.slides.entries()) if (!pages.includes(i)) deck.removeSlide(page.id);
  await writeFile(folder + '/' + output, (await generatePptx(deck)).bytes, { flag: 'wx' });
} finally { session.dispose(); }
