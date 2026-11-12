import { mkdir, writeFile } from 'node:fs/promises';
import { importSimpleSvg } from '../packages/svg-parser/dist/index.js';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const source = '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" fill="none" stroke-width="24"><path d="M100 300 L220 120 L340 300" stroke="red" stroke-linecap="round" stroke-linejoin="round"/><path d="M500 130 H760 V310 H500 Z" stroke="blue" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const result = importSimpleSvg(source);
if (result.status !== 'editable') throw new Error(JSON.stringify(result.issues));
const deck = createPresentation({size: {width: result.width, height: result.height}});
deck.addSlide({elements: [...result.elements]});
const directory = process.argv[2] ?? '.tmp/native-svg-path';
await mkdir(directory, {recursive: true});
await writeFile(`${directory}/source.svg`, source, {flag: 'wx'});
await writeFile(`${directory}/source.pptx`, (await generatePptx(deck)).bytes, {flag: 'wx'});
