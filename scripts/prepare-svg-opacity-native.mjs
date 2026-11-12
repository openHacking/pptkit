import { mkdir, writeFile } from 'node:fs/promises';
import { importSimpleSvg } from '../packages/svg-parser/dist/index.js';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const source = '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540"><rect x="100" y="100" width="150" height="120" fill="#f008" fill-opacity=".5"/><ellipse cx="430" cy="160" rx="100" ry="70" fill="none" stroke="#00ff0080" stroke-opacity=".8" stroke-width="24"/><rect x="650" y="100" width="150" height="120" fill="rgba(0,0,255,.5)" fill-opacity="50%"/></svg>';
const result = importSimpleSvg(source);
if (result.status !== 'editable') throw new Error(JSON.stringify(result.issues));
const deck = createPresentation({size: {width: result.width, height: result.height}});
deck.addSlide({elements: [...result.elements]});
await mkdir('.tmp/native-svg-opacity', {recursive: true});
await writeFile('.tmp/native-svg-opacity/source.svg', source, {flag: 'wx'});
await writeFile('.tmp/native-svg-opacity/source.pptx', (await generatePptx(deck)).bytes, {flag: 'wx'});
