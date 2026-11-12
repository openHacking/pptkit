import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { nativeTableFontBytes } from "../tests/fixtures/native-table-fonts.mjs";

if (process.argv.length > 3 || (process.argv[2] && process.argv[2] !== "--filled-collection")) throw Error("Usage: node scripts/prepare-native-table-fonts.mjs [--filled-collection]");
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const fonts = JSON.parse(await readFile("docs/testing/real-font-native-evidence.json", "utf8")).fonts;
for (const font of fonts) assert.equal(hash(await readFile(font.file)), font.sha256);
const fixture = await nativeTableFontBytes(process.argv[2] ? ["collection-generic"] : undefined), options = { projection: "content", textFontProfile: "script-families-v3" };
const session = await importPptx(fixture.bytes, options), projections = [];
try {
  const document = normalizePresentation(session.document);
  for (const [index, control] of fixture.controls.entries()) {
    const table = document.slides[index].elements.find(e => e.name === "ProbeTable"), run = table.rows[0].cells[0].content[0].runs[0];
    assert.equal(run.text, control.text); assert.equal(run.style.language, control.language);
    const selected = resolveTextFontSegments(run.text, run.style, document.theme);
    assert.deepEqual([...new Set(selected.map(s => s.family))], [control.targetFamily]);
    projections.push({ id: control.id, style: run.style, selected });
  }
  const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(fixture.bytes));
  await mkdir(".tmp/native-table-fonts", { recursive: true });
  const directory = await mkdtemp(path.resolve(".tmp/native-table-fonts/run-"));
  await writeFile(path.join(directory, "table-fonts.pptx"), fixture.bytes);
  await writeFile(path.join(directory, "expected.json"), JSON.stringify({ recordedAt: new Date().toISOString(), sourceSha256: hash(fixture.bytes), controls: fixture.controls, maps: fixture.maps, fonts, options, projections, properties: session.report.tableStyleInheritance.properties }, null, 2) + "\n");
  await writeFile(".tmp/native-table-fonts/latest.txt", directory);
  console.log(JSON.stringify({ directory, file: path.join(directory, "table-fonts.pptx"), controls: fixture.controls.length, sourceNoOpExact: true }));
} finally { session.dispose(); }
