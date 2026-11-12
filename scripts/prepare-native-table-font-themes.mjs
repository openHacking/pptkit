import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, writeFile, copyFile } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { nativeFontFamilies as F } from "./research/font-native-cases.mjs";

if (process.argv.length !== 3) throw Error("Usage: node scripts/prepare-native-table-font-themes.mjs <verified-native-style-run>");
const sourceDirectory = path.resolve(process.argv[2]), hash = bytes => createHash("sha256").update(bytes).digest("hex"), root = path.resolve(".tmp/native-table-font-themes");
await mkdir(root, { recursive: true }); const directory = await mkdtemp(path.join(root, "run-"));
for (const sourceGroup of ["reference", "filled"]) for (const local of [false, true]) {
  const group = sourceGroup + (local ? "-local" : ""), folder = path.join(directory, group); await mkdir(folder);
  const source = path.join(sourceDirectory, sourceGroup), input = await readFile(path.join(source, "regular-round3.pptx")), expected = JSON.parse(await readFile(path.join(source, "expected.json"), "utf8"));
  const baseline = readGeneratedZip(input), parts = new Map(baseline);
  if (local) for (const index of expected.editTargets) {
    const part = `ppt/slides/slide${index}.xml`; let count = 0;
    const text = parts.get(part).toString().replace(/<a:tbl>[\s\S]*?<\/a:tbl>/, table => table.replace(/<a:rPr\b([^>]*)\/>/g, (_, attributes) => {
      count++; return `<a:rPr${attributes}><a:latin typeface="${F.latin}"/><a:ea typeface="${F.generic}"/><a:cs typeface="${F.generic}"/></a:rPr>`;
    }));
    assert.equal(count, 2); parts.set(part, Buffer.from(text));
  }
  const bytes = local ? zip([...parts].map(([name, body]) => ({ name, body }))).bytes : input;
  await writeFile(path.join(folder, "baseline.pptx"), bytes);
  if (!local) await copyFile(path.join(source, "regular-round3.pdf"), path.join(folder, "baseline.pdf"));
  const pdf = await readFile(path.join(source, "regular-round3.pdf"));
  Object.assign(expected, { group, localOverrideTargets: local ? expected.editTargets : [], localFonts: { latin: F.latin, ea: F.generic, cs: F.generic }, sourceNative: { directory: source, file: path.join(source, "regular-round3.pptx"), sha256: hash(input), pdfSha256: hash(pdf) }, themeBaseline: { pptx: hash(bytes), pdf: local ? null : hash(pdf) }, sourceAssembly: local ? "Named local-slot source patch on verified native regular input; requires separate native baseline PDF" : "Exact verified native regular input copy" });
  await writeFile(path.join(folder, "expected.json"), JSON.stringify(expected, null, 2) + "\n");
}
console.log(directory);
