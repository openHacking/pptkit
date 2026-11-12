import { readFile, writeFile } from "node:fs/promises";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
const [input, output] = process.argv.slice(2);
if (!input || !output) throw new Error("Usage: node scripts/prepare-native-workbook-only-chart.mjs input.pptx output.pptx");
const parts = readGeneratedZip(await readFile(input));
for (let i = 1; i <= 3; i++) {
  const name = `ppt/charts/chart${i}.xml`, xml = parts.get(name)?.toString();
  if (!xml) throw new Error(`Missing ${name}`);
  const changed = xml.replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, "");
  if (changed === xml) throw new Error(`Missing numeric cache in ${name}`);
  parts.set(name, Buffer.from(changed));
}
await writeFile(output, zip([...parts].map(([name, body]) => ({ name, body }))).bytes, { flag: "wx" });
