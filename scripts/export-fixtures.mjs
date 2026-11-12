import { mkdir, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { createFixture, fixtureCatalog } from "../tests/fixtures/presentations.mjs";

const args = process.argv.slice(2);
if (args.length > 1) throw new Error("Usage: node scripts/export-fixtures.mjs [output-directory]");
const directory = path.resolve(args[0] ?? ".tmp/fixtures");
await mkdir(directory, { recursive: true });
const manifest = { schemaVersion: 1, fixtures: [] };
for (const fixture of fixtureCatalog) {
  const presentation = createFixture(fixture.id, createPresentation);
  const exported = await generatePptx(presentation);
  const preview = await renderPresentationToSvg(presentation);
  const filename = `${fixture.id}.pptx`;
  await writeFile(path.join(directory, filename), exported.bytes);
  await writeFile(path.join(directory, `${fixture.id}.ir.json`), `${JSON.stringify(normalizePresentation(presentation), null, 2)}\n`);
  for (const [index, slide] of preview.slides.entries()) await writeFile(path.join(directory, `${fixture.id}-${index + 1}.svg`), slide.svg);
  manifest.fixtures.push({
    id: fixture.id, capabilities: fixture.capabilities, filename,
    sha256: createHash("sha256").update(exported.bytes).digest("hex"), bytes: exported.byteLength,
    slideCount: exported.slideCount, exportWarnings: exported.warnings, previewWarnings: preview.warnings,
  });
}
await writeFile(path.join(directory, "manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`Exported ${manifest.fixtures.length} fixtures to ${directory}`);
