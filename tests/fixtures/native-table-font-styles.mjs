import { nativeTableFontBytes } from "./native-table-fonts.mjs";
import { readGeneratedZip } from "../support/read-generated-zip.mjs";
import { zip } from "../../packages/pptx-parser/test/zip-fixture.mjs";

/** Source-only control reproducing observed run/style declarations, without native execution or font loading. */
export async function nativeTableStyleSourceBytes(stage, filled = false) {
  const flags = { bold: ' b="1"', italic: ' b="0" i="1"', regular: ' b="0" i="0"' }[stage];
  if (flags === undefined) throw Error("Unknown native table style source stage");
  const fixture = await nativeTableFontBytes(filled ? ["collection-generic"] : undefined), parts = readGeneratedZip(fixture.bytes);
  const targets = filled ? [3, 5, 6, 7, 8] : [3, 5, 6, 7, 12, 14, 15, 21, 25, 26];
  for (const index of targets) {
    const name = `ppt/slides/slide${index}.xml`, control = fixture.controls[index - 1]; let count = 0;
    const source = parts.get(name).toString().replace(/<a:tbl>[\s\S]*?<\/a:tbl>/, table => table.replace(/<a:r><a:rPr\b[^>]*\/><a:t>[^<]*<\/a:t><\/a:r>/, run => {
      count++; const alternate = ["ar", "he"].includes(control.sampleId) ? "" : ` altLang="${control.language}"`;
      return run.replace('/><a:t>', `${flags} dirty="0"/><a:t>`) + `<a:r><a:rPr lang="el-GR"${alternate} sz="4800"${flags} dirty="0"/><a:t>Ω</a:t></a:r>`;
    }));
    if (count !== 1) throw Error("Expected exactly one patched cell run");
    parts.set(name, Buffer.from(source));
  }
  return { ...fixture, bytes: zip([...parts].map(([name, body]) => ({ name, body }))).bytes, targets };
}
