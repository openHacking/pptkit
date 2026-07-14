import { copyFile, mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import type { ExtractedSource } from "./contracts.js";

const inputs = process.argv.slice(2);
if (inputs.length === 0) throw new Error("Usage: npm run extract -- <source paths...>");

await mkdir("content", { recursive: true });
await mkdir("assets", { recursive: true });

function sourceId(file: string, index: number) {
  const stem = path.basename(file, path.extname(file)).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "source";
  return `src-${String(index + 1).padStart(2, "0")}-${stem}`;
}

async function extractPdf(file: string) {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const data = new Uint8Array(await readFile(file));
  const pdf = await pdfjs.getDocument({ data, useWorkerFetch: false, isEvalSupported: false, useSystemFonts: true }).promise;
  const pages: string[] = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const text = await page.getTextContent();
    pages.push(text.items.map((item) => "str" in item ? item.str : "").join(" "));
  }
  return pages.join("\n\n");
}

async function extractOne(input: string, index: number): Promise<ExtractedSource> {
  const file = path.resolve(input);
  const id = sourceId(file, index);
  const extension = path.extname(file).toLowerCase();
  const base = { id, path: file, warnings: [] as string[] };
  if ([".md", ".markdown", ".txt"].includes(extension)) return { ...base, type: "text", content: await readFile(file, "utf8") };
  if (extension === ".pdf") return { ...base, type: "document", content: await extractPdf(file) };
  if (extension === ".docx") {
    const mammoth = await import("mammoth");
    const result = await mammoth.extractRawText({ path: file });
    return { ...base, type: "document", content: result.value, warnings: result.messages.map((message) => message.message) };
  }
  if ([".csv", ".xlsx", ".xls"].includes(extension)) {
    const XLSX = await import("xlsx");
    const workbook = XLSX.read(await readFile(file), { type: "buffer" });
    const sheets = workbook.SheetNames.map((name) => ({ name, rows: XLSX.utils.sheet_to_json(workbook.Sheets[name]!, { header: 1, raw: false }) as unknown[][] }));
    return { ...base, type: "table", sheets };
  }
  if ([".png", ".jpg", ".jpeg", ".gif", ".svg"].includes(extension)) {
    const { imageSize } = await import("image-size");
    const destination = path.resolve("assets", `${id}${extension}`);
    await copyFile(file, destination);
    let dimensions: { width?: number; height?: number } = {};
    try {
      const measured = imageSize(new Uint8Array(await readFile(file)));
      dimensions = { width: measured.width, height: measured.height };
    } catch (error) {
      base.warnings.push(`Could not determine image dimensions: ${error instanceof Error ? error.message : String(error)}`);
    }
    return { ...base, type: "image", image: { copiedPath: path.relative(process.cwd(), destination), ...dimensions } };
  }
  if (extension === ".pptx") return { ...base, type: "document", warnings: ["PPTX parsing and template reuse are not supported in v1; treat this file as a manual visual reference."] };
  return { ...base, type: "document", warnings: [`Unsupported source type: ${extension || "no extension"}`] };
}

const sources: ExtractedSource[] = [];
let failures = 0;
for (let index = 0; index < inputs.length; index += 1) {
  try {
    sources.push(await extractOne(inputs[index]!, index));
  } catch (error) {
    failures += 1;
    sources.push({ id: sourceId(inputs[index]!, index), path: path.resolve(inputs[index]!), type: "document", warnings: [`Extraction failed: ${error instanceof Error ? error.message : String(error)}`] });
  }
}

await writeFile("content/sources.json", `${JSON.stringify({ generatedAt: new Date().toISOString(), sources }, null, 2)}\n`);
process.stdout.write(`${JSON.stringify({ sources: sources.length, failures, output: path.resolve("content/sources.json") }, null, 2)}\n`);
if (failures > 0) process.exitCode = 1;
