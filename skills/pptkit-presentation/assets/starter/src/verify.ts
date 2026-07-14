import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { normalizePresentation, type NormalizedElement, type NormalizedPresentation, type PresentationDocument } from "@pptkit/core";
import { unzipSync } from "fflate";

import type { StructuralIssue } from "./contracts.js";

function visitElements(elements: NormalizedElement[], visit: (element: NormalizedElement) => void) {
  for (const element of elements) {
    visit(element);
    if (element.type === "group") visitElements(element.children, visit);
  }
}

function overlapArea(a: NormalizedElement["box"], b: NormalizedElement["box"]) {
  const width = Math.max(0, Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x));
  const height = Math.max(0, Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y));
  return width * height;
}

export function inspectStructure(input: PresentationDocument | NormalizedPresentation): StructuralIssue[] {
  const document = "irVersion" in input ? input : normalizePresentation(input);
  const issues: StructuralIssue[] = [];

  for (const slide of document.slides) {
    const visible: NormalizedElement[] = [];
    visitElements(slide.elements, (element) => {
      if (element.hidden) return;
      const { x, y, width, height } = element.box;
      if (x < -0.01 || y < -0.01 || x + width > document.size.width + 0.01 || y + height > document.size.height + 0.01) {
        issues.push({ severity: "error", code: "out-of-bounds", message: `Element ${element.name} exceeds the ${document.size.width}×${document.size.height} slide.`, slideId: slide.id, elementId: element.id });
      }
      if (element.type !== "connector" && !element.accessibility.decorative) visible.push(element);
    });

    for (let left = 0; left < visible.length; left += 1) {
      for (let right = left + 1; right < visible.length; right += 1) {
        const a = visible[left]!;
        const b = visible[right]!;
        const overlap = overlapArea(a.box, b.box);
        const smaller = Math.min(a.box.width * a.box.height, b.box.width * b.box.height);
        if (smaller > 0 && overlap / smaller > 0.35) {
          issues.push({ severity: "error", code: "high-overlap", message: `${a.name} substantially overlaps ${b.name}; adjust the layout and inspect the rendered slide.`, slideId: slide.id, elementId: b.id });
        }
      }
    }
  }
  return issues;
}

export async function inspectPptxPackage(file: string) {
  const bytes = new Uint8Array(await readFile(file));
  const files = unzipSync(bytes);
  const names = Object.keys(files);
  const issues: string[] = [];
  for (const required of ["[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"]) {
    if (!names.includes(required)) issues.push(`Missing required package part: ${required}`);
  }
  const slideParts = names.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  for (const name of names.filter((entry) => entry.endsWith(".xml") || entry.endsWith(".rels"))) {
    const text = new TextDecoder().decode(files[name]);
    if (!text.includes("<") || !text.includes(">")) issues.push(`Malformed XML-shaped part: ${name}`);
  }
  return { valid: issues.length === 0, parts: names.length, slideParts: slideParts.length, issues };
}

async function main() {
  const file = path.resolve(process.argv[2] ?? "output/deck.pptx");
  const result = await inspectPptxPackage(file);
  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
  if (!result.valid) process.exitCode = 1;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  main().catch(async (error) => {
    await writeFile("output/verify-error.txt", `${error instanceof Error ? error.stack : String(error)}\n`).catch(() => undefined);
    throw error;
  });
}
