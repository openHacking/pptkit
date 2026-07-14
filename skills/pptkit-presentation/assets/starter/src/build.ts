import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { writePptx } from "@pptkit/pptx-exporter/node";

import type { BuildReport, StructuralIssue } from "./contracts.js";
import { deckSpec } from "./deck-spec.js";
import { authorPresentation, validateDeckSpec } from "./runtime/author.js";
import { inspectPptxPackage, inspectStructure } from "./verify.js";

const outputDir = path.resolve("output");
const output = path.join(outputDir, "deck.pptx");
const reportPath = path.join(outputDir, "build-report.json");

await mkdir(outputDir, { recursive: true });

const specIssues = validateDeckSpec(deckSpec);
const document = authorPresentation(deckSpec);
const diagnostics = validatePresentation(document);
const diagnosticIssues: StructuralIssue[] = diagnostics
  .filter((item) => item.severity === "error")
  .map((item) => ({ severity: "error", code: item.code, message: item.message, slideId: item.slideId, elementId: item.elementId }));
const structuralIssues = diagnostics.some((item) => item.severity === "error")
  ? [...specIssues, ...diagnosticIssues]
  : [...specIssues, ...inspectStructure(normalizePresentation(document))];

let report: BuildReport = {
  output,
  slideCount: document.slides.length,
  byteLength: 0,
  diagnostics: diagnostics.map(({ severity, code, message, path: diagnosticPath }) => ({ severity, code, message, path: diagnosticPath })),
  exportWarnings: [],
  structuralIssues,
  packageChecks: { valid: false, parts: 0, slideParts: 0, issues: ["Build did not run."] },
  renderStatus: "not-run",
  generatedAt: new Date().toISOString(),
};

if (structuralIssues.some((issue) => issue.severity === "error")) {
  await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);
  throw new Error(`Deck failed pre-export checks. See ${reportPath}`);
}

const result = await writePptx(document, { output });
const packageChecks = await inspectPptxPackage(output);
const exportIssues: StructuralIssue[] = result.warnings.map((warning) => ({ severity: "error", code: `export-${warning.code}`, message: warning.message, slideId: warning.slideId }));
if (!packageChecks.valid) exportIssues.push(...packageChecks.issues.map((message) => ({ severity: "error" as const, code: "invalid-package", message })));
if (packageChecks.slideParts !== document.slides.length) exportIssues.push({ severity: "error", code: "slide-part-count", message: `Expected ${document.slides.length} slide parts, found ${packageChecks.slideParts}.` });

report = {
  ...report,
  byteLength: result.byteLength,
  exportWarnings: result.warnings,
  structuralIssues: [...structuralIssues, ...exportIssues],
  packageChecks,
};
await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);

process.stdout.write(`${JSON.stringify({ output, report: reportPath, slideCount: result.slideCount, warnings: result.warnings.length }, null, 2)}\n`);
if (report.structuralIssues.some((issue) => issue.severity === "error")) process.exitCode = 1;
