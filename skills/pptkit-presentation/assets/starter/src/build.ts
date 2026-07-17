import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter/node";
import { authorDeck, inspectPptxPackage, inspectStructure, validateDeckSpec } from "@pptkit/presentation-workflow";

import type { BuildReport, StructuralIssue } from "./contracts.js";
import { deckSpec } from "./deck-spec.js";
import { mimeTypeForAsset, resolveNodeAsset } from "./runtime/node-assets.js";

const outputDir = path.resolve("output");
const output = path.join(outputDir, "deck.pptx");
const reportPath = path.join(outputDir, "build-report.json");

await mkdir(outputDir, { recursive: true });

const availableAssetIds = new Set(deckSpec.slides.flatMap((slide) => slide.image ? [slide.image.assetId] : []));
const specIssues = validateDeckSpec(deckSpec, availableAssetIds);
const { presentation: document, layoutDecisions } = authorDeck(deckSpec, (assetId) => resolveNodeAsset(assetId, mimeTypeForAsset(assetId)));
const diagnostics = validatePresentation(document);
const diagnosticIssues: StructuralIssue[] = diagnostics
  .filter((item) => item.severity === "error")
  .map((item) => ({ severity: "error", code: item.code, message: item.message, slideId: item.slideId, elementId: item.elementId }));
const structuralIssues = diagnostics.some((item) => item.severity === "error")
  ? [...specIssues, ...diagnosticIssues]
  : [...specIssues, ...inspectStructure(normalizePresentation(document))];

let report: BuildReport = {
  runtime: "node",
  output,
  slideCount: document.slides.length,
  byteLength: 0,
  diagnostics: diagnostics.map(({ severity, code, message, path: diagnosticPath }) => ({ severity, code, message, path: diagnosticPath })),
  exportWarnings: [],
  structuralIssues,
  layoutDecisions,
  packageChecks: { status: "not-run", valid: false, parts: 0, slideParts: 0, issues: [] },
  previewStatus: "not-run",
  exportStatus: "not-run",
  renderStatus: "not-run",
  generatedAt: new Date().toISOString(),
};

if (structuralIssues.some((issue) => issue.severity === "error")) {
  await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);
  throw new Error(`Deck failed pre-export checks. See ${reportPath}`);
}

const result = await generatePptx(document);
await writeFile(output, result.bytes);
const packageChecks = inspectPptxPackage(result.bytes);
const exportIssues: StructuralIssue[] = result.warnings.map((warning) => ({ severity: "error", code: `export-${warning.code}`, message: warning.message, slideId: warning.slideId }));
if (!packageChecks.valid) exportIssues.push(...packageChecks.issues.map((message) => ({ severity: "error" as const, code: "invalid-package", message })));
if (packageChecks.slideParts !== document.slides.length) exportIssues.push({ severity: "error", code: "slide-part-count", message: `Expected ${document.slides.length} slide parts, found ${packageChecks.slideParts}.` });

report = {
  ...report,
  byteLength: result.byteLength,
  exportWarnings: result.warnings,
  structuralIssues: [...structuralIssues, ...exportIssues],
  packageChecks,
  exportStatus: exportIssues.some((issue) => issue.severity === "error") ? "failed" : "generated",
};
await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);

process.stdout.write(`${JSON.stringify({ output, report: reportPath, slideCount: result.slideCount, warnings: result.warnings.length }, null, 2)}\n`);
if (report.structuralIssues.some((issue) => issue.severity === "error")) process.exitCode = 1;
