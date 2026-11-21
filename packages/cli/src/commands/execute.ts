import { applyCustomPropertyEdits, customPropertyCandidates, parseCustomPropertyEdits } from "./custom-property-edits.js";
import path from "node:path";
import { compareOpcPackages, importPptx, inspectPptxNotes, describeImportCapabilities, PackageInspectionError, PptxImportError } from "@pptkit/pptx-parser";
import { PptxFileError, writePptxFile } from "@pptkit/pptx-parser/node";
import { validatePresentation, inspectMissingDescriptions, inspectEffectiveMissingDescriptions, inspectSlideTextContrast, TextContrastLimitError, EffectiveDescriptionLimitError, PresentationValidationError } from "@pptkit/core";
import { renderPresentationToSvg, SvgRenderLimitError, SvgRenderAbortError } from "@pptkit/svg-renderer";
import type { ImportReport, ImportCapabilityReport } from "@pptkit/pptx-parser";
import { CliError, HELP, parseCommand } from "./arguments.js";
import { boundedJson, boundedText } from "./format.js";
import { readInput, checkAbort } from "./read-file.js";
import { exportAuthoring, parseAuthoring } from "./generate.js";
import { publishPreview } from "./publish-preview.js";
import { previewHtml } from "./preview-html.js";
import { applyTextEdits, parseTextEdits } from "./text-edits.js";
import { textCandidates } from "./text-candidates.js";
import { applyChartEdits, parseChartEdits } from "./chart-edits.js";
import { chartCandidates } from "./chart-candidates.js";
import { applyMetadataEdits, metadataCandidates, parseMetadataEdits } from "./core-metadata-edits.js";
import { applyVisibilityEdits, parseVisibilityEdits, visibilityCandidates } from "./slide-visibility-edits.js";
import { applyDescriptionEdits, descriptionCandidates, parseDescriptionEdits, applyNameEdits, nameCandidates, parseNameEdits, applyDecorativeEdits, decorativeCandidates, parseDecorativeEdits } from "./drawing-property-edits.js";
import { applyPictureEdit } from "./picture-edit.js";
import { packageSummary } from "./package-summary.js";
import type { CliRunResult, CliExecutionOptions } from "../index.js";
export const CLI_REPORT_SCHEMA_VERSION = "pptkit-cli-report-v1";
function failure(error: unknown): { code: string; message: string; exitCode: number; source?: unknown } {
  if (error instanceof CliError) return error;
  if (error instanceof TextContrastLimitError) return { code: "output-limit-exceeded", message: "Contrast findings exceed max-output-bytes.", exitCode: 4 };
  if (error instanceof EffectiveDescriptionLimitError) return { code: "output-limit-exceeded", message: "Description occurrences exceed max-output-bytes.", exitCode: 4 };
  if (error instanceof PackageInspectionError || error instanceof PptxImportError || error instanceof PptxFileError) return {
    code: error.code, message: error.message,
    exitCode: error.code === "operation-aborted" ? 130 : error.code.endsWith("limit-exceeded") ? 4 : 3,
    ...(error instanceof PackageInspectionError ? error.entry ? { source: { part: error.entry } } : {} : error instanceof PptxImportError && error.source ? { source: error.source } : error instanceof PptxFileError && error.temporaryPath ? { source: { temporaryPath: error.temporaryPath } } : {}) };
  return { code: "internal-error", message: "Command failed unexpectedly.", exitCode: 5 };
}
function capabilityLines(report: ImportReport): string[] {
  const counts = new Map<string, number>();
  for (const f of report.findings) { const key = `${f.operation}: ${f.tier}`; counts.set(key, (counts.get(key) ?? 0) + 1); }
  return [...counts].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0).map(([key, count]) => `  ${key}: ${count}`);
}
export async function execute(args: readonly string[], options: CliExecutionOptions): Promise<CliRunResult> {
  let json = Array.isArray(args) && args.slice(0, args.indexOf("--") < 0 ? args.length : args.indexOf("--")).includes("--json"), limit = 1024 * 1024;
  try {
    if (!options || typeof options !== "object" || Array.isArray(options) || options.cwd !== undefined && typeof options.cwd !== "string" || options.signal !== undefined && !(options.signal instanceof AbortSignal)) throw new CliError("usage", "Invalid execution options.", 2);
    const cwd = path.resolve(options.cwd ?? process.cwd()), signal = options.signal;
    const command = parseCommand(args); if (!command) return { exitCode: 0, output: HELP };
    json = command.json; limit = command.maxOutputBytes; checkAbort(signal);
    const captures = [];
    // Sequential captures keep peak memory predictable; every file gets its own input bound.
    for (const file of command.kind === "save" || command.kind === "generate" ? command.files.slice(0, 1) : command.files) captures.push(await readInput(file, cwd, command.maxInputBytes, signal));
    const textCapture = command.textEdits !== undefined ? await readInput(command.textEdits, cwd, command.maxInputBytes, signal) : undefined;
    const chartCapture = command.chartEdits !== undefined ? await readInput(command.chartEdits, cwd, command.maxInputBytes, signal) : undefined;
    const descriptionCapture = command.descriptionEdits !== undefined ? await readInput(command.descriptionEdits, cwd, command.maxInputBytes, signal) : undefined;
    const nameCapture = command.nameEdits !== undefined ? await readInput(command.nameEdits, cwd, command.maxInputBytes, signal) : undefined;
    const decorativeCapture = command.decorativeEdits !== undefined ? await readInput(command.decorativeEdits, cwd, command.maxInputBytes, signal) : undefined;
    const visibilityCapture = command.visibilityEdits !== undefined ? await readInput(command.visibilityEdits, cwd, command.maxInputBytes, signal) : undefined;
    const metadataCapture = command.metadataEdits !== undefined ? await readInput(command.metadataEdits, cwd, command.maxInputBytes, signal) : undefined;
    const customPropertyCapture = command.customPropertyEdits !== undefined ? await readInput(command.customPropertyEdits, cwd, command.maxInputBytes, signal) : undefined;
    const pictureCapture = command.pictureEdit !== undefined ? await readInput(command.pictureEdit, cwd, command.maxInputBytes, signal) : undefined;
    const imageCapture = command.pictureBytes !== undefined ? await readInput(command.pictureBytes, cwd, command.maxInputBytes, signal) : undefined;
    if (textCapture) captures.push(textCapture); if (chartCapture) captures.push(chartCapture); if (descriptionCapture) captures.push(descriptionCapture); if (nameCapture) captures.push(nameCapture); if (decorativeCapture) captures.push(decorativeCapture); if (visibilityCapture) captures.push(visibilityCapture); if (metadataCapture) captures.push(metadataCapture); if (customPropertyCapture) captures.push(customPropertyCapture); if (pictureCapture) captures.push(pictureCapture); if (imageCapture) captures.push(imageCapture);
    const inputs = captures.map(c => c.input), common = { ...(signal ? { signal } : {}), zipLimits: { maxInputBytes: command.maxInputBytes } };
    const authoringOptions = { cwd, maxAssetBytes: command.maxInputBytes, ...(command.assetBase !== undefined ? { assetBase: command.assetBase } : {}), ...(signal ? { signal } : {}) };
    if (command.kind === "preview") {
      const imported = command.previewInput === "pptx" ? await importPptx(captures[0]!.bytes, { ...common, projection: "content" }) : undefined;
      try {
        const authored = imported ? undefined : await parseAuthoring(captures[0]!.bytes, authoringOptions);
        const document = imported?.document ?? authored!.document, index = command.slideNumber === "all" ? undefined : command.slideNumber! - 1;
        if (index !== undefined && index >= document.slides.length) throw new CliError("preview-slide-out-of-range", "Requested slide does not exist.", 3);
        let rendered;
        try { rendered = await renderPresentationToSvg(document, { maxSvgBytes: limit, ...(signal ? { signal } : {}), ...(index === undefined ? {} : { slideIndices: [index] }) }); }
        catch (error) { if (error instanceof PresentationValidationError) throw new CliError("authoring-validation-failed", error.message, 3, { diagnostics: error.diagnostics }); if (error instanceof SvgRenderLimitError) throw new CliError("output-limit-exceeded", "Rendered SVG exceeds max-output-bytes.", 4); if (error instanceof SvgRenderAbortError) throw new CliError("operation-aborted", "Command cancelled.", 130); throw error; }
        checkAbort(signal);
        if (command.previewOutput !== undefined) {
          const destination = path.resolve(cwd, command.previewOutput), displayPath = path.relative(cwd, destination).split(path.sep).join("/");
          const all = index === undefined;
          const content = all ? previewHtml(rendered.slides, rendered.width, rendered.height, limit) : rendered.slides[0]!.svg;
          const slideMetadata = rendered.slides.map(({ svg: _svg, ...slide }) => slide);
          const base = { schemaVersion: CLI_REPORT_SCHEMA_VERSION, command: "preview", status: "ok", inputs, ...(authored?.assetInputs.length ? { assetInputs: authored.assetInputs } : {}), nativeVerification: "not-performed", destination: displayPath };
          const report = { scope: "svg-static", inputFormat: command.previewInput, width: rendered.width, height: rendered.height, ...(all ? { slides: slideMetadata } : { slide: slideMetadata[0] }), warnings: rendered.warnings, status: rendered.status, ...(imported ? { importReport: imported.report } : {}), publication: { status: "committed", byteLength: Buffer.byteLength(content), ...(all ? { format: "html" } : {}) } };
          const fallback = { ...base, report: { publication: { status: "committed", detailsTruncated: true } } };
          boundedJson(fallback, limit); boundedJson({ ...base, report }, limit);
          const published = await publishPreview(destination, content, signal);
          try { return { exitCode: 0, output: boundedJson({ ...base, report: { ...report, publication: { ...report.publication, ...(published.temporaryPath ? { temporaryPath: published.temporaryPath } : {}) } } }, limit) }; }
          catch { return { exitCode: 0, output: boundedJson(fallback, limit) }; }
        }
        const report = { scope: "svg-static", inputFormat: command.previewInput, width: rendered.width, height: rendered.height, ...(index === undefined ? { slides: rendered.slides } : { slide: rendered.slides[0] }), warnings: rendered.warnings, status: rendered.status, ...(imported ? { importReport: imported.report } : {}) };
        return { exitCode: 0, output: boundedJson({ schemaVersion: CLI_REPORT_SCHEMA_VERSION, command: "preview", status: "ok", inputs, ...(authored?.assetInputs.length ? { assetInputs: authored.assetInputs } : {}), nativeVerification: "not-performed", report }, limit) };
      } finally { imported?.dispose(); }
    }
    if (command.kind === "validate") {
      const authored = await parseAuthoring(captures[0]!.bytes, authoringOptions), document = authored.document, diagnostics = validatePresentation(document);
      checkAbort(signal);
      const coreValid = !diagnostics.some(d => d.severity === "error");
      const descriptions = command.checkDescriptions ? coreValid ? { status: "checked" as const, ...(command.descriptionOccurrences ? inspectEffectiveMissingDescriptions(document, { maxFindings: Math.floor(limit / 32) }) : inspectMissingDescriptions(document)) } : { status: "skipped-invalid-core" as const } : undefined;
      const contrast = command.checkContrast ? coreValid ? { status: "checked" as const, ...inspectSlideTextContrast(document, { maxFindings: Math.floor(limit / 32) }) } : { status: "skipped-invalid-core" as const } : undefined;
      checkAbort(signal);
      let exported: { status: "checked"; slideCount: number; byteLength: number; warnings: Awaited<ReturnType<typeof exportAuthoring>>["warnings"] } | { status: "skipped-invalid-core" } | undefined;
      if (command.validateExport) {
        if (!coreValid) exported = { status: "skipped-invalid-core" };
        else {
          const result = await exportAuthoring(document, signal, { imageDiagnostics: command.imageDiagnostics, validateImageData: command.validateImageData });
          const session = await importPptx(result.bytes, { projection: "structure", ...(signal ? { signal } : {}) });
          try { checkAbort(signal); exported = { status: "checked", slideCount: result.slideCount, byteLength: result.byteLength, warnings: result.warnings }; }
          finally { session.dispose(); }
        }
      }
      const valid = coreValid && !(command.failOnExportWarning && exported?.status === "checked" && exported.warnings.length > 0) && !(command.failOnMissingDescription && descriptions?.status === "checked" && descriptions.findings.length > 0) && !(command.failOnContrast && contrast?.status === "checked" && contrast.findings.some(f => f.status !== "meets-threshold"));
      const report = { valid, diagnostics, scope: command.validateExport ? "core-and-export" : "core-authoring", ...(command.failOnExportWarning ? { failOnExportWarning: true } : {}), ...(command.failOnMissingDescription ? { failOnMissingDescription: true } : {}), ...(command.failOnContrast ? { failOnContrast: true } : {}), ...(descriptions ? { descriptions } : {}), ...(contrast ? { contrast } : {}), ...(exported ? { export: exported } : {}) };
      const body = { schemaVersion: CLI_REPORT_SCHEMA_VERSION, command: "validate", status: valid ? "ok" : "invalid", inputs, ...(authored.assetInputs.length ? { assetInputs: authored.assetInputs } : {}), nativeVerification: "not-performed", report };
      const lines = [`PPTKit validate: ${inputs[0]!.path}`, `Valid: ${valid}`, `Diagnostics: ${diagnostics.length}`, ...diagnostics.map(d => `  ${d.severity} ${d.code} ${d.path}: ${d.message}`), ...(descriptions ? [`Description check: ${descriptions.status}`, ...(descriptions.status === "checked" ? descriptions.findings.map(f => `  ${f.code} ${f.path} (${f.elementId})${'source' in f ? ` slide=${f.slideId} source=${f.source.kind}:${f.source.id}` : ''}`) : [])] : []), ...(contrast ? [`Contrast check: ${contrast.status}`, ...(contrast.status === "checked" ? contrast.findings.map(f => `  ${f.status} ${f.path} (${f.elementId})${f.source ? ` slide=${f.slideId} source=${f.source.kind}:${f.source.id}` : ""}`) : [])] : []), ...(exported ? [`Export: ${exported.status}`, ...(exported.status === "checked" ? exported.warnings.map(w => `Export warning: ${JSON.stringify(w)}`) : []), "Native application checks were not performed."] : ["Core authoring validation only; export and native application checks were not performed."])];
      return { exitCode: valid ? 0 : 1, output: json ? boundedJson(body, limit) : boundedText(lines, limit) };
    }
    if (command.kind === "generate") {
      const authored = await parseAuthoring(captures[0]!.bytes, authoringOptions);
      const destination = path.resolve(cwd, command.files[1]!), displayPath = path.relative(cwd, destination).split(path.sep).join("/");
      const base = { schemaVersion: CLI_REPORT_SCHEMA_VERSION, command: "generate", status: "ok", inputs, ...(authored.assetInputs.length ? { assetInputs: authored.assetInputs } : {}), nativeVerification: "not-performed", destination: displayPath };
      const fallback = { ...base, report: { status: "committed", directorySynced: false, detailsTruncated: true } };
      const fallbackText = [`PPTKit generate committed: ${displayPath}`, "Directory sync or report details may be omitted; inspect output."];
      if (json) boundedJson(fallback, limit); else boundedText(fallbackText, limit);
      const generated = await exportAuthoring(authored.document, signal, { imageDiagnostics: command.imageDiagnostics, validateImageData: command.validateImageData });
      checkAbort(signal);
      if (command.failOnExportWarning && generated.warnings.length) throw new CliError("export-warnings-rejected", "Export warnings were rejected before publication.", 3, { warnings: generated.warnings });
      const session = await importPptx(generated.bytes, { projection: "structure", ...(signal ? { signal } : {}) });
      try {
        const prepared = await session.prepareSave({ mode: "no-op", ...(signal ? { signal } : {}) });
        if (prepared.status !== "ready") throw new CliError("generate-publication-blocked", "Generated package could not be prepared for publication.", 3);
        const written = await writePptxFile(session, prepared.plan, destination, { mode: "create", ...(signal ? { signal } : {}) });
        const report = { status: written.status, slideCount: generated.slideCount, byteLength: generated.byteLength, exportWarnings: generated.warnings, directorySynced: written.directorySynced, warnings: written.warnings };
        try { return { exitCode: 0, output: json ? boundedJson({ ...base, report }, limit) : boundedText([`PPTKit generate committed: ${displayPath}`, `Slides: ${generated.slideCount}`, `Bytes: ${generated.byteLength}`, `Directory synced: ${written.directorySynced}`, ...generated.warnings.map(w => `Export warning: ${JSON.stringify(w)}`), ...written.warnings.map(w => `Warning: ${JSON.stringify(w)}`)], limit) }; }
        catch { return { exitCode: 0, output: json ? boundedJson(fallback, limit) : boundedText(fallbackText, limit) }; }
      } finally { session.dispose(); }
    }
    if (command.kind === "save") {
      const destination = path.resolve(cwd, command.files[1]!);
      const displayPath = path.relative(cwd, destination).split(path.sep).join("/");
      const edit = command.saveMode === "edit" ? { ...(command.textEdits !== undefined ? { textEdits: true } : {}), ...(command.chartEdits !== undefined ? { chartEdits: true } : {}), ...(command.descriptionEdits !== undefined ? { descriptionEdits: true } : {}), ...(command.nameEdits !== undefined ? { nameEdits: true } : {}), ...(command.decorativeEdits !== undefined ? { decorativeEdits: true } : {}), ...(command.visibilityEdits !== undefined ? { visibilityEdits: true } : {}), ...(command.metadataEdits !== undefined ? { metadataEdits: true } : {}), ...(command.customPropertyEdits !== undefined ? { customPropertyEdits: true } : {}), ...(command.pictureEdit !== undefined ? { pictureEdit: true } : {}), ...(command.firstSlideNumber !== undefined ? { firstSlideNumber: command.firstSlideNumber } : {}) } : undefined;
      const base = { schemaVersion: CLI_REPORT_SCHEMA_VERSION, command: "save", status: "ok", inputs, nativeVerification: "not-performed", destination: displayPath, ...(edit ? { edit } : {}) };
      const assurance = command.saveMode === "edit" ? "scoped-preservation" : "original-bytes";
      // Ensure a committed result always fits before any output is published.
      const fallback = { ...base, report: { status: "committed", assurance, directorySynced: false, detailsTruncated: true } };
      const editText = [...(command.textEdits !== undefined ? ["Text edits applied."] : []), ...(command.chartEdits !== undefined ? ["Chart edits applied."] : []), ...(command.descriptionEdits !== undefined ? ["Description edits applied."] : []), ...(command.nameEdits !== undefined ? ["Name edits applied."] : []), ...(command.decorativeEdits !== undefined ? ["Decorative edits applied."] : []), ...(command.visibilityEdits !== undefined ? ["Visibility edits applied."] : []), ...(command.metadataEdits !== undefined ? ["Metadata edits applied."] : []), ...(command.customPropertyEdits !== undefined ? ["Custom property edits applied."] : []), ...(command.pictureEdit !== undefined ? ["Picture edit applied."] : []), ...(command.firstSlideNumber !== undefined ? [`First slide number: ${command.firstSlideNumber === null ? "clear" : command.firstSlideNumber}`] : [])];
      const fallbackText = [`PPTKit save committed: ${displayPath}`, `Assurance: ${assurance}`, ...editText, "Directory sync or report details may be omitted; inspect output."];
      if (!command.dryRun) { if (json) boundedJson(fallback, limit); else boundedText(fallbackText, limit); }
      const textPlan = textCapture ? parseTextEdits(textCapture.bytes, inputs[0]!.sha256) : undefined, chartPlan = chartCapture ? parseChartEdits(chartCapture.bytes, inputs[0]!.sha256) : undefined, descriptionPlan = descriptionCapture ? parseDescriptionEdits(descriptionCapture.bytes, inputs[0]!.sha256) : undefined, namePlan = nameCapture ? parseNameEdits(nameCapture.bytes, inputs[0]!.sha256) : undefined, decorativePlan = decorativeCapture ? parseDecorativeEdits(decorativeCapture.bytes, inputs[0]!.sha256) : undefined, visibilityPlan = visibilityCapture ? parseVisibilityEdits(visibilityCapture.bytes, inputs[0]!.sha256) : undefined, metadataPlan = metadataCapture ? parseMetadataEdits(metadataCapture.bytes, inputs[0]!.sha256) : undefined;
      const customPropertyPlan = customPropertyCapture ? parseCustomPropertyEdits(customPropertyCapture.bytes, inputs[0]!.sha256) : undefined;
      if ((customPropertyPlan?.version === "pptkit-custom-property-edits-v6" || customPropertyPlan?.version === "pptkit-custom-property-edits-v7") && (chartPlan || descriptionPlan || namePlan || decorativePlan || pictureCapture || command.firstSlideNumber !== undefined)) throw new CliError("custom-property-edits-incompatible", "Exact custom-property edits cannot be combined with chart, drawing, picture or starting-number edits.", 3);
      if ((descriptionPlan || namePlan || decorativePlan) && textPlan && textPlan.version !== "pptkit-text-edits-v3") throw new CliError(descriptionPlan ? "description-edits-incompatible" : namePlan ? "name-edits-incompatible" : "decorative-edits-incompatible", "Combined drawing property and text edits require a v3 text manifest.", 3);
      const session = await importPptx(captures[0]!.bytes, command.saveMode === "edit" ? { ...common, projection: "content", ...(descriptionPlan || namePlan || decorativePlan ? { editSaveProfile: "text-runs-v3" as const, descriptionEditSaveProfile: "mixed-v5" as const, ...(namePlan ? { drawingNameSaveProfile: "existing-v1" as const } : {}), ...(decorativePlan ? { decorativeSaveProfile: "existing-v1" as const } : {}) } : textPlan ? { editSaveProfile: textPlan.version === "pptkit-text-edits-v3" ? "text-runs-v3" : textPlan.version === "pptkit-text-edits-v2" ? "text-runs-v2" : "text-runs-v1" } : {}), ...(chartPlan ? { chartEditSaveProfile: chartPlan.version === "v6" ? "existing-values-and-series-names-v5" : chartPlan.version === "v5" ? "existing-values-and-labels-v4" : chartPlan.version === "v4" ? "existing-values-v3" : chartPlan.version === "v3" ? "existing-values-v2" : "existing-values-v1" } : {}), ...(visibilityPlan ? { slideVisibilitySaveProfile: "attribute-v1" } : {}), ...(metadataPlan ? { metadataSaveProfile: metadataPlan.version === "pptkit-metadata-edits-v17" ? "metadata-default-remove-v17" : metadataPlan.version === "pptkit-metadata-edits-v16" ? "metadata-revision-remove-v16" : metadataPlan.version === "pptkit-metadata-edits-v15" ? "metadata-keywords-remove-v15" : metadataPlan.version === "pptkit-metadata-edits-v14" ? "metadata-date-remove-v14" : metadataPlan.version === "pptkit-metadata-edits-v13" ? "metadata-remove-v13" : metadataPlan.version === "pptkit-metadata-edits-v12" ? "extended-create-v12" : metadataPlan.version === "pptkit-metadata-edits-v11" ? "extended-company-insert-v11" : metadataPlan.version === "pptkit-metadata-edits-v10" ? "core-keywords-insert-v10" : metadataPlan.version === "pptkit-metadata-edits-v9" ? "core-revision-insert-v9" : metadataPlan.version === "pptkit-metadata-edits-v8" ? "core-create-v8" : metadataPlan.version === "pptkit-metadata-edits-v7" ? "core-date-insert-v7" : metadataPlan.version === "pptkit-metadata-edits-v6" ? "core-date-values-v6" : metadataPlan.version === "pptkit-metadata-edits-v5" ? "core-insert-v5" : metadataPlan.version === "pptkit-metadata-edits-v4" ? "extended-company-v4" : metadataPlan.version === "pptkit-metadata-edits-v3" ? "core-keywords-v3" : metadataPlan.version === "pptkit-metadata-edits-v2" ? "core-values-v2" : "core-strings-v1" } : {}), ...(customPropertyPlan?.version === "pptkit-custom-property-edits-v6" ? { customCurrencyProjection: "exact-save-v2" as const } : {}), ...(customPropertyPlan?.version === "pptkit-custom-property-edits-v7" ? { customDecimalProjection: "exact-save-v3" as const } : {}), ...(customPropertyPlan ? { customPropertySaveProfile: customPropertyPlan.version === "pptkit-custom-property-edits-v5" || customPropertyPlan.version === "pptkit-custom-property-edits-v7" ? "create-and-append-v5" : customPropertyPlan.version === "pptkit-custom-property-edits-v4" ? "append-and-remove-v4" : customPropertyPlan.version === "pptkit-custom-property-edits-v3" ? "remove-v3" : customPropertyPlan.version === "pptkit-custom-property-edits-v2" ? "scalar-and-name-v2" : "scalar-v1" } : {}), ...(pictureCapture ? { pictureEditSaveProfile: "replace-v1" } : {}), ...(command.firstSlideNumber !== undefined ? { firstSlideNumberSaveProfile: "attribute-v2" } : {}) } : { ...common, projection: "structure" });
      try {
        if (textPlan) {
          applyTextEdits(session, textPlan, signal);
        }
        if (chartPlan) applyChartEdits(session, chartPlan.edits, signal, chartPlan.version);
        if (descriptionPlan) applyDescriptionEdits(session, descriptionPlan, signal);
        if (namePlan) applyNameEdits(session, namePlan, signal);
        if (decorativePlan) applyDecorativeEdits(session, decorativePlan, signal);
        if (visibilityPlan) applyVisibilityEdits(session, visibilityPlan, signal);
        const coreMetadataChanges = metadataPlan ? applyMetadataEdits(session, metadataPlan, signal) : { insertions: [], removals: [] };
        const customPropertyChanges = customPropertyPlan ? applyCustomPropertyEdits(session, customPropertyPlan, signal) : { removals: [], additions: [] };
        if (pictureCapture && imageCapture) applyPictureEdit(session, pictureCapture.bytes, imageCapture.bytes, inputs[0]!.sha256, imageCapture.input.sha256, signal);
        if (command.firstSlideNumber !== undefined) {
          if (!session.document.setFirstSlideNumber) throw new CliError("save-unsupported", "Imported document cannot edit its starting number.", 3);
          session.document.setFirstSlideNumber(command.firstSlideNumber === null ? undefined : command.firstSlideNumber);
        }
        const prepared = await session.prepareSave({ mode: command.saveMode!, ...(coreMetadataChanges.insertions.length ? { coreMetadataInsertions: coreMetadataChanges.insertions } : {}), ...(coreMetadataChanges.removals.length ? { coreMetadataRemovals: coreMetadataChanges.removals } : {}), ...(customPropertyChanges.removals.length ? { customPropertyRemovals: customPropertyChanges.removals } : {}), ...(customPropertyChanges.additions.length ? { customPropertyAdditions: customPropertyChanges.additions } : {}), ...(signal ? { signal } : {}) });
        if (command.dryRun) {
          checkAbort(signal);
          const planned = prepared.status === "ready" ? prepared.plan.report : prepared.report;
          const result = { ...base, status: prepared.status, dryRun: true, report: planned };
          const lines = [`PPTKit save dry run: ${displayPath}`, `Plan: ${prepared.status}`, `Assurance: ${planned.assurance}`, ...(planned.changes ?? []).map(change => `Change: ${JSON.stringify(change)}`), ...planned.losses.map(loss => `Loss: ${JSON.stringify(loss)}`), ...planned.diagnostics.map(diagnostic => `Diagnostic: ${JSON.stringify(diagnostic)}`), "No destination was published; filesystem publication was not checked."];
          return { exitCode: prepared.status === "ready" ? 0 : 1, output: json ? boundedJson(result, limit) : boundedText(lines, limit) };
        }
        if (prepared.status !== "ready") { const diagnostic = prepared.report.diagnostics[0]; throw new CliError(diagnostic?.code ?? "save-blocked", diagnostic?.message ?? "Preservation plan was blocked.", 3, diagnostic?.source); }
        const result = await writePptxFile(session, prepared.plan, destination, { mode: "create", ...(signal ? { signal } : {}) });
        const report = { status: result.status, assurance: result.report.assurance, directorySynced: result.directorySynced, warnings: result.warnings, changes: result.report.changes };
        try { return { exitCode: 0, output: json ? boundedJson({ ...base, report }, limit) : boundedText([`PPTKit save committed: ${displayPath}`, `Assurance: ${result.report.assurance}`, ...editText, `Directory synced: ${result.directorySynced}`, ...result.warnings.map(w => `Warning: ${JSON.stringify(w)}`)], limit) }; }
        catch { return { exitCode: 0, output: json ? boundedJson(fallback, limit) : boundedText(fallbackText, limit) }; }
      } finally { session.dispose(); }
    }
    let capabilities: ImportCapabilityReport | undefined;
    let textEditCandidates: ReturnType<typeof textCandidates> | undefined;
    let chartEditCandidates: ReturnType<typeof chartCandidates> | undefined;
    let chartEditCandidatesV2: ReturnType<typeof chartCandidates> | undefined;
    let chartEditCandidatesV6: ReturnType<typeof chartCandidates> | undefined;
    let chartEditCandidatesV5: ReturnType<typeof chartCandidates> | undefined;
    let chartEditCandidatesV4: ReturnType<typeof chartCandidates> | undefined;
    let chartEditCandidatesV3: ReturnType<typeof chartCandidates> | undefined;
    let descriptionEditCandidates: ReturnType<typeof descriptionCandidates> | undefined;
    let nameEditCandidates: ReturnType<typeof nameCandidates> | undefined;
    let decorativeEditCandidates: ReturnType<typeof decorativeCandidates> | undefined;
    let visibilityEditCandidates: ReturnType<typeof visibilityCandidates> | undefined;
    let metadataEditCandidates: ReturnType<typeof metadataCandidates> | undefined;
    let customPropertyEditCandidates: ReturnType<typeof customPropertyCandidates> | undefined;
    let report, exitCode: number, lines: string[];
    if (command.kind === "inspect") {
      const session = await importPptx(captures[0]!.bytes, { ...common, projection: command.projection,
        ...(command.diagramPreview ? { diagramPreview: command.diagramPreview } : {}), ...(command.diagramTextComparison ? { diagramTextComparison: command.diagramTextComparison } : {}),
        ...((command.descriptionValues || command.nameValues || command.decorativeValues) && (command.editSaveProfile || command.chartEditSaveProfile || Number(command.descriptionValues) + Number(command.nameValues) + Number(command.decorativeValues) + Number(command.visibilityValues) + Number(command.metadataValues) + Number(command.customPropertyValues) > 1) ? { editSaveProfile: "text-runs-v3" as const, descriptionEditSaveProfile: "mixed-v5" as const, ...(command.nameValues ? { drawingNameSaveProfile: "existing-v1" as const } : {}), ...(command.decorativeValues ? { decorativeSaveProfile: "existing-v1" as const } : {}) } : command.descriptionValues ? { descriptionEditSaveProfile: "batch-v3" as const } : command.nameValues ? { drawingNameSaveProfile: "existing-v1" as const } : command.decorativeValues ? { decorativeSaveProfile: "existing-v1" as const } : command.editSaveProfile ? { editSaveProfile: command.editSaveProfile } : {}), ...(command.chartEditSaveProfile ? { chartEditSaveProfile: command.chartEditSaveProfile } : {}), ...(command.visibilityValues ? { slideVisibilitySaveProfile: "attribute-v1" } : {}), ...(command.metadataValues ? { metadataSaveProfile: command.metadataDefaultRemove ? "metadata-default-remove-v17" : command.metadataRevisionRemove ? "metadata-revision-remove-v16" : command.metadataKeywordsRemove ? "metadata-keywords-remove-v15" : command.metadataDateRemove ? "metadata-date-remove-v14" : command.metadataRemove ? "metadata-remove-v13" : command.metadataExtendedCreate ? "extended-create-v12" : command.metadataCompanyInsert ? "extended-company-insert-v11" : command.metadataKeywordsInsert ? "core-keywords-insert-v10" : command.metadataRevisionInsert ? "core-revision-insert-v9" : command.metadataCreate ? "core-create-v8" : command.metadataDateInsert ? "core-date-insert-v7" : command.metadataDates ? "core-date-values-v6" : command.metadataInsert ? "core-insert-v5" : command.metadataCompany ? "extended-company-v4" : command.metadataKeywords ? "core-keywords-v3" : command.metadataRevision ? "core-values-v2" : "core-strings-v1" } : {}), ...(command.customPropertyCurrency ? { customCurrencyProjection: "exact-save-v2" as const } : {}), ...(command.customPropertyDecimal ? { customDecimalProjection: "exact-save-v3" as const } : {}), ...(command.customPropertyValues ? { customPropertySaveProfile: command.customPropertyCreate || command.customPropertyDecimal ? "create-and-append-v5" : command.customPropertyAdditions ? "append-and-remove-v4" : command.customPropertyDeletions ? "remove-v3" : command.customPropertyNames ? "scalar-and-name-v2" : "scalar-v1" } : {}), ...(command.pictureSourceProfile ? { pictureSourceProfile: command.pictureSourceProfile } : {}) });
      try { report = session.report;
        if (command.capabilities) capabilities = describeImportCapabilities(report, { limits: { maxSubjects: command.maxCapabilitySubjects }, filter: { ...(command.capabilityFeature !== undefined ? { feature: command.capabilityFeature } : {}), ...(command.capabilityPart !== undefined ? { part: command.capabilityPart } : {}) } });
        if (command.textValues) textEditCandidates = textCandidates(session, command.editSaveProfile!, inputs[0]!.sha256, signal);
        if (command.chartValues) {
          if (command.chartEditSaveProfile === "existing-values-and-series-names-v5") chartEditCandidatesV6 = chartCandidates(session, inputs[0]!.sha256, signal, "v6");
          else if (command.chartEditSaveProfile === "existing-values-and-labels-v4") chartEditCandidatesV5 = chartCandidates(session, inputs[0]!.sha256, signal, "v5");
          else if (command.chartEditSaveProfile === "existing-values-v3") chartEditCandidatesV4 = chartCandidates(session, inputs[0]!.sha256, signal, "v4");
          else if (command.chartEditSaveProfile === "existing-values-v2") chartEditCandidatesV3 = chartCandidates(session, inputs[0]!.sha256, signal, "v3");
          else {
            chartEditCandidates = chartCandidates(session, inputs[0]!.sha256, signal);
            if (session.report.chartValues?.some(binding => { checkAbort(signal); return binding.coordinate !== undefined; })) chartEditCandidatesV2 = chartCandidates(session, inputs[0]!.sha256, signal, "v2");
          }
        }
        if (command.descriptionValues) descriptionEditCandidates = descriptionCandidates(session, inputs[0]!.sha256, signal);
        if (command.nameValues) nameEditCandidates = nameCandidates(session, inputs[0]!.sha256, signal);
        if (command.decorativeValues) decorativeEditCandidates = decorativeCandidates(session, inputs[0]!.sha256, signal);
        if (command.visibilityValues) visibilityEditCandidates = visibilityCandidates(session, inputs[0]!.sha256, signal);
        if (command.metadataValues) metadataEditCandidates = metadataCandidates(session, inputs[0]!.sha256, signal, command.metadataRevision, command.metadataKeywords, command.metadataCompany, command.metadataInsert, command.metadataDates, command.metadataDateInsert, command.metadataCreate, command.metadataRevisionInsert, command.metadataKeywordsInsert, command.metadataCompanyInsert, command.metadataExtendedCreate, command.metadataRemove, command.metadataDateRemove, command.metadataKeywordsRemove, command.metadataRevisionRemove, command.metadataDefaultRemove);
        if (command.customPropertyValues) customPropertyEditCandidates = customPropertyCandidates(session, inputs[0]!.sha256, signal, command.customPropertyNames, command.customPropertyDeletions, command.customPropertyAdditions, command.customPropertyCreate, command.customPropertyCurrency, command.customPropertyDecimal);
      } finally { session.dispose(); }
      exitCode = 0;
      lines = [`PPTKit inspect: ${inputs[0]!.path}`, `Projection: ${report.projection}`, `Slides: ${report.inspection.slides.length}`, `Package parts: ${report.inspection.package.parts.length}`,
        `Capability findings: ${report.findings.length}`, ...capabilityLines(report),
        ...report.findings.map(f => `  ${f.feature} | ${f.operation} | ${f.tier} | ${f.source.part}${f.source.path ? ` ${f.source.path}` : ""} | ${f.reason}`),
        `Diagnostics: ${report.diagnostics.length}`, ...report.diagnostics.map(d => `  ${d.severity} ${d.code}${d.source ? ` ${d.source.part}` : ""}: ${d.message}`),
        "Findings describe the selected projection/profile; absent findings do not establish support.", "Native application verification: not performed. Source is unchanged."];
    } else {
      report = await compareOpcPackages(captures[0]!.bytes, captures[1]!.bytes, { ...common, profile: command.profile, comparisonLimits: { maxDifferences: command.maxDifferences } });
      exitCode = report.equal ? 0 : 1;
      lines = [`PPTKit compare: ${inputs[0]!.path} -> ${inputs[1]!.path}`, `Profile: ${report.profile}`, `Equal: ${report.equal}`, `Differences: ${report.differenceCount}`, `Truncated: ${report.truncated}`,
        ...report.differences.map(d => `  ${d.code}: ${d.part}${d.relationshipId ? ` [${d.relationshipId}]` : ""}${d.path ? ` ${d.path}` : ""}`),
        "Equality is profile-scoped; it does not prove visual parity or native editability."];
    }
    if (capabilities) lines = [`PPTKit capability view: ${inputs[0]!.path}`, `Projection: ${capabilities.projection}`, `Subjects: ${capabilities.returnedSubjects}/${capabilities.totalSubjects}`, `Truncated: ${capabilities.truncated}`,
      ...capabilities.subjects.flatMap(s => [`${s.id}: ${s.feature} | ${s.source.part}${s.source.path ? ` ${s.source.path}` : ""}`,
        ...s.operations.map(o => `  ${o.operation}: ${o.status}${o.tier ? ` (${o.tier})` : ""}${o.observations.length ? ` | ${o.observations.map(f => f.reason).join("; ")}` : " | No observation; no support is inferred."}`),
        ...(s.references ? [`  Tasks: ${s.references.tasks.join(", ")}`] : [])]),
      "Reference artifacts do not verify this input; native application and candidate-save validation were not performed."];
    const notesInventory = command.kind === "inspect" && command.notesInventory ? await inspectPptxNotes(captures[0]!.bytes, { ...common, notesLimits: command.notesLimits }) : undefined;
    if (notesInventory && !json) {
      lines.push(`Notes source inventory: ${notesInventory.pages.length} slide(s), ${notesInventory.parts.length} notes/master part(s)`,
        `Notes size: ${notesInventory.notesSize.widthEmu} x ${notesInventory.notesSize.heightEmu} EMU`,
        `Inventory findings: ${notesInventory.findings.length}; structural findings: ${notesInventory.structureFindings.length}`);
      for (const page of notesInventory.pages) {
        checkAbort(signal);
        lines.push(`  Slide ${page.slideId} | ${page.slidePart} | notes: ${page.notesPart ?? "absent"} | master: ${page.masterPart ?? "absent"} | body: ${page.body.status} | placeholder matches: ${page.placeholderMatches.length}`);
      }
      for (const part of notesInventory.parts) {
        checkAbort(signal);
        lines.push(`  ${part.kind} ${part.part} | drawings: ${part.drawings.length} | geometry records: ${part.geometries.length} | field records: ${part.fields.length}`);
      }
      lines.push("Notes geometry is source-only; fields were not evaluated. Use --json for complete records.");
    }
    checkAbort(signal);
    const envelope = { schemaVersion: CLI_REPORT_SCHEMA_VERSION, command: command.kind, status: exitCode ? "different" : "ok", inputs, nativeVerification: "not-performed", ...(capabilities ? { reportKind: "capabilities" } : {}), ...(command.packageSummary && command.kind === "inspect" ? { packageSummary: packageSummary((report as ImportReport).inspection.package) } : {}), ...(notesInventory ? { notesInventory } : {}), ...(textEditCandidates ? { textEditCandidates } : {}), ...(chartEditCandidates ? { chartEditCandidates } : {}), ...(chartEditCandidatesV2 ? { chartEditCandidatesV2 } : {}), ...(chartEditCandidatesV3 ? { chartEditCandidatesV3 } : {}), ...(chartEditCandidatesV4 ? { chartEditCandidatesV4 } : {}), ...(chartEditCandidatesV5 ? { chartEditCandidatesV5 } : {}), ...(chartEditCandidatesV6 ? { chartEditCandidatesV6 } : {}), ...(descriptionEditCandidates ? { descriptionEditCandidates } : {}), ...(nameEditCandidates ? { nameEditCandidates } : {}), ...(decorativeEditCandidates ? { decorativeEditCandidates } : {}), ...(visibilityEditCandidates ? { visibilityEditCandidates } : {}), ...(metadataEditCandidates ? { metadataEditCandidates } : {}), ...(customPropertyEditCandidates ? { customPropertyEditCandidates } : {}), report: capabilities ?? report };
    return { exitCode, output: json ? boundedJson(envelope, limit) : boundedText(lines, limit) };
  } catch (error) {
    let f = failure(error);
    let output;
    try { output = json ? boundedJson({ schemaVersion: CLI_REPORT_SCHEMA_VERSION, status: "error", error: { code: f.code, message: f.message, ...(f.source ? { source: f.source } : {}) } }, limit) : boundedText([`PPTKit error (${f.code}): ${f.message}`, ...(f.source ? [`Source: ${JSON.stringify(f.source)}`] : [])], limit); }
    catch { f = failure(new CliError("output-limit-exceeded", "Report exceeds max-output-bytes; increase the limit.", 4)); output = json ? JSON.stringify({ schemaVersion: CLI_REPORT_SCHEMA_VERSION, status: "error", error: { code: f.code, message: f.message } }) : `PPTKit error (${f.code}): ${f.message}`; }
    return { exitCode: f.exitCode, output, channel: "stderr" };
  }
}
