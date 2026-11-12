import { createPresentation, PresentationValidationError } from "@pptkit/core";
import type { PresentationInit, PresentationSlideInput, PresentationAssetInput, SlideLayoutInput } from "@pptkit/core";
import { generatePptx, PptxExportError } from "@pptkit/pptx-exporter";
import { CliError } from "./arguments.js";
import path from "node:path";
import { readInput } from "./read-file.js";

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** The only optional I/O is explicitly requested local asset capture. */
export async function parseAuthoring(input: Uint8Array, options: { cwd: string; assetBase?: string; maxAssetBytes: number; signal?: AbortSignal }) {
  let data: unknown;
  try { data = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(input)); }
  catch { throw new CliError("authoring-invalid-json", "Authoring input must be valid UTF-8 JSON.", 3); }
  if (!record(data) || data.version !== "pptkit-authoring-v1" || Object.keys(data).some(k => !["version", "presentation", "assets", "layouts", "slides"].includes(k))) throw new CliError("authoring-invalid", "Expected a pptkit-authoring-v1 object with only documented top-level fields.", 3);
  if (data.presentation !== undefined && !record(data.presentation) || !Array.isArray(data.slides) || data.slides.some(s => !record(s)) || data.assets !== undefined && (!Array.isArray(data.assets) || data.assets.some(a => !record(a))) || data.layouts !== undefined && (!Array.isArray(data.layouts) || data.layouts.some(l => !record(l)))) throw new CliError("authoring-invalid", "Presentation must be an object; slides, assets and layouts must be arrays of objects.", 3);
  const known = (item: Record<string, unknown>, keys: readonly string[]) => Object.keys(item).every(key => keys.includes(key));
  if (record(data.presentation) && !known(data.presentation, ["id", "firstSlideNumber", "metadata", "size", "theme", "masters", "textStylePresets"]) || (data.slides as Record<string, unknown>[]).some(s => !known(s, ["id", "layoutId", "background", "elements", "notes", "hidden", "section", "tags", "customData", "showInheritedShapes", "transition", "eventTimeline", "readingOrder"])) || ((data.layouts ?? []) as Record<string, unknown>[]).some(l => !known(l, ["id", "name", "masterId", "showMasterShapes", "background", "elements", "placeholders"])) || ((data.assets ?? []) as Record<string, unknown>[]).some(a => !known(a, ["id", "kind", "source", "mimeType", "width", "height", "accessibility", "dedupeKey"]))) throw new CliError("authoring-invalid", "Unknown presentation, slide, asset or layout field.", 3);
  // Avoid passing malformed containers to Core, whose typed authoring API
  // assumes a box is either omitted or an object. Bound this traversal.
  const pending: { value: unknown; depth: number }[] = [];
  for (const owner of [...(data.slides as Record<string, unknown>[]), ...((data.layouts ?? []) as Record<string, unknown>[]), ...(record(data.presentation) && Array.isArray(data.presentation.masters) ? data.presentation.masters.filter(record) : [])]) {
    if (owner.elements !== undefined) pending.push({ value: owner.elements, depth: 0 });
  }
  let visited = 0;
  while (pending.length) {
    const { value, depth } = pending.pop()!;
    if (depth > 128 || !Array.isArray(value)) throw new CliError("authoring-invalid", "Element collections must be arrays within the nesting limit.", 3);
    for (const element of value) {
      if (++visited > 100000) throw new CliError("authoring-limit-exceeded", "Authoring input exceeds the element limit.", 4);
      if (!record(element) || element.box !== undefined && !record(element.box)) throw new CliError("authoring-invalid", "Each element must be an object with an optional box object.", 3);
      if (element.type === "connector" && (!record(element.start) || !record(element.end))) throw new CliError("authoring-invalid", "Connector endpoints must be objects.", 3);
      if (element.type === "table" && (!Array.isArray(element.columns) || !Array.isArray(element.rows))) throw new CliError("authoring-invalid", "Table columns and rows must be arrays.", 3);
      if (element.type === "group" && element.children !== undefined) pending.push({ value: element.children, depth: depth + 1 });
    }
  }
  const assets = (data.assets ?? []) as Record<string, unknown>[];
  const assetInputs: { path: string; byteLength: number; sha256: string }[] = [];
  let remaining = options.maxAssetBytes;
  for (const asset of assets) {
    if (record(asset.source) && asset.source.type === "path") {
      const source = asset.source;
      if (options.assetBase === undefined || Object.keys(source).some(k => k !== "type" && k !== "value") || typeof source.value !== "string" || !source.value || source.value.includes("\\") || source.value.includes("\0") || source.value.includes(":") || path.posix.isAbsolute(source.value) || source.value.split("/").some(segment => segment === ".." || segment === "")) throw new CliError("authoring-asset-unsupported", "Local asset paths require --asset-base and a relative slash-separated path without parent segments.", 3);
      const captured = await readInput(source.value, path.resolve(options.cwd, options.assetBase), remaining, options.signal);
      remaining -= captured.bytes.byteLength;
      assetInputs.push({ ...captured.input, path: path.relative(options.cwd, path.resolve(options.cwd, options.assetBase, source.value)).split(path.sep).join("/") });
      asset.source = { type: "bytes", encoding: "base64", value: Buffer.from(captured.bytes).toString("base64") };
    }
    if (!record(asset.source) || asset.source.type !== "bytes" || asset.source.encoding !== "base64" || typeof asset.source.value !== "string") throw new CliError("authoring-asset-unsupported", "Authoring JSON supports only base64 byte assets; paths and URLs are not loaded.", 3);
  }
  let document;
  try {
    document = createPresentation((data.presentation ?? {}) as PresentationInit);
    for (const asset of assets) document.registerAsset(asset as unknown as PresentationAssetInput);
    for (const layout of (data.layouts ?? []) as SlideLayoutInput[]) document.defineSlideLayout(layout);
    for (const slide of data.slides as PresentationSlideInput[]) document.addSlide(slide);
  } catch (error) {
    if (error instanceof Error) throw new CliError("authoring-invalid", error.message, 3);
    throw error;
  }
  return { document, assetInputs };
}

export async function exportAuthoring(document: Awaited<ReturnType<typeof parseAuthoring>>["document"], signal?: AbortSignal, options: { imageDiagnostics?: boolean; validateImageData?: boolean } = {}) {
  const { imageDiagnostics, validateImageData } = options;
  try {
    return await generatePptx(document, {
      ...(signal ? { signal } : {}),
      ...(imageDiagnostics ? { jpegOrientationDiagnostics: "warn", jpegIccDiagnostics: "warn" } : {}),
      ...(validateImageData ? { pngDiagnostics: "reject-scanlines", gifDiagnostics: "reject-lzw", jpegValidation: "structure" }
        : imageDiagnostics ? { pngDiagnostics: "warn", gifDiagnostics: "warn" } : {}),
    });
  }
  catch (error) {
    if (signal?.aborted) throw new CliError("operation-aborted", "Command cancelled.", 130);
    if (error instanceof PresentationValidationError) throw new CliError("authoring-validation-failed", error.message, 3, { diagnostics: error.diagnostics });
    if (error instanceof PptxExportError) throw new CliError(error.code, error.message, 3, { diagnostics: error.diagnostics });
    if (error instanceof RangeError && ["Cannot serialize a character forbidden by XML 1.0.", "Overlapping OOXML escape-like sequences are not supported for custom properties."].includes(error.message)) throw new CliError("authoring-export-invalid", error.message, 3);
    throw error;
  }
}
