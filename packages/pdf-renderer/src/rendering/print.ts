import type { LayoutResult } from "@pptkit/layout";
import { renderLayoutToSvg, SvgRenderAbortError } from "@pptkit/svg-renderer";
import type { SvgRenderOptions, SvgRenderWarning } from "@pptkit/svg-renderer";
import { jpegDataUrlPrecision } from "./jpeg-prefix.js";
import { captureImageDataUrl, ImageInspectionLimitError } from "./image-data.js";
import { inspectStaticImage, validateJpegStructure } from "@pptkit/media-inspection";

/** Minimal Playwright-compatible interfaces; the caller owns and launches Chromium. */
export interface PdfBrowser {
  newContext(): Promise<PdfBrowserContext>;
}
export interface PdfBrowserContext {
  route(url: string, handler: (route: PdfRoute) => Promise<void>): Promise<unknown>;
  newPage(): Promise<PdfPage>;
  close(): Promise<void>;
}
export interface PdfRoute { abort(): Promise<void> }
export interface PdfPage {
  setContent(html: string): Promise<void>;
  evaluate(expression: string): Promise<unknown>;
  pdf(options: { preferCSSPageSize: boolean; printBackground: boolean }): Promise<Uint8Array>;
}

export interface PdfRenderOptions extends Omit<SvgRenderOptions, "portabilityDiagnostics" | "slideActionMode"> {
  browser: PdfBrowser;
  /** Require a supplied and browser-loadable embedded face for every emitted font request. */
  fontPolicy?: "allow-fallback" | "require-embedded";
  /** Require caller verification of each emitted raster image use. */
  imageAnimationPolicy?: "allow" | "require-static" | "require-structurally-static";
  isStaticImage?: (request: Readonly<PdfStaticImageRequest>) => boolean;
  /** Cumulative decoded bytes across unique image data URLs inspected for static structure. */
  maxStaticInspectionBytes?: number;
  /** Optionally reject declared JPEG resources that decode with transparent pixels. */
  jpegDecodePolicy?: "browser" | "require-opaque";
  /** Require an 8-bit JPEG frame header within the bounded source prefix. */
  jpegPrecisionPolicy?: "browser" | "require-8-bit";
  /** Require the supported complete 8-bit Huffman DCT JPEG structure profile. */
  jpegStructurePolicy?: "browser" | "require-valid-8-bit";
  /** Cumulative decoded bytes across all unique inspected image data URLs. */
  maxStructureInspectionBytes?: number;
  /** Maximum cumulative decoded pixels across unique JPEG data URLs. */
  maxJpegInspectionPixels?: number;
  /** Maximum PDF bytes, checked after Chromium has produced the file. Defaults to 64 MiB. */
  maxOutputBytes?: number;
}
export interface PdfStaticImageRequest { dataUrl: string; mimeType: string; slideId: string; slideIndex: number }
export interface PdfRenderResult {
  bytes: Uint8Array;
  pages: { slideId: string; index: number; hidden: boolean }[];
  warnings: SvgRenderWarning[];
}
export class PdfRenderError extends Error {
  constructor(readonly code: "pdf-resource-unavailable" | "pdf-font-unavailable" | "pdf-image-not-static" | "pdf-image-static-unverified" | "pdf-jpeg-unsupported-precision" | "pdf-jpeg-precision-unavailable" | "pdf-jpeg-structure-unavailable" | "pdf-empty-selection" | "pdf-limit-exceeded" | "pdf-print-failed", message: string) {
    super(message); this.name = "PdfRenderError";
  }
}

const unavailable = new Set(["asset-missing", "asset-unresolved", "asset-resolver-failed", "asset-path-unsupported", "asset-bytes-unsupported", "unsafe-asset-url", "svg-external-image-dependency", "svg-inline-svg-image-unverified"]);
class InvalidStaticImageResult extends TypeError {}
function checkAbort(signal: AbortSignal | undefined): void { if (signal?.aborted) throw new SvgRenderAbortError(); }
function wait<T>(operation: Promise<T>, signal: AbortSignal | undefined): Promise<T> {
  if (!signal) return operation;
  if (signal.aborted) { void operation.catch(() => {}); return Promise.reject(new SvgRenderAbortError()); }
  return new Promise<T>((resolve, reject) => {
    const abort = () => { signal.removeEventListener("abort", abort); reject(new SvgRenderAbortError()); };
    signal.addEventListener("abort", abort, { once: true });
    operation.then(value => { signal.removeEventListener("abort", abort); resolve(value); }, error => { signal.removeEventListener("abort", abort); reject(error); });
    if (signal.aborted) abort();
  });
}

/** Print selected SVG slides in a fresh, network-denied Chromium context. */
export async function renderLayoutToPdf(layout: LayoutResult, options: PdfRenderOptions): Promise<PdfRenderResult> {
  if (!options || typeof options.browser?.newContext !== "function") throw new TypeError("A Chromium-compatible browser is required.");
  const { browser, maxOutputBytes = 64 * 1024 * 1024, signal, fontPolicy = "allow-fallback", imageAnimationPolicy = "allow", isStaticImage, maxStaticInspectionBytes, jpegDecodePolicy = "browser", jpegPrecisionPolicy = "browser", jpegStructurePolicy = "browser", maxJpegInspectionPixels, maxStructureInspectionBytes, ...svgOptions } = options;
  if (fontPolicy !== "allow-fallback" && fontPolicy !== "require-embedded") throw new TypeError("Invalid fontPolicy.");
  if (imageAnimationPolicy !== "allow" && imageAnimationPolicy !== "require-static" && imageAnimationPolicy !== "require-structurally-static") throw new TypeError("Invalid imageAnimationPolicy.");
  if (imageAnimationPolicy === "require-static" ? typeof isStaticImage !== "function" : isStaticImage !== undefined) throw new TypeError("isStaticImage is required only with imageAnimationPolicy require-static.");
  const maxStaticBytes = maxStaticInspectionBytes === undefined ? 16 * 1024 * 1024 : maxStaticInspectionBytes;
  if (!Number.isSafeInteger(maxStaticBytes) || maxStaticBytes < 0 || maxStaticBytes > 64 * 1024 * 1024 || (imageAnimationPolicy !== "require-structurally-static" && maxStaticInspectionBytes !== undefined)) throw new RangeError("maxStaticInspectionBytes requires require-structurally-static and must be between zero and 67108864.");
  if (jpegDecodePolicy !== "browser" && jpegDecodePolicy !== "require-opaque") throw new TypeError("Invalid jpegDecodePolicy.");
  if (jpegPrecisionPolicy !== "browser" && jpegPrecisionPolicy !== "require-8-bit") throw new TypeError("Invalid jpegPrecisionPolicy.");
  if (jpegStructurePolicy !== "browser" && jpegStructurePolicy !== "require-valid-8-bit") throw new TypeError("Invalid jpegStructurePolicy.");
  const maxStructureBytes = maxStructureInspectionBytes === undefined ? 16 * 1024 * 1024 : maxStructureInspectionBytes;
  if (!Number.isSafeInteger(maxStructureBytes) || maxStructureBytes < 0 || maxStructureBytes > 64 * 1024 * 1024 || (jpegStructurePolicy === "browser" && maxStructureInspectionBytes !== undefined)) throw new RangeError("maxStructureInspectionBytes requires require-valid-8-bit and must be between zero and 67108864.");
  const maxJpegPixels = maxJpegInspectionPixels === undefined ? 16_000_000 : maxJpegInspectionPixels;
  if (!Number.isSafeInteger(maxJpegPixels) || maxJpegPixels < 0 || maxJpegPixels > 64_000_000 || (jpegDecodePolicy === "browser" && maxJpegInspectionPixels !== undefined)) throw new RangeError("maxJpegInspectionPixels requires require-opaque and must be between zero and 64000000.");
  let embeddedFonts: SvgRenderOptions["embeddedFonts"];
  if (svgOptions.embeddedFonts !== undefined) {
    if (!Array.isArray(svgOptions.embeddedFonts) || svgOptions.embeddedFonts.length > 32) throw new TypeError("embeddedFonts must contain at most 32 font faces.");
    embeddedFonts = svgOptions.embeddedFonts.map(face => {
      if (!face || typeof face !== "object") throw new TypeError("Each embedded font must be a font face object.");
      return { family: face.family, weight: face.weight, italic: face.italic, format: face.format, base64: face.base64 };
    });
  }
  if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes < 0) throw new RangeError("maxOutputBytes must be a nonnegative safe integer.");
  if (![layout.size.width, layout.size.height].every(value => typeof value === "number" && Number.isFinite(value) && value >= 1 && value <= 14400)) throw new RangeError("PDF page dimensions must be between 1 and 14400 points.");
  checkAbort(signal);
  const rendered = await renderLayoutToSvg(layout, { ...svgOptions, ...(embeddedFonts === undefined ? {} : { embeddedFonts }), maxSvgBytes: svgOptions.maxSvgBytes === undefined ? 64 * 1024 * 1024 : svgOptions.maxSvgBytes, ...(signal === undefined ? {} : { signal }), portabilityDiagnostics: "warn", slideActionMode: "document-links" });
  checkAbort(signal);
  if (rendered.slides.length === 0) throw new PdfRenderError("pdf-empty-selection", "PDF output requires at least one selected slide.");
  const problem = rendered.warnings.find(warning => unavailable.has(warning.code));
  if (problem) throw new PdfRenderError("pdf-resource-unavailable", `PDF output requires resolved inline resources (${problem.code}).`);
  const needed = new Map<string, { family: string; weight: 400 | 700; italic: boolean }>();
  if (fontPolicy === "require-embedded") {
    const supplied = new Set((embeddedFonts ?? []).map(face => JSON.stringify([face.family.toLowerCase(), face.weight, face.italic])));
    for (const warning of rendered.warnings) for (const font of warning.fontDependencies ?? []) {
      const weight = font.bold ? 700 : 400, key = JSON.stringify([font.family.toLowerCase(), weight, font.italic]);
      if (!supplied.has(key)) throw new PdfRenderError("pdf-font-unavailable", "A requested font family and style has no matching embedded face.");
      needed.set(key, { family: font.family, weight, italic: font.italic });
    }
  }
  const css = `@page{size:${rendered.width}pt ${rendered.height}pt;margin:0}html,body{margin:0;padding:0}section{width:${rendered.width}pt;height:${rendered.height}pt;break-after:page}svg{display:block;width:100%;height:100%}`;
  const html = `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'"><style>${css}</style>${rendered.slides.map(slide => `<section>${slide.svg}</section>`).join("")}`;
  // Do not race context creation: if it eventually succeeds we must close it.
  let context: PdfBrowserContext | undefined, failed = false, blockedRequests = 0;
  try {
    context = await browser.newContext();
    checkAbort(signal);
    await wait(context.route("**/*", async route => { blockedRequests++; await route.abort(); }), signal);
    const page = await wait(context.newPage(), signal);
    await wait(page.setContent(html), signal);
    await wait(page.evaluate("document.fonts.ready.then(() => true)"), signal);
    let attested = 0, structurallyStatic = 0, inventoried = 0;
    if (imageAnimationPolicy !== "allow" || jpegPrecisionPolicy === "require-8-bit" || jpegStructurePolicy === "require-valid-8-bit") {
      const uses = await wait(page.evaluate("Array.from(document.querySelectorAll('section'), section => Array.from(section.querySelectorAll('svg image'), node => node.getAttribute('href')))"), signal);
      if (!Array.isArray(uses) || uses.length !== rendered.slides.length || uses.some(item => !Array.isArray(item))) throw new PdfRenderError("pdf-resource-unavailable", "Browser image inventory was unavailable.");
      const checked = new Set<string>();
      const structurallyChecked = new Set<string>(); let structureBytes = 0, staticBytes = 0;
      for (let index = 0; index < uses.length; index++) for (const href of uses[index] as unknown[]) {
        checkAbort(signal);
        if (typeof href !== "string" || !/^data:image\//i.test(href) || /^data:image\/svg\+xml(?:;|,)/i.test(href)) throw new PdfRenderError("pdf-resource-unavailable", "PDF output requires inline, non-SVG image resources.");
        inventoried++;
        if (jpegPrecisionPolicy === "require-8-bit" && !checked.has(href)) {
          const precision = jpegDataUrlPrecision(href); checked.add(href);
          if (precision === "unavailable") throw new PdfRenderError("pdf-jpeg-precision-unavailable", "The bounded image prefix could not establish JPEG precision.");
          if (typeof precision === "number" && precision !== 8) throw new PdfRenderError("pdf-jpeg-unsupported-precision", "A JPEG frame does not use 8-bit sample precision.");
        }
        if ((jpegStructurePolicy === "require-valid-8-bit" || imageAnimationPolicy === "require-structurally-static") && !structurallyChecked.has(href)) {
          let bytes: Uint8Array;
          const remaining = Math.min(jpegStructurePolicy === "require-valid-8-bit" ? maxStructureBytes - structureBytes : Infinity, imageAnimationPolicy === "require-structurally-static" ? maxStaticBytes - staticBytes : Infinity);
          try { bytes = captureImageDataUrl(href, remaining); }
          catch (error) {
            if (error instanceof ImageInspectionLimitError) throw new PdfRenderError("pdf-limit-exceeded", "Image structure inspection byte budget exceeded.");
            throw new PdfRenderError(imageAnimationPolicy === "require-structurally-static" ? "pdf-image-static-unverified" : "pdf-jpeg-structure-unavailable", "An image data URL could not be structurally inspected.");
          }
          structurallyChecked.add(href);
          if (jpegStructurePolicy === "require-valid-8-bit") structureBytes += bytes.byteLength;
          if (imageAnimationPolicy === "require-structurally-static") staticBytes += bytes.byteLength;
          const jpeg = bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xd8;
          const declared = /^data:image\/(?:jpe?g|pjpeg)(?:;|,)/i.test(href);
          if (!jpeg && declared && jpegStructurePolicy === "require-valid-8-bit") throw new PdfRenderError("pdf-jpeg-structure-unavailable", "A declared JPEG does not have a JPEG signature.");
          if (jpeg && jpegStructurePolicy === "require-valid-8-bit") {
            let precision: number;
            try { precision = validateJpegStructure(bytes, signal).precision; }
            catch (error) { if (signal?.aborted) throw new SvgRenderAbortError(); throw new PdfRenderError("pdf-jpeg-structure-unavailable", "JPEG structure is invalid or outside the supported profile."); }
            if (precision !== 8) throw new PdfRenderError("pdf-jpeg-unsupported-precision", "A JPEG frame does not use 8-bit sample precision.");
          }
          if (imageAnimationPolicy === "require-structurally-static") {
            let inspection;
            try { inspection = inspectStaticImage(bytes, signal === undefined ? {} : { signal }); }
            catch (error) { if (signal?.aborted) throw new SvgRenderAbortError(); throw new PdfRenderError("pdf-image-static-unverified", "Image structure is invalid or outside the static profile."); }
            if (inspection.status === "animation-declared") throw new PdfRenderError("pdf-image-not-static", "The image declares animation.");
            if (inspection.status !== "static") throw new PdfRenderError("pdf-image-static-unverified", "Image format or structure is outside the static profile.");
            structurallyStatic++;
          }
        }
        if (imageAnimationPolicy !== "require-static") continue;
        const slide = rendered.slides[index]!;
        let accepted: unknown;
        try { accepted = isStaticImage!(Object.freeze({ dataUrl: href, mimeType: /^data:([^;,]+)/i.exec(href)?.[1]?.toLowerCase() ?? "", slideId: slide.slideId, slideIndex: slide.index })); }
        catch (error) { if (signal?.aborted || error instanceof SvgRenderAbortError) throw new SvgRenderAbortError(); throw new PdfRenderError("pdf-image-not-static", "The image static-frame callback failed."); }
        if (typeof accepted !== "boolean") { void Promise.resolve(accepted).catch(() => {}); checkAbort(signal); throw new InvalidStaticImageResult("isStaticImage must return a synchronous boolean."); }
        checkAbort(signal);
        if (!accepted) throw new PdfRenderError("pdf-image-not-static", "The image static-frame callback rejected a resource.");
        attested++;
      }
    }
    if (needed.size) {
      let loaded: unknown;
      try {
        loaded = await wait(page.evaluate(`(async () => { const faces = ${JSON.stringify([...needed.values()])}; for (const face of faces) { const css = (face.italic ? 'italic' : 'normal') + ' ' + face.weight + ' 16px ' + JSON.stringify(face.family); const matches = await document.fonts.load(css); if (!matches.length || matches.some(item => item.status !== 'loaded')) return false; } return true; })()`), signal);
      } catch (error) {
        if (error instanceof SvgRenderAbortError) throw error;
        throw new PdfRenderError("pdf-font-unavailable", "The browser could not load a required embedded font face.");
      }
      if (loaded !== true) throw new PdfRenderError("pdf-font-unavailable", "The browser could not load a required embedded font face.");
    }
    const imageCount = await wait(page.evaluate(`(async () => { const nodes = document.querySelectorAll('svg image'); for (const node of nodes) { const href = node.getAttribute('href'); if (!href || !/^data:image\\//i.test(href) || /^data:image\\/svg\\+xml(?:;|,)/i.test(href)) return -1; const image = new Image(); image.src = href; try { await image.decode(); } catch { return -1; } if (!image.naturalWidth || !image.naturalHeight) return -1; } return nodes.length; })()`), signal);
    if (typeof imageCount !== "number" || !Number.isSafeInteger(imageCount) || imageCount < 0 || blockedRequests || ((imageAnimationPolicy !== "allow" || jpegPrecisionPolicy === "require-8-bit" || jpegStructurePolicy === "require-valid-8-bit") && imageCount !== inventoried)) throw new PdfRenderError("pdf-resource-unavailable", "An image resource could not be decoded, changed during inspection, or requested an external resource.");
    if (jpegDecodePolicy === "require-opaque") {
      const result = await wait(page.evaluate(String.raw`(async () => {
        const seen = new Set(); let used = 0;
        for (const node of document.querySelectorAll('svg image')) {
          const href = node.getAttribute('href');
          if (!href || !/^data:image\/(?:jpe?g|pjpeg)(?:;|,)/i.test(href) || seen.has(href)) continue;
          seen.add(href); const image = new Image(); image.src = href;
          try { await image.decode(); } catch { return 'unavailable'; }
          const width = image.naturalWidth, height = image.naturalHeight, pixels = width * height;
          if (!width || !height || !Number.isSafeInteger(pixels) || width > 16384 || height > 16384 || pixels > ${maxJpegPixels} - used) return 'limit';
          const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height;
          try {
            const context = canvas.getContext('2d'); if (!context) return 'unavailable';
            context.drawImage(image, 0, 0, width, height);
            for (let y = 0; y < height; y += 64) {
              const data = context.getImageData(0, y, width, Math.min(64, height - y)).data;
              for (let i = 3; i < data.length; i += 4) if (data[i] !== 255) return 'unavailable';
            }
          } catch { return 'unavailable'; } finally { canvas.width = 0; canvas.height = 0; }
          used += pixels;
        }
        return 'ok';
      })()`), signal);
      if (result === "limit") throw new PdfRenderError("pdf-limit-exceeded", "Decoded JPEG inspection exceeds its pixel budget.");
      if (result !== "ok") throw new PdfRenderError("pdf-resource-unavailable", "A JPEG resource decoded with transparent or unreadable pixels.");
    }
    const bytes = await wait(page.pdf({ preferCSSPageSize: true, printBackground: true }), signal);
    checkAbort(signal);
    if (blockedRequests) throw new PdfRenderError("pdf-resource-unavailable", "Printing requested an external resource.");
    if (!(bytes instanceof Uint8Array) || bytes.length < 5 || String.fromCharCode(...bytes.subarray(0, 5)) !== "%PDF-") throw new PdfRenderError("pdf-print-failed", "Chromium did not return PDF bytes.");
    if (bytes.byteLength > maxOutputBytes) throw new PdfRenderError("pdf-limit-exceeded", "PDF output byte budget exceeded.");
    return { bytes: new Uint8Array(bytes), pages: rendered.slides.map(({ slideId, index, hidden }) => ({ slideId, index, hidden })), warnings: [...rendered.warnings, ...(imageCount ? [{ code: "pdf-image-frame-unverified", message: "Image decoding, animation frame selection and painted pixels are not fully verified." }] : []), ...(attested ? [{ code: "pdf-image-static-attested", message: "The caller attested that each emitted image is static; the assertion was not independently verified." }] : []), ...(structurallyStatic ? [{ code: "pdf-image-static-structure-checked", message: "Image structures matched the bounded static profile; decoding and painted pixels remain unverified." }] : []), { code: "pdf-font-selection-unverified", message: "Browser font selection, glyph coverage and fallback are not verified." }, { code: "pdf-page-size-rounding", message: "Chromium may round requested page dimensions." }] };
  } catch (error) {
    failed = true;
    if (signal?.aborted) throw new SvgRenderAbortError();
    if (error instanceof SvgRenderAbortError || error instanceof PdfRenderError || error instanceof InvalidStaticImageResult) throw error;
    throw new PdfRenderError("pdf-print-failed", "The browser could not print the document.");
  } finally {
    if (context) try { await context.close(); } catch { if (!failed) throw new PdfRenderError("pdf-print-failed", "The browser context could not be closed."); }
    if (!failed) checkAbort(signal);
  }
}
