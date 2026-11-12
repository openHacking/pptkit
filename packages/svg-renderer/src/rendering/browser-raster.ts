import type { LayoutResult } from "@pptkit/layout";
import { checkRenderAbort, waitForResolver } from "./abort.js";
import { renderResolvedLayout } from "./render.js";
import type { SvgRenderOptions, SvgRenderWarning } from "../types/public.js";
import { SvgRenderAbortError } from "../types/public.js";
import { captureEmbeddedFonts, embeddedFontMediaType } from "./embedded-fonts.js";

export interface BrowserRasterOptions extends Omit<SvgRenderOptions, "portabilityDiagnostics"> {
  format: "png" | "jpeg";
  scale?: number;
  /** JPEG compositing color, six hexadecimal digits prefixed by #. */
  jpegBackground?: string;
  jpegQuality?: number;
  /** Cumulative pixels across selected slides; hard ceiling is 64 million. */
  maxTotalPixels?: number;
  /** Cumulative encoded bytes; defaults to 64 MiB. */
  maxOutputBytes?: number;
  /** Optional crop in slide coordinates; it must be contained in the slide. */
  crop?: { x: number; y: number; width: number; height: number };
  /** Require matching supplied face declarations that the browser can load. */
  fontPolicy?: "allow-fallback" | "require-embedded" | "require-covered";
  /** Synchronous, caller-owned proof of coverage for each emitted text fragment. */
  fontCoverage?: (request: Readonly<BrowserRasterFontCoverageRequest>) => boolean;
  /** Require caller verification that every embedded raster image is static. */
  imageAnimationPolicy?: "allow" | "require-static";
  isStaticImage?: (request: Readonly<BrowserRasterStaticImageRequest>) => boolean;
  /** Optionally reject JPEG resources whose decoded pixels are not fully opaque. */
  jpegDecodePolicy?: "browser" | "require-opaque";
  /** Cumulative decoded JPEG pixels checked by require-opaque, at most 64 million. */
  maxJpegInspectionPixels?: number;
}

export interface BrowserRasterStaticImageRequest {
  dataUrl: string;
  mimeType: string;
  slideId: string;
  slideIndex: number;
}

export interface BrowserRasterFontCoverageRequest {
  family: string;
  bold: boolean;
  italic: boolean;
  text: string;
  fontSize: number;
  language?: string;
  slideId: string;
  elementId?: string;
}

export interface BrowserRasterSlide {
  slideId: string;
  index: number;
  hidden: boolean;
  width: number;
  height: number;
  mimeType: "image/png" | "image/jpeg";
  bytes: Uint8Array;
}
export interface BrowserRasterResult { slides: BrowserRasterSlide[]; warnings: SvgRenderWarning[] }

export class BrowserRasterError extends Error {
  constructor(readonly code: "raster-environment-unavailable" | "raster-limit-exceeded" | "raster-resource-unavailable" | "raster-font-unavailable" | "raster-image-not-static" | "raster-decode-failed" | "raster-encoding-failed", message: string, readonly slideIndex?: number, readonly elementId?: string) {
    super(message); this.name = "BrowserRasterError";
  }
}

const unavailable = new Set(["asset-missing", "asset-unresolved", "asset-resolver-failed", "asset-path-unsupported", "asset-bytes-unsupported", "unsafe-asset-url", "svg-external-image-dependency", "svg-inline-svg-image-unverified"]);

// Keep browser globals out of public types and the package's Node-oriented TS libs.
interface RasterImage { src: string; naturalWidth: number; naturalHeight: number; decode(): Promise<void> }
interface ParsedSvg {
  getElementsByTagNameNS(namespace: string, name: string): { readonly length: number; item(index: number): { getAttribute(name: string): string | null } | null };
}
interface RasterCanvas {
  width: number; height: number;
  getContext(kind: "2d"): { fillStyle: string; fillRect(x: number, y: number, width: number, height: number): void; drawImage(image: RasterImage, sx: number, sy: number, sw: number, sh: number, dx: number, dy: number, dw: number, dh: number): void; getImageData(x: number, y: number, width: number, height: number): { data: Uint8ClampedArray } } | null;
  toBlob(callback: (blob: Blob | null) => void, type: string, quality?: number): void;
}

/** Browser-only, atomic multi-slide conversion; returned bytes are not published until all slides succeed. */
export async function renderLayoutToBrowserRaster(layout: LayoutResult, options: BrowserRasterOptions): Promise<BrowserRasterResult> {
  if (!options || (options.format !== "png" && options.format !== "jpeg")) throw new TypeError("Raster format must be png or jpeg.");
  const { format, signal } = options;
  checkRenderAbort(signal);
  const fontPolicy = options.fontPolicy === undefined ? "allow-fallback" : options.fontPolicy;
  if (fontPolicy !== "allow-fallback" && fontPolicy !== "require-embedded" && fontPolicy !== "require-covered") throw new TypeError("Invalid fontPolicy.");
  const coverage = options.fontCoverage;
  if (fontPolicy === "require-covered" ? typeof coverage !== "function" : coverage !== undefined) throw new TypeError("fontCoverage is required only with fontPolicy require-covered.");
  const imageAnimationPolicy = options.imageAnimationPolicy === undefined ? "allow" : options.imageAnimationPolicy;
  if (imageAnimationPolicy !== "allow" && imageAnimationPolicy !== "require-static") throw new TypeError("Invalid imageAnimationPolicy.");
  const isStaticImage = options.isStaticImage;
  if (imageAnimationPolicy === "require-static" ? typeof isStaticImage !== "function" : isStaticImage !== undefined) throw new TypeError("isStaticImage is required only with imageAnimationPolicy require-static.");
  const jpegDecodePolicy = options.jpegDecodePolicy === undefined ? "browser" : options.jpegDecodePolicy;
  if (jpegDecodePolicy !== "browser" && jpegDecodePolicy !== "require-opaque") throw new TypeError("Invalid jpegDecodePolicy.");
  const maxJpegPixels = options.maxJpegInspectionPixels === undefined ? 16_000_000 : options.maxJpegInspectionPixels;
  if (!Number.isSafeInteger(maxJpegPixels) || maxJpegPixels < 0 || maxJpegPixels > 64_000_000 || (jpegDecodePolicy === "browser" && options.maxJpegInspectionPixels !== undefined)) throw new RangeError("maxJpegInspectionPixels requires require-opaque and must be between zero and 64000000.");
  const embeddedFonts = captureEmbeddedFonts(options.embeddedFonts).faces;
  const supplied = new Map(embeddedFonts.map(face => [JSON.stringify([face.family.toLowerCase(), face.weight, face.italic]), face]));
  const scale = options.scale === undefined ? 1 : options.scale;
  if (typeof scale !== "number" || !Number.isFinite(scale) || scale <= 0) throw new RangeError("Raster scale must be positive and finite.");
  const jpegBackground = options.jpegBackground === undefined ? "#FFFFFF" : options.jpegBackground;
  if (typeof jpegBackground !== "string" || !/^#[0-9a-f]{6}$/i.test(jpegBackground)) throw new TypeError("jpegBackground must be a six-digit hexadecimal color.");
  const quality = options.jpegQuality === undefined ? 0.92 : options.jpegQuality;
  if (typeof quality !== "number" || !Number.isFinite(quality) || quality < 0 || quality > 1 || (format === "png" && (options.jpegQuality !== undefined || options.jpegBackground !== undefined))) throw new TypeError("JPEG options require JPEG format and quality between zero and one.");
  const maxPixels = options.maxTotalPixels === undefined ? 64_000_000 : options.maxTotalPixels;
  if (!Number.isSafeInteger(maxPixels) || maxPixels < 0 || maxPixels > 64_000_000) throw new RangeError("maxTotalPixels must be a safe integer from zero to 64000000.");
  const maxBytes = options.maxOutputBytes === undefined ? 64 * 1024 * 1024 : options.maxOutputBytes;
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 0) throw new RangeError("maxOutputBytes must be a nonnegative safe integer.");
  if (!Number.isFinite(layout.size.width) || !Number.isFinite(layout.size.height) || layout.size.width <= 0 || layout.size.height <= 0) throw new RangeError("Slide dimensions must be positive and finite.");
  const crop = options.crop === undefined ? { x: 0, y: 0, width: layout.size.width, height: layout.size.height } : { ...options.crop };
  if (![crop.x, crop.y, crop.width, crop.height].every(value => typeof value === "number" && Number.isFinite(value)) || crop.x < 0 || crop.y < 0 || crop.width <= 0 || crop.height <= 0 || crop.x + crop.width > layout.size.width || crop.y + crop.height > layout.size.height) throw new RangeError("Raster crop must be a positive rectangle contained in the slide.");
  const host = globalThis as unknown as { document?: { createElement(name: "canvas"): RasterCanvas }; Image?: new () => RasterImage; DOMParser?: new () => { parseFromString(source: string, type: "image/svg+xml"): ParsedSvg }; FontFace?: new (family: string, source: string, descriptors: { weight: string; style: string }) => { load(): Promise<unknown> } };
  const browserDocument = host.document, ImageConstructor = host.Image, Parser = host.DOMParser;
  if (!browserDocument || !ImageConstructor || !Parser || typeof URL.createObjectURL !== "function") throw new BrowserRasterError("raster-environment-unavailable", "Browser image, SVG parser and canvas APIs are required.");
  const slidePositions = fontPolicy === "require-covered" ? new Map(layout.slides.map((slide, index) => [slide.id, index])) : undefined;
  const verifyCoverage = fontPolicy !== "require-covered" ? undefined : (request: BrowserRasterFontCoverageRequest): void => {
    checkRenderAbort(signal);
    const slideIndex = slidePositions!.get(request.slideId);
    const key = JSON.stringify([request.family.toLowerCase(), request.bold ? 700 : 400, request.italic]);
    if (!supplied.has(key)) throw new BrowserRasterError("raster-font-unavailable", "A requested font family and style has no matching embedded face.", slideIndex, request.elementId);
    let covered: unknown;
    try { covered = coverage!(Object.freeze({ ...request })); }
    catch (error) {
      if (signal?.aborted || error instanceof SvgRenderAbortError) throw new SvgRenderAbortError();
      throw new BrowserRasterError("raster-font-unavailable", "The font coverage callback failed.", slideIndex, request.elementId);
    }
    if (typeof covered !== "boolean") {
      void Promise.resolve(covered).catch(() => {});
      checkRenderAbort(signal);
      throw new TypeError("fontCoverage must return a synchronous boolean.");
    }
    checkRenderAbort(signal);
    if (!covered) throw new BrowserRasterError("raster-font-unavailable", "A font coverage callback rejected an emitted text fragment.", slideIndex, request.elementId);
  };
  const rendered = await renderResolvedLayout(layout, { ...options, embeddedFonts, portabilityDiagnostics: "warn" }, verifyCoverage);
  checkRenderAbort(signal);
  const problematic = rendered.warnings.find(warning => unavailable.has(warning.code));
  if (problematic) throw new BrowserRasterError("raster-resource-unavailable", `Raster output requires resolved, non-external image resources (${problematic.code}).`, rendered.slides.find(slide => slide.slideId === problematic.slideId)?.index);
  if (rendered.slides.length === 0) return { slides: [], warnings: rendered.warnings };
  const width = Math.round(crop.width * scale), height = Math.round(crop.height * scale);
  if (![width, height].every(value => Number.isSafeInteger(value) && value > 0 && value <= 16384)) throw new BrowserRasterError("raster-limit-exceeded", "Raster dimensions must be positive and no greater than 16384 pixels.");
  const perSlide = width * height, total = perSlide * rendered.slides.length;
  if (!Number.isSafeInteger(total) || perSlide > 16_000_000 || total > maxPixels) throw new BrowserRasterError("raster-limit-exceeded", "Raster pixel budget exceeded.");
  if (fontPolicy === "require-embedded" || fontPolicy === "require-covered") {
    const needed = new Map<string, { slideIndex: number | undefined; elementId: string | undefined }>();
    for (const warning of rendered.warnings) for (const font of warning.fontDependencies ?? []) {
      const key = JSON.stringify([font.family.toLowerCase(), font.bold ? 700 : 400, font.italic]);
      if (!needed.has(key)) needed.set(key, { slideIndex: rendered.slides.find(slide => slide.slideId === warning.slideId)?.index, elementId: warning.elementId });
    }
    for (const [key, location] of needed) {
      checkRenderAbort(signal);
      const face = supplied.get(key);
      if (!face) throw new BrowserRasterError("raster-font-unavailable", "A requested font family and style has no matching embedded face.", location.slideIndex, location.elementId);
      if (!host.FontFace) throw new BrowserRasterError("raster-environment-unavailable", "Browser FontFace loading is required by fontPolicy.", location.slideIndex, location.elementId);
      try {
        const resource = `url(data:${embeddedFontMediaType(face.format)};base64,${face.base64})`;
        const font = new host.FontFace(face.family, resource, { weight: String(face.weight), style: face.italic ? "italic" : "normal" });
        await waitForResolver(font.load(), signal);
      } catch (error) {
        if (error instanceof SvgRenderAbortError) throw error;
        throw new BrowserRasterError("raster-font-unavailable", "The browser could not load a required embedded font face.", location.slideIndex, location.elementId);
      }
    }
  }
  const mimeType = format === "png" ? "image/png" : "image/jpeg";
  const slides: BrowserRasterSlide[] = [];
  const decodedAssets = new Set<string>();
  let used = 0, inspectedJpegPixels = 0;
  for (const slide of rendered.slides) {
    checkRenderAbort(signal);
    const parsed = new Parser().parseFromString(slide.svg, "image/svg+xml");
    if (parsed.getElementsByTagNameNS("*", "parsererror").length > 0) throw new BrowserRasterError("raster-decode-failed", "The rendered SVG is not valid XML.", slide.index);
    const images = parsed.getElementsByTagNameNS("http://www.w3.org/2000/svg", "image");
    if (images.length > 0) rendered.warnings.push(imageAnimationPolicy === "require-static"
      ? { code: "raster-image-static-attested", message: "The caller attested that each embedded image is static; this assertion was not independently verified.", slideId: slide.slideId }
      : { code: "raster-image-frame-unverified", message: "Embedded image animation, if present, is sampled by the browser; frame timing and playback are not controlled.", slideId: slide.slideId });
    for (let i = 0; i < images.length; i++) {
      const href = images.item(i)?.getAttribute("href");
      if (!href || !/^data:image\//i.test(href) || /^data:image\/svg\+xml(?:;|,)/i.test(href)) throw new BrowserRasterError("raster-resource-unavailable", "Raster output requires inline, non-SVG image resources.", slide.index);
      if (imageAnimationPolicy === "require-static") {
        checkRenderAbort(signal);
        const mimeType = /^data:([^;,]+)/i.exec(href)?.[1]?.toLowerCase() ?? "";
        let isStatic: unknown;
        try { isStatic = isStaticImage!(Object.freeze({ dataUrl: href, mimeType, slideId: slide.slideId, slideIndex: slide.index })); }
        catch (error) {
          if (signal?.aborted || error instanceof SvgRenderAbortError) throw new SvgRenderAbortError();
          throw new BrowserRasterError("raster-image-not-static", "The image static-frame callback failed.", slide.index);
        }
        if (typeof isStatic !== "boolean") {
          void Promise.resolve(isStatic).catch(() => {});
          checkRenderAbort(signal);
          throw new TypeError("isStaticImage must return a synchronous boolean.");
        }
        checkRenderAbort(signal);
        if (!isStatic) throw new BrowserRasterError("raster-image-not-static", "The image static-frame callback rejected a resource.", slide.index);
      }
      if (decodedAssets.has(href)) continue;
      const asset = new ImageConstructor();
      try {
        asset.src = href;
        try { await waitForResolver(asset.decode(), signal); }
        catch (error) { if (error instanceof SvgRenderAbortError) throw error; throw new BrowserRasterError("raster-resource-unavailable", "An embedded image resource could not be decoded.", slide.index); }
        if (!asset.naturalWidth || !asset.naturalHeight) throw new BrowserRasterError("raster-resource-unavailable", "An embedded image resource has no decodable dimensions.", slide.index);
        if (jpegDecodePolicy === "require-opaque" && /^data:image\/(?:jpe?g|pjpeg)(?:;|,)/i.test(href)) {
          const pixels = asset.naturalWidth * asset.naturalHeight;
          if (!Number.isSafeInteger(pixels) || asset.naturalWidth > 16384 || asset.naturalHeight > 16384 || pixels > maxJpegPixels - inspectedJpegPixels) throw new BrowserRasterError("raster-limit-exceeded", "Decoded JPEG inspection exceeds its pixel budget.", slide.index);
          const probe = browserDocument.createElement("canvas"); probe.width = asset.naturalWidth; probe.height = asset.naturalHeight;
          const probeContext = probe.getContext("2d");
          if (!probeContext) throw new BrowserRasterError("raster-environment-unavailable", "A 2D canvas is unavailable for JPEG inspection.", slide.index);
          try {
            probeContext.drawImage(asset, 0, 0, asset.naturalWidth, asset.naturalHeight, 0, 0, asset.naturalWidth, asset.naturalHeight);
            for (let y = 0; y < asset.naturalHeight; y += 64) {
              checkRenderAbort(signal);
              const data = probeContext.getImageData(0, y, asset.naturalWidth, Math.min(64, asset.naturalHeight - y)).data;
              checkRenderAbort(signal);
              for (let alpha = 3; alpha < data.length; alpha += 4) if (data[alpha] !== 255) throw new BrowserRasterError("raster-resource-unavailable", "A JPEG resource decoded with transparent pixels.", slide.index);
              checkRenderAbort(signal);
            }
          } catch (error) {
            if (error instanceof BrowserRasterError || error instanceof SvgRenderAbortError) throw error;
            throw new BrowserRasterError("raster-resource-unavailable", "Decoded JPEG pixels could not be inspected.", slide.index);
          } finally { probe.width = 0; probe.height = 0; }
          inspectedJpegPixels += pixels;
        }
        decodedAssets.add(href);
      } finally { asset.src = ""; }
    }
    const image = new ImageConstructor();
    const url = URL.createObjectURL(new Blob([slide.svg], { type: "image/svg+xml" }));
    try {
      image.src = url;
      try { await waitForResolver(image.decode(), signal); }
      catch (error) { if (error instanceof SvgRenderAbortError) throw error; throw new BrowserRasterError("raster-decode-failed", "The browser could not decode a rendered SVG.", slide.index); }
      if (!image.naturalWidth || !image.naturalHeight) throw new BrowserRasterError("raster-decode-failed", "The rendered SVG has no decodable dimensions.", slide.index);
      checkRenderAbort(signal);
      const canvas = browserDocument.createElement("canvas"); canvas.width = width; canvas.height = height;
      const context = canvas.getContext("2d");
      if (!context) throw new BrowserRasterError("raster-environment-unavailable", "A 2D canvas is unavailable.", slide.index);
      if (format === "jpeg") { context.fillStyle = jpegBackground; context.fillRect(0, 0, width, height); }
      try { context.drawImage(image, crop.x / rendered.width * image.naturalWidth, crop.y / rendered.height * image.naturalHeight, crop.width / rendered.width * image.naturalWidth, crop.height / rendered.height * image.naturalHeight, 0, 0, width, height); }
      catch { throw new BrowserRasterError("raster-decode-failed", "The browser could not draw the decoded SVG.", slide.index); }
      let blob: Blob | null;
      try { blob = await waitForResolver(new Promise<Blob | null>((resolve, reject) => { try { canvas.toBlob(resolve, mimeType, format === "jpeg" ? quality : undefined); } catch (error) { reject(error); } }), signal); }
      catch (error) { if (error instanceof SvgRenderAbortError) throw error; throw new BrowserRasterError("raster-encoding-failed", "The browser could not encode the raster image.", slide.index); }
      if (!blob || blob.type !== mimeType) throw new BrowserRasterError("raster-encoding-failed", "The browser did not produce the requested image format.", slide.index);
      if (blob.size > maxBytes - used) throw new BrowserRasterError("raster-limit-exceeded", "Raster output byte budget exceeded.", slide.index);
      const bytes = new Uint8Array(await waitForResolver(blob.arrayBuffer(), signal));
      used += bytes.byteLength;
      slides.push({ slideId: slide.slideId, index: slide.index, hidden: slide.hidden, width, height, mimeType, bytes });
    } finally { image.src = ""; URL.revokeObjectURL(url); }
  }
  checkRenderAbort(signal);
  return { slides, warnings: rendered.warnings };
}
