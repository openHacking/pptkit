import type { NativePlaceholderInheritance, NativeInheritedPlaceholderDefaults, NativeInheritedRuns } from "../types/native-placeholder-inheritance.js";
import { inspectJpegExifOrientation } from "../assets/inspect-jpeg-exif.js";
import { inspectJpegIcc } from "../assets/inspect-jpeg-icc.js";
import { inspectGif, inspectGifImageData } from "../assets/inspect-gif.js";
import { validateJpegStructure } from "@pptkit/media-inspection";
import { validatePngScanlines } from "../assets/validate-png-scanlines.js";
import { inspectApng } from "../assets/inspect-apng.js";
import { inspectCapturedPngIcc } from "../assets/inspect-png-icc.js";
import { nativeDrawingContext } from "../ooxml/native-connectors.js";
import { placeholderIndices } from "../ooxml/placeholder-indices.js";
import type { ElementAction, NormalizedPlaceholderDefinition, NormalizedPresentationTheme } from "@pptkit/core";
import type { LayoutChartElement, LayoutElement, LayoutResult } from "@pptkit/layout";
import { encodeUtf8 } from "../binary/encode.js";
import { REL } from "../constants/ooxml.js";
import { elementXml, type ElementXmlContext } from "../ooxml/elements.js";
import {
  appPropertiesXml,
  contentTypesXml,
  corePropertiesXml,
  customPropertiesXml,
  notesMasterXml,
  notesSlideXml,
  presentationPropertiesXml,
  presentationXml,
  slideLayoutXml,
  slideMasterXml,
  slideXml,
  tableStylesXml,
  themeXml,
  viewPropertiesXml,
} from "../ooxml/package-parts.js";
import { chartPartXml } from "../ooxml/chart.js";
import { chartWorkbook } from "../ooxml/chart-workbook.js";
import { bubbleSeriesParts } from "../ooxml/bubble-series.js";
import { scatterWorkbook } from "../ooxml/scatter-workbook.js";
import { relationshipsXml, rootRelationshipsXml } from "../ooxml/relationships.js";
import type { ExportWarning } from "../types/export.js";
import type { AssetLoadContext, AssetLoader, PackagedMedia, Relationship, ZipPart } from "../types/internal.js";
import type { WaveAudioResource } from "../generation/wave-audio.js";
import type { Mp4VideoResource } from "../generation/mp4-video.js";
import { decodeAssetBytes } from "../assets/load-bytes.js";

interface PartContextResult {
  context: ElementXmlContext;
  relationships: Relationship[];
  chartParts: { id: number; chart: LayoutChartElement }[];
  audioTargets: { objectId: number; durationMs: number }[];
  videoTargets: { objectId: number; durationMs: number }[];
  mediaTargets: { kind: "audio" | "video"; objectId: number; durationMs: number }[];
}

function mediaFilename(value: string): string {
  return value.slice(value.lastIndexOf("/") + 1);
}

function sameBytes(a: Uint8Array | undefined, b: Uint8Array | undefined): boolean {
  if (!a || !b) return a === b;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function byteFingerprint(bytes: Uint8Array | undefined): string {
  if (!bytes) return "absent";
  let hash = 2166136261;
  for (const value of bytes) hash = Math.imul(hash ^ value, 16777619);
  return `${bytes.length}:${hash >>> 0}`;
}

function createPartContext(options: {
  theme: NormalizedPresentationTheme;
  textPath?: string;
  baseRelationships: Relationship[];
  media: ReadonlyMap<string, PackagedMedia>;
  audio?: ReadonlyMap<string, WaveAudioResource>;
  video?: ReadonlyMap<string, Mp4VideoResource>;
  fallbackPoster?: PackagedMedia;
  fallbackVideoPoster?: PackagedMedia;
  slideIndexes: ReadonlyMap<string, number>;
  placeholders?: NormalizedPlaceholderDefinition[];
  warnings: ExportWarning[];
  slideId?: string;
  nextChartId: () => number;
  nativeElements?: readonly LayoutElement[];
  layoutId?: string;
  masterId?: string;
  sourcePart?: "slide" | "other";
  nativeInheritedPlaceholders?: NativeInheritedPlaceholderDefaults;
  nativeInheritedRuns?: NativeInheritedRuns;
}): PartContextResult {
  const relationships = [...options.baseRelationships];
  let nextRelationship = relationships.length + 1;
  let nextObject = 2;
  const chartParts: { id: number; chart: LayoutChartElement }[] = [];
  const audioTargets: { objectId: number; durationMs: number }[] = [];
  const videoTargets: { objectId: number; durationMs: number }[] = [];
  const mediaTargets: { kind: "audio" | "video"; objectId: number; durationMs: number }[] = [];
  const native = options.nativeElements ? nativeDrawingContext(options.nativeElements) : undefined;
  const context: ElementXmlContext = {
    theme: options.theme,
    ...(options.nativeInheritedPlaceholders ? { nativeInheritedPlaceholders: options.nativeInheritedPlaceholders } : {}),
    ...(options.nativeInheritedRuns ? { nativeInheritedRuns: options.nativeInheritedRuns } : {}),
    ...(options.textPath ? { textPath: options.textPath } : {}),
    ...(native ? { objectId: native.objectId, nativeConnector: native.connector } : {}),
    ...(options.layoutId ? { layoutId: options.layoutId } : {}),
    ...(options.masterId ? { masterId: options.masterId } : {}),
    ...(options.slideId !== undefined ? { slideId: options.slideId } : {}),
    warnings: options.warnings,
    placeholders: new Map((options.placeholders ?? []).map((placeholder) => [placeholder.key, placeholder])),
    placeholderIndices: placeholderIndices(options.placeholders ?? []),
    nextObjectId: native?.nextObjectId ?? (() => nextObject++),
    imageRelationship(assetId) {
      const packaged = options.media.get(assetId);
      if (packaged === undefined) {
        options.warnings.push({ code: "image-omitted", message: `Image asset "${assetId}" was omitted from the part.`, ...(options.slideId !== undefined ? { slideId: options.slideId } : {}), assetId, ...(options.layoutId !== undefined ? { layoutId: options.layoutId } : {}), ...(options.masterId !== undefined ? { masterId: options.masterId } : {}) });
        return undefined;
      }
      const id = `rId${nextRelationship++}`;
      relationships.push({ id, type: REL.image, target: `../media/${mediaFilename(packaged.name)}` });
      if (!packaged.svgName) return { base: id };
      const svg = `rId${nextRelationship++}`;
      relationships.push({ id: svg, type: REL.image, target: `../${packaged.svgName.slice(4)}` });
      return { base: id, svg };
    },
    audioRelationship(assetId, posterAssetId, objectId) {
      const packaged = options.media.get(assetId), resource = options.audio?.get(assetId);
      const poster = posterAssetId === undefined ? options.fallbackPoster : options.media.get(posterAssetId);
      if (!packaged || !resource || !poster || poster.loaded.mimeType.startsWith("audio/")) throw new TypeError("PPTX audio or its poster could not be packaged.");
      const audio = `rId${nextRelationship++}`, media = `rId${nextRelationship++}`, image = `rId${nextRelationship++}`;
      relationships.push({ id: audio, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio", target: `../media/${mediaFilename(packaged.name)}` });
      relationships.push({ id: media, type: "http://schemas.microsoft.com/office/2007/relationships/media", target: `../media/${mediaFilename(packaged.name)}` });
      relationships.push({ id: image, type: REL.image, target: `../media/${mediaFilename(poster.name)}` });
      audioTargets.push({ objectId, durationMs: resource.durationMs });
      mediaTargets.push({ kind: "audio", objectId, durationMs: resource.durationMs });
      return { audio, media, poster: image };
    },
    videoRelationship(assetId, posterAssetId, objectId) {
      const packaged = options.media.get(assetId), resource = options.video?.get(assetId);
      const poster = posterAssetId === undefined ? options.fallbackVideoPoster : options.media.get(posterAssetId);
      if (!packaged || !resource || !poster || !poster.loaded.mimeType.startsWith("image/")) throw new TypeError("PPTX video or its poster could not be packaged.");
      const video = `rId${nextRelationship++}`, media = `rId${nextRelationship++}`, image = `rId${nextRelationship++}`;
      relationships.push({ id: video, type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video", target: `../media/${mediaFilename(packaged.name)}` });
      relationships.push({ id: media, type: "http://schemas.microsoft.com/office/2007/relationships/media", target: `../media/${mediaFilename(packaged.name)}` });
      relationships.push({ id: image, type: REL.image, target: `../media/${mediaFilename(poster.name)}` });
      videoTargets.push({ objectId, durationMs: resource.durationMs });
      mediaTargets.push({ kind: "video", objectId, durationMs: resource.durationMs });
      return { video, media, poster: image };
    },
    actionRelationship(action: ElementAction) {
      const id = `rId${nextRelationship++}`;
      if (action.type === "url") {
        relationships.push({ id, type: REL.hyperlink, target: action.url, targetMode: "External" });
        return id;
      }
      const index = options.slideIndexes.get(action.slideId);
      if (index === undefined) return undefined;
      const target = `slide${index + 1}.xml`;
      relationships.push({
        id,
        type: REL.slide,
        target: options.sourcePart === "slide" ? target : `../slides/${target}`,
      });
      return id;
    },
    nextChartId: options.nextChartId,
    chartRelationship(chartId) {
      const id = `rId${nextRelationship++}`;
      relationships.push({ id, type: REL.chart, target: `../charts/chart${chartId}.xml` });
      return id;
    },
    registerChart(chartId, chart) {
      chartParts.push({ id: chartId, chart });
    },
  };
  return { context, relationships, chartParts, audioTargets, videoTargets, mediaTargets };
}

function elementsXml(elements: readonly LayoutElement[], context: ElementXmlContext): string[] {
  return elements.map((element) => elementXml(element, context)).filter((value) => value !== "");
}

export async function buildPackage(
  layout: LayoutResult,
  warnings: ExportWarning[],
  loadAsset: AssetLoader,
  loadContext: AssetLoadContext = {},
  connectorMode: "path" | "native" = "path",
  nativeInheritance?: NativePlaceholderInheritance,
  gifDiagnostics: "off" | "warn" | "reject-lzw" = "off",
  jpegOrientationDiagnostics: "off" | "warn" = "off",
  pngDiagnostics: "off" | "warn" | "reject-structure" | "reject-scanlines" = "off",
  deduplicateResolvedAssets = false,
  jpegIccDiagnostics: "off" | "warn" = "off",
  firstSlideNumber?: number,
  jpegValidation?: "structure",
  audio: ReadonlyMap<string, WaveAudioResource> = new Map(),
  video: ReadonlyMap<string, Mp4VideoResource> = new Map(),
): Promise<ZipPart[]> {
  const masters = layout.masters ?? [];
  const masterIndexes = new Map(masters.map((master, index) => [master.id, index]));
  const masterIndex = (id: string | undefined) => {
    if (id === undefined) return 1;
    const index = masterIndexes.get(id);
    if (index === undefined) throw new Error(`Unknown layout master "${id}".`);
    return index + 2;
  };
  for (const item of layout.layouts) { masterIndex(item.masterId); if (item.masterId !== undefined && !item.theme) throw new TypeError(`Layout "${item.id}" has no resolved master theme.`); }
  let chartCount = 0;
  const nextChartId = () => ++chartCount;
  const parts: ZipPart[] = [];
  const media = new Map<string, PackagedMedia>();
  const mediaBuckets = new Map<string, PackagedMedia[]>();
  let svgCount = 0;
  let generatedMediaId = 0;
  const reservedIds = new Set(layout.assets.map(asset => asset.id.toLowerCase()));
  const usedMediaNames = new Set<string>();
  const mediaName = (id: string, extension: string) => {
    const candidate = `${id}${extension}`;
    if (/^[A-Za-z0-9._-]+$/.test(id) && !usedMediaNames.has(candidate.toLowerCase())) {
      usedMediaNames.add(candidate.toLowerCase()); return candidate;
    }
    let name: string;
    do { name = `pptkit-asset-${++generatedMediaId}`; } while (reservedIds.has(name.toLowerCase()) || usedMediaNames.has(`${name}${extension}`.toLowerCase()));
    const result = `${name}${extension}`;
    usedMediaNames.add(result.toLowerCase()); return result;
  };
  for (const asset of layout.assets) {
    // Media resources have separate handling and must never enter image loaders.
    if (asset.kind === "audio" || asset.kind === "video") continue;
    try {
      loadContext.signal?.throwIfAborted();
      const loaded = await loadAsset(asset, loadContext);
      loadContext.signal?.throwIfAborted();
      const inspectPng = () => inspectApng(loaded.data, { ...(loadContext.signal ? { signal: loadContext.signal } : {}) });
      const pngMime = pngDiagnostics !== "off" && /^image\/a?png$/.test(loaded.mimeType.trim().toLowerCase());
      const pngCandidate = pngDiagnostics !== "off" && (pngMime || [137,80,78,71,13,10,26,10].every((value,index) => loaded.data[index] === value));
      const verifiedPng = pngCandidate && (pngDiagnostics === "reject-structure" || pngDiagnostics === "reject-scanlines") ? inspectPng() : undefined;
      if (verifiedPng) {
        if (!pngMime || loaded.extension !== ".png") throw new Error("PNG signature does not match the loaded media type or extension.");
        if (!verifiedPng.png.idatBytes) throw new Error("Invalid PNG: IDAT has no compressed data.");
        if (verifiedPng.png.unknownCriticalChunks.length) throw new Error(`PNG has an unsupported critical chunk: ${verifiedPng.png.unknownCriticalChunks[0]}.`);
        if (pngDiagnostics === "reject-scanlines") await validatePngScanlines(loaded.data, verifiedPng, loadContext.signal);
      }
      const gifMime = gifDiagnostics !== "off" && loaded.mimeType.trim().toLowerCase() === "image/gif";
      const gifCandidate = gifDiagnostics !== "off" && (gifMime || /^(GIF87a|GIF89a)$/.test(String.fromCharCode(...loaded.data.subarray(0, 6))));
      const gifOptions = { ...(loadContext.signal ? { signal: loadContext.signal } : {}) };
      const verifiedGif = gifCandidate && gifDiagnostics === "reject-lzw" ? inspectGifImageData(loaded.data, gifOptions) : undefined;
      if (verifiedGif && (!gifMime || loaded.extension !== ".gif")) throw new Error("GIF signature does not match the loaded media type or extension.");
      if (jpegValidation === "structure") {
        const jpegMime = loaded.mimeType.trim().toLowerCase() === "image/jpeg";
        if (jpegMime || (loaded.data[0] === 0xff && loaded.data[1] === 0xd8)) {
          validateJpegStructure(loaded.data, loadContext.signal);
          if (!jpegMime || loaded.extension !== ".jpg") throw new Error("JPEG signature does not match the loaded media type or extension.");
        }
      }
      if (loaded.orientationNormalization === "pixels") warnings.push({ code: "asset-orientation-normalized", assetId: asset.id,
        message: "The asset resolver reports EXIF orientation was normalized into pixels; image bytes or metadata may have changed, and authored geometry and crop remain unchanged." });
      if (loaded.animationFallback === "first-frame") warnings.push({ code: "asset-animation-flattened", assetId: asset.id,
        message: "The asset resolver reports that animated input was flattened to a first-frame image; animation and timing are lost." });
      if (loaded.vectorRasterization !== undefined) warnings.push({ code: "asset-vector-rasterized", assetId: asset.id,
        message: `The asset resolver reports ${loaded.vectorRasterization.toUpperCase()} was rasterized; vector editability and any animation are lost, and fidelity is caller-owned.` });
      if (loaded.svgPair !== undefined) warnings.push({ code: "asset-svg-pair-unverified", assetId: asset.id,
        message: "SVG and fallback bytes are caller-supplied; their validity, correspondence, external resources and application rendering are not verified." });
      if (gifCandidate) {
        try {
          const info = verifiedGif ?? inspectGif(loaded.data, gifOptions);
          if (info.frames.length > 1) warnings.push({ code: "gif-animation-playback-unverified", assetId: asset.id,
            message: "Animated GIF bytes are preserved; playback and timing depend on the target application." });
          if (info.plainTextBlocks || info.unknownExtensionBlocks || info.frames.some(frame => frame.userInputFlag || frame.reservedControlBits || frame.disposal >= 4)) warnings.push({ code: "gif-inspection-incomplete", assetId: asset.id,
            message: "GIF contains plain-text graphics, unrecognized extensions, a user-input flag or reserved graphic-control values whose visual and timing effects were not verified." });
        } catch (error) {
          loadContext.signal?.throwIfAborted();
          warnings.push({ code: "gif-inspection-incomplete", assetId: asset.id,
            message: `GIF structure could not be fully inspected: ${error instanceof Error ? error.message : String(error)}` });
        }
      }
      if (jpegOrientationDiagnostics === "warn") {
        const jpegSignature = loaded.data.length >= 2 && loaded.data[0] === 0xff && loaded.data[1] === 0xd8;
        if (jpegSignature || loaded.mimeType.trim().toLowerCase() === "image/jpeg") {
          try {
            const info = inspectJpegExifOrientation(loaded.data, { ...(loadContext.signal ? { signal: loadContext.signal } : {}) });
            if (info.orientation !== undefined && info.orientation !== 1) warnings.push({ code: "jpeg-exif-orientation-unverified", assetId: asset.id,
              message: `JPEG EXIF orientation ${info.orientation} is preserved; target application and preview orientation handling may differ.` });
          } catch (error) {
            loadContext.signal?.throwIfAborted();
            warnings.push({ code: "jpeg-exif-inspection-incomplete", assetId: asset.id,
              message: `JPEG pre-scan EXIF orientation could not be inspected: ${error instanceof Error ? error.message : String(error)}` });
          }
        }
      }
      if (jpegIccDiagnostics === "warn") {
        const jpegSignature = loaded.data.length >= 2 && loaded.data[0] === 0xff && loaded.data[1] === 0xd8;
        if (jpegSignature || loaded.mimeType.trim().toLowerCase() === "image/jpeg") {
          try {
            const info = inspectJpegIcc(loaded.data, { ...(loadContext.signal ? { signal: loadContext.signal } : {}) });
            if (info.status === "complete") warnings.push({ code: "jpeg-icc-color-unverified", assetId: asset.id, message: "JPEG ICC chunks are present; profile validity, color conversion and application fidelity are not verified." });
            if (info.status === "incomplete") warnings.push({ code: "jpeg-icc-inspection-incomplete", assetId: asset.id, message: "JPEG ICC chunks are incomplete before the first scan; color behavior is not verified." });
          } catch (error) {
            loadContext.signal?.throwIfAborted();
            warnings.push({ code: "jpeg-icc-inspection-incomplete", assetId: asset.id, message: `JPEG ICC structure could not be fully inspected: ${error instanceof Error ? error.message : String(error)}` });
          }
        }
      }
      if (pngCandidate) {
        try {
          const info = verifiedPng ?? inspectPng();
          if (info.animation) warnings.push({ code: "apng-animation-playback-unverified", assetId: asset.id, message: "APNG bytes are preserved; animation playback and timing depend on the target application." });
          if (info.png.iccProfileChunks) {
            warnings.push({ code: "png-icc-color-unverified", assetId: asset.id, message: "PNG iCCP chunks are present; profile validity, color conversion and application fidelity are not verified." });
            try { inspectCapturedPngIcc(loaded.data, { ...(loadContext.signal ? { signal: loadContext.signal } : {}) }); }
            catch (error) {
              loadContext.signal?.throwIfAborted();
              warnings.push({ code: "png-icc-inspection-incomplete", assetId: asset.id, message: `PNG ICC envelope could not be fully inspected: ${error instanceof Error ? error.message : String(error)}` });
            }
          }
          if (info.png.unknownCriticalChunks.length) warnings.push({ code: "png-inspection-incomplete", assetId: asset.id, message: "PNG contains unrecognized critical chunks; their decoding and visual effects were not inspected." });
        } catch (error) {
          loadContext.signal?.throwIfAborted();
          warnings.push({ code: "png-inspection-incomplete", assetId: asset.id, message: `PNG/APNG structure could not be fully inspected: ${error instanceof Error ? error.message : String(error)}` });
        }
      }
      const bucketKey = deduplicateResolvedAssets ? JSON.stringify([loaded.extension, loaded.mimeType.trim().toLowerCase(), byteFingerprint(loaded.data), byteFingerprint(loaded.svgPair)]) : undefined;
      const bucket = bucketKey === undefined ? undefined : mediaBuckets.get(bucketKey);
      const existing = bucket?.find(item => sameBytes(item.loaded.data, loaded.data) && sameBytes(item.loaded.svgPair, loaded.svgPair));
      if (existing) { media.set(asset.id, existing); continue; }
      const name = `ppt/media/${mediaName(asset.id, loaded.extension)}`;
      const svgName = loaded.svgPair ? `ppt/media/pptkit-svg-pairs/${++svgCount}.svg` : undefined;
      const packaged = { name, loaded, ...(svgName ? { svgName } : {}) };
      media.set(asset.id, packaged);
      if (bucketKey !== undefined) { if (bucket) bucket.push(packaged); else mediaBuckets.set(bucketKey, [packaged]); }
      parts.push({ name, data: loaded.data });
      if (svgName && loaded.svgPair) parts.push({ name: svgName, data: loaded.svgPair });
    } catch (error) {
      loadContext.signal?.throwIfAborted();
      warnings.push({ code: "asset-read-failed", message: `Could not package image asset "${asset.id}": ${error instanceof Error ? error.message : String(error)}`, assetId: asset.id });
    }
  }

  for (const [id, resource] of audio) {
    const name = `ppt/media/${mediaName(id, ".wav")}`;
    const packaged = { name, loaded: { data: resource.data, mimeType: "audio/x-wav", extension: ".wav" } };
    media.set(id, packaged); parts.push({ name, data: resource.data });
  }
  for (const [id, resource] of video) {
    const name = `ppt/media/${mediaName(id, ".mp4")}`;
    const packaged = { name, loaded: { data: resource.data, mimeType: "video/mp4", extension: ".mp4" } };
    media.set(id, packaged); parts.push({ name, data: resource.data });
  }
  let fallbackPoster: PackagedMedia | undefined;
  if (layout.slides.some(slide => slide.elements.some(element => element.type === "audio" && element.posterAssetId === undefined))) {
    const data = decodeAssetBytes("iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAZElEQVR42mP4/OXb/4HEDKMOGHXAqAOGnQPiM4sHzgEgywfEATCLB8QB6JZTxQHYDMVmODFqaOYAYh1JdQcQg0cdMLwdMOCJkC7ZcMALoiFVFA+KymjQVMejLaJRB4w6YEg6AACOBIj3YosGvAAAAABJRU5ErkJggg==");
    const name = `ppt/media/${mediaName("pptkit-audio-poster", ".png")}`;
    fallbackPoster = { name, loaded: { data, mimeType: "image/png", extension: ".png" } };
    let key = "pptkit-audio-poster"; while (media.has(key)) key += "-";
    media.set(key, fallbackPoster); parts.push({ name, data });
  }
  let fallbackVideoPoster: PackagedMedia | undefined;
  if (layout.slides.some(slide => slide.elements.some(element => element.type === "video" && element.posterAssetId === undefined))) {
    const data = decodeAssetBytes("iVBORw0KGgoAAAANSUhEUgAAACAAAAASCAYAAAA6yNxSAAAARElEQVR42mN4/fr1/4HEDKMOGHYOCA0NHRgHgCyGYbo6ANliujoAm8V0cQA+i0eGAwY8CgZNIhw02XDQFESjdcGQdQAAbRsYlRxPJ3MAAAAASUVORK5CYII=");
    const name = `ppt/media/${mediaName("pptkit-video-poster", ".png")}`;
    fallbackVideoPoster = { name, loaded: { data, mimeType: "image/png", extension: ".png" } };
    let key = "pptkit-video-poster"; while (media.has(key)) key += "-";
    media.set(key, fallbackVideoPoster); parts.push({ name, data });
  }

  const slideIndexes = new Map(layout.slides.map((slide, index) => [slide.id, index]));
  const layoutIndexes = new Map(layout.layouts.map((item, index) => [item.id, index]));
  const allChartParts: { id: number; chart: LayoutChartElement }[] = [];

  for (let index = 0; index < layout.layouts.length; index += 1) {
    const item = layout.layouts[index]!;
    const partContext = createPartContext({
      theme: item.theme ?? layout.theme,
      layoutId: item.id,
      ...(nativeInheritance?.layouts.get(item.id) ? { nativeInheritedPlaceholders: nativeInheritance.layouts.get(item.id)! } : {}),
      nextChartId,
      baseRelationships: [{ id: "rId1", type: REL.slideMaster, target: `../slideMasters/slideMaster${masterIndex(item.masterId)}.xml` }],
      media,
      slideIndexes,
      ...(connectorMode === "native" ? { nativeElements: item.elements } : {}),
      placeholders: item.placeholders,
      warnings,
    });
    const xml = elementsXml(item.elements, partContext.context);
    allChartParts.push(...partContext.chartParts);
    parts.push({ name: `ppt/slideLayouts/slideLayout${index + 1}.xml`, data: encodeUtf8(slideLayoutXml(item, xml, partContext.context)) });
    parts.push({ name: `ppt/slideLayouts/_rels/slideLayout${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml(partContext.relationships)) });
  }

  for (let index = 0; index < layout.slides.length; index += 1) {
    const slide = layout.slides[index]!;
    const layoutIndex = layoutIndexes.get(slide.layoutId)!;
    const slideLayout = layout.layouts[layoutIndex]!;
    const partContext = createPartContext({
      theme: slideLayout.theme ?? layout.theme,
      nextChartId,
      baseRelationships: [
        { id: "rId1", type: REL.slideLayout, target: `../slideLayouts/slideLayout${layoutIndex + 1}.xml` },
        { id: "rId2", type: REL.notesSlide, target: `../notesSlides/notesSlide${index + 1}.xml` },
      ],
      media,
      audio,
      video,
      ...(fallbackPoster ? { fallbackPoster } : {}),
      ...(fallbackVideoPoster ? { fallbackVideoPoster } : {}),
      slideIndexes,
      ...(connectorMode === "native" ? { nativeElements: slide.elements } : {}),
      placeholders: slideLayout.placeholders,
      warnings,
      slideId: slide.id,
      ...(nativeInheritance?.slides.get(slide.id) ? { nativeInheritedRuns: nativeInheritance.slides.get(slide.id)! } : {}),
      sourcePart: "slide",
    });
    const xml = elementsXml(slide.elements, partContext.context);
    allChartParts.push(...partContext.chartParts);
    parts.push({
      name: `ppt/slides/slide${index + 1}.xml`,
      data: encodeUtf8(slideXml(xml, slide.backgroundSource === "slide" ? slide.background : undefined, slide.hidden, {
        ...(slide.section !== undefined ? { section: slide.section } : {}),
        ...(slide.showInheritedShapes === undefined ? {} : { showInheritedShapes: slide.showInheritedShapes }),
        tags: slide.tags,
        customData: slide.customData,
        ...(slide.transition !== undefined ? { transition: slide.transition } : {}),
        ...(partContext.audioTargets.length && partContext.videoTargets.length ? { mixedMediaTiming: partContext.mediaTargets } : partContext.audioTargets.length ? { audioTiming: partContext.audioTargets } : partContext.videoTargets.length ? { videoTiming: partContext.videoTargets } : {}),
      })),
    });
    parts.push({ name: `ppt/slides/_rels/slide${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml(partContext.relationships)) });

    const notesContext = createPartContext({
      theme: layout.theme,
      textPath: "notes",
      nextChartId,
      baseRelationships: [
        { id: "rId1", type: REL.notesMaster, target: "../notesMasters/notesMaster1.xml" },
        { id: "rId2", type: REL.slide, target: `../slides/slide${index + 1}.xml` },
      ],
      media,
      slideIndexes,
      warnings,
      slideId: slide.id,
    });
    parts.push({ name: `ppt/notesSlides/notesSlide${index + 1}.xml`, data: encodeUtf8(notesSlideXml(slide.notes, notesContext.context)) });
    parts.push({ name: `ppt/notesSlides/_rels/notesSlide${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml(notesContext.relationships)) });
  }

  const masterParts = masters.map((master, index) => {
    const indices = layout.layouts.flatMap((item, i) => item.masterId === master.id ? [i] : []);
    const owned = indices.map(i => layout.layouts[i]!);
    const relationships: Relationship[] = indices.map((i, j) => ({ id: `rId${j + 1}`, type: REL.slideLayout, target: `../slideLayouts/slideLayout${i + 1}.xml` }));
    relationships.push({ id: `rId${owned.length + 1}`, type: REL.theme, target: `../theme/theme${index + 3}.xml` });
    const partContext = createPartContext({ theme: master.theme, masterId: master.id, nextChartId, baseRelationships: relationships, media, slideIndexes, warnings,
      ...(connectorMode === "native" ? { nativeElements: master.elements ?? [] } : {}),
    });
    const artwork = elementsXml(master.elements ?? [], partContext.context);
    allChartParts.push(...partContext.chartParts);
    return { xml: slideMasterXml(owned, { ...master, layoutIndexes: indices, preserve: true }, artwork), relationships: partContext.relationships };
  });

  for (const chartPart of allChartParts) {
    const chart = chartPart.chart;
    const bubble = chart.chartType === "bubble" ? bubbleSeriesParts(chart.series) : undefined;
    parts.push({ name: `ppt/charts/chart${chartPart.id}.xml`, data: encodeUtf8(chartPartXml(chart, bubble?.xml)) });
    parts.push({ name: `ppt/charts/_rels/chart${chartPart.id}.xml.rels`, data: encodeUtf8(relationshipsXml([
      { id: "rId1", type: REL.package, target: `../embeddings/Microsoft_Excel_Worksheet${chartPart.id}.xlsx` },
    ])) });
    parts.push({ name: `ppt/embeddings/Microsoft_Excel_Worksheet${chartPart.id}.xlsx`, data: chart.chartType === "bubble" ? bubble!.workbook : chart.chartType === "scatter" ? scatterWorkbook(chart.series) : chartWorkbook(chart) });
  }

  const presentationRelationships: Relationship[] = [
    { id: "rId1", type: REL.slideMaster, target: "slideMasters/slideMaster1.xml" },
  ];
  layout.slides.forEach((_, index) => presentationRelationships.push({ id: `rId${index + 2}`, type: REL.slide, target: `slides/slide${index + 1}.xml` }));
  presentationRelationships.push(
    { id: `rId${layout.slideCount + 2}`, type: REL.notesMaster, target: "notesMasters/notesMaster1.xml" },
    { id: `rId${layout.slideCount + 3}`, type: REL.presProps, target: "presProps.xml" },
    { id: `rId${layout.slideCount + 4}`, type: REL.viewProps, target: "viewProps.xml" },
    { id: `rId${layout.slideCount + 5}`, type: REL.theme, target: "theme/theme1.xml" },
    { id: `rId${layout.slideCount + 6}`, type: REL.tableStyles, target: "tableStyles.xml" },
  );
  const defaultLayouts = layout.layouts.filter(item => item.masterId === undefined);
  const masterRelationships: Relationship[] = defaultLayouts.map((_, index) => ({ id: `rId${index + 1}`, type: REL.slideLayout, target: `../slideLayouts/slideLayout${layout.layouts.indexOf(defaultLayouts[index]!) + 1}.xml` }));
  masterRelationships.push({ id: `rId${defaultLayouts.length + 1}`, type: REL.theme, target: "../theme/theme1.xml" });

  masters.forEach((master, index) => {
    presentationRelationships.push({ id: `rId${layout.slideCount + 7 + index}`, type: REL.slideMaster, target: `slideMasters/slideMaster${index + 2}.xml` });
    parts.push({ name: `ppt/slideMasters/slideMaster${index + 2}.xml`, data: encodeUtf8(masterParts[index]!.xml) });
    parts.push({ name: `ppt/slideMasters/_rels/slideMaster${index + 2}.xml.rels`, data: encodeUtf8(relationshipsXml(masterParts[index]!.relationships)) });
    parts.push({ name: `ppt/theme/theme${index + 3}.xml`, data: encodeUtf8(themeXml(master.theme)) });
  });
  parts.push(
    { name: "[Content_Types].xml", data: encodeUtf8(contentTypesXml(media, layout.slideCount, layout.layouts.length, allChartParts.length, masters.length, layout.metadata.customProperties !== undefined)) },
    { name: "_rels/.rels", data: encodeUtf8(rootRelationshipsXml(layout.metadata.customProperties !== undefined)) },
    { name: "ppt/presentation.xml", data: encodeUtf8(presentationXml(layout, firstSlideNumber)) },
    { name: "ppt/_rels/presentation.xml.rels", data: encodeUtf8(relationshipsXml(presentationRelationships)) },
    { name: "ppt/presProps.xml", data: encodeUtf8(presentationPropertiesXml()) },
    { name: "ppt/viewProps.xml", data: encodeUtf8(viewPropertiesXml()) },
    { name: "ppt/tableStyles.xml", data: encodeUtf8(tableStylesXml()) },
    { name: "ppt/slideMasters/slideMaster1.xml", data: encodeUtf8(slideMasterXml(defaultLayouts, masters.length ? { layoutIndexes: defaultLayouts.map(item => layout.layouts.indexOf(item)) } : undefined)) },
    { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: encodeUtf8(relationshipsXml(masterRelationships)) },
    { name: "ppt/theme/theme1.xml", data: encodeUtf8(themeXml(layout.theme)) },
    { name: "ppt/theme/theme2.xml", data: encodeUtf8(themeXml(layout.theme)) },
    { name: "ppt/notesMasters/notesMaster1.xml", data: encodeUtf8(notesMasterXml(layout.size)) },
    { name: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: encodeUtf8(relationshipsXml([
      { id: "rId1", type: REL.theme, target: "../theme/theme2.xml" },
    ])) },
    { name: "docProps/core.xml", data: encodeUtf8(corePropertiesXml(layout.metadata)) },
    { name: "docProps/app.xml", data: encodeUtf8(appPropertiesXml(layout.slideCount, layout.metadata, layout.size)) },
    ...(layout.metadata.customProperties === undefined ? [] : [{ name: "docProps/custom.xml", data: encodeUtf8(customPropertiesXml(layout.metadata)) }]),
  );
  return parts;
}
