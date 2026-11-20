import { chartHoleWarnings } from "./chart-hole-warnings.js";
import { backgroundGradientWarnings } from "./background-gradient-warnings.js";
import { nativePlaceholderInheritance } from "./native-placeholder-inheritance.js";
import { prepareEventTimelineOutput } from "./event-timelines.js";
import { createAssetLoader } from "../assets/create-asset-loader.js";
import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import { resolveNormalizedLayout } from "@pptkit/layout";
import { createZip } from "../archive/create-zip.js";
import { buildPackage } from "../packaging/build-package.js";
import type { ExportWarning, GeneratePptxOptions, GeneratePptxResult } from "../types/export.js";
import type { AssetLoader } from "../types/internal.js";
import { emu } from "../ooxml/xml.js";
import { listContinuationWarnings } from "./list-continuation-warnings.js";
import { planWaveAudio } from "./wave-audio.js";
import { rejectVideoOutput } from "./video-preflight.js";

export async function generatePptxWith(
  document: PresentationDocument,
  loadAsset: AssetLoader,
  options?: GeneratePptxOptions,
): Promise<GeneratePptxResult> {
  const { signal, resolveAsset, maxAssetBytes, audioMode, externalImageValidation, connectorMode = "path", nativePlaceholderText = "materialize", eventTimelineMode = "reject", gifDiagnostics = "off", pngDiagnostics = "off", jpegOrientationDiagnostics = "off", jpegIccDiagnostics = "off", jpegValidation, deduplicateResolvedAssets = false, firstSlideNumber, layout: layoutOptions } = options ?? {};
  signal?.throwIfAborted();
  if (firstSlideNumber !== undefined && (!Number.isSafeInteger(firstSlideNumber) || firstSlideNumber < 0 || firstSlideNumber > 9999)) throw new RangeError("firstSlideNumber must be an integer from 0 to 9999.");
  if (typeof deduplicateResolvedAssets !== "boolean") throw new TypeError("deduplicateResolvedAssets must be a boolean.");
  if (jpegOrientationDiagnostics !== "off" && jpegOrientationDiagnostics !== "warn") throw new TypeError("jpegOrientationDiagnostics must be off or warn.");
  if (jpegIccDiagnostics !== "off" && jpegIccDiagnostics !== "warn") throw new TypeError("jpegIccDiagnostics must be off or warn.");
  if (jpegValidation !== undefined && jpegValidation !== "structure") throw new TypeError("jpegValidation must be structure when provided.");
  if (gifDiagnostics !== "off" && gifDiagnostics !== "warn" && gifDiagnostics !== "reject-lzw") throw new TypeError("gifDiagnostics must be off, warn or reject-lzw.");
  if (pngDiagnostics !== "off" && pngDiagnostics !== "warn" && pngDiagnostics !== "reject-structure" && pngDiagnostics !== "reject-scanlines") throw new TypeError("pngDiagnostics must be off, warn, reject-structure or reject-scanlines.");
  if (eventTimelineMode !== "reject" && eventTimelineMode !== "omit") throw new TypeError("eventTimelineMode must be reject or omit.");
  if (nativePlaceholderText !== "materialize" && nativePlaceholderText !== "inherit") throw new TypeError("nativePlaceholderText must be materialize or inherit.");
  if (connectorMode !== "path" && connectorMode !== "native") throw new TypeError("connectorMode must be path or native.");
  if (maxAssetBytes !== undefined && (!Number.isSafeInteger(maxAssetBytes) || maxAssetBytes <= 0)) throw new RangeError("maxAssetBytes must be a positive safe integer.");
  if (externalImageValidation !== undefined && externalImageValidation !== "signature") throw new TypeError("externalImageValidation must be signature when provided.");
  if (resolveAsset !== undefined && typeof resolveAsset !== "function") throw new TypeError("resolveAsset must be a function.");
  if (audioMode !== undefined && audioMode !== "wave-pcm-v1") throw new TypeError("audioMode must be wave-pcm-v1 when provided.");
  const loadContext = { ...(signal ? { signal } : {}), ...(maxAssetBytes !== undefined ? { maxAssetBytes } : {}), ...(externalImageValidation ? { externalImageValidation } : {}) };
  const resolvedLoader = createAssetLoader(loadAsset, resolveAsset, signal, maxAssetBytes);
  const normalized = normalizePresentation(document);
  rejectVideoOutput(normalized);
  const audio = planWaveAudio(normalized, audioMode, maxAssetBytes, signal);
  // PresentationML ST_SlideSizeCoordinate bounds apply to the serialized
  // integer, not the pre-rounding Core value. Reject before any asset I/O.
  if ([normalized.size.width, normalized.size.height].some(value => {
    const coordinate = emu(value);
    return !Number.isSafeInteger(coordinate) || coordinate < 914400 || coordinate > 51206400;
  })) throw new RangeError("PPTX slide width and height must serialize to 914400–51206400 EMU (72–4032 pt).");
  const effectiveFirstSlideNumber = firstSlideNumber ?? normalized.firstSlideNumber;
  if (effectiveFirstSlideNumber !== undefined && effectiveFirstSlideNumber > 9999) throw new RangeError("The native first slide number must be from 0 to 9999.");
  const timelineWarnings = prepareEventTimelineOutput(normalized, eventTimelineMode);
  const layout = resolveNormalizedLayout(normalized, layoutOptions);
  const warnings: ExportWarning[] = [
    ...timelineWarnings,
    ...normalized.slides.flatMap((slide, index) => slide.readingOrder?.length ? [{ code: "reading-order-not-exported", message: "PPTX does not retain authored independent reading order.", slideId: slide.id, path: `slides.${index}.readingOrder` }] : []),
    ...chartHoleWarnings(document),
    ...backgroundGradientWarnings(document),
    ...listContinuationWarnings(document),
    ...(layout.analysis?.diagnostics ?? []).map(({ severity, bounds, available, ...warning }) => warning),
  ];
  const nativeInheritance = nativePlaceholderInheritance(document, layout, warnings, nativePlaceholderText);
  const parts = await buildPackage(layout, warnings, resolvedLoader, loadContext, connectorMode, nativeInheritance, gifDiagnostics, jpegOrientationDiagnostics, pngDiagnostics, deduplicateResolvedAssets, jpegIccDiagnostics, effectiveFirstSlideNumber, jpegValidation, audio);
  signal?.throwIfAborted();
  const bytes = createZip(parts);
  signal?.throwIfAborted();
  return {
    bytes,
    slideCount: layout.slideCount,
    byteLength: bytes.byteLength,
    warnings,
    status: warnings.length === 0 ? "generated" : "generated-with-warnings",
  };
}
