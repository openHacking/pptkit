import { normalizePresentation, validatePresentation, type PresentationDocument } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";
import {
  authorPresentation,
  inspectPptxPackage,
  inspectStructure,
  NOT_RUN_PACKAGE_CHECK,
  validateDeckSpec,
  type BuildReport,
  type DeckSessionV1,
  type StructuralIssue,
} from "@pptkit/presentation-workflow";
import { renderPresentationToSvg, type SvgRenderResult } from "@pptkit/svg-renderer";

import { listTransfers, loadAssetBlob, loadSession, type StoredTransfer } from "./storage.js";
import {
  MAX_TRANSFER_CHUNK_BYTES,
  PPTKIT_TRANSFER_PROTOCOL,
  receiveTransferChunk,
  type TransferProgress,
} from "./transfer.js";
import "./styles.css";

const byId = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const deckTitle = byId<HTMLHeadingElement>("deck-title");
const transferInput = byId<HTMLTextAreaElement>("transfer-input");
const transferButton = byId<HTMLButtonElement>("transfer-submit");
const transferError = byId<HTMLParagraphElement>("transfer-error");
const transferProgress = byId<HTMLDivElement>("transfer-progress");
const status = byId<HTMLSpanElement>("status");
const deckMeta = byId<HTMLSpanElement>("deck-meta");
const thumbnails = byId<HTMLDivElement>("thumbnails");
const stage = byId<HTMLDivElement>("stage");
const issuesContainer = byId<HTMLDivElement>("issues");
const previousButton = byId<HTMLButtonElement>("previous");
const nextButton = byId<HTMLButtonElement>("next");
const pageStatus = byId<HTMLSpanElement>("page-status");
const downloadButton = byId<HTMLButtonElement>("download");

let session: DeckSessionV1 | undefined;
let presentation: PresentationDocument | undefined;
let preview: SvgRenderResult | undefined;
let currentIndex = 0;
let findings: StructuralIssue[] = [];
let currentDiagnostics: BuildReport["diagnostics"] = [];
let currentExportStatus: BuildReport["exportStatus"] = "not-run";
let objectUrls: string[] = [];
let persisted = true;
let transfers: TransferProgress[] = [];

function revokeObjectUrls() {
  for (const url of objectUrls) URL.revokeObjectURL(url);
  objectUrls = [];
}

async function createAssetResolver(activeSession: DeckSessionV1) {
  const resolved = new Map<string, { source: { type: "url"; value: string }; mimeType: string; dedupeKey: string }>();
  for (const asset of activeSession.assets) {
    const blob = await loadAssetBlob(activeSession.id, asset).catch(() => undefined);
    if (!blob) continue;
    const url = URL.createObjectURL(blob);
    objectUrls.push(url);
    resolved.set(asset.id, { source: { type: "url", value: url }, mimeType: asset.mimeType, dedupeKey: asset.id });
  }
  return (assetId: string) => resolved.get(assetId);
}

function changedSlides(previous: DeckSessionV1 | undefined, next: DeckSessionV1) {
  if (!previous || previous.id !== next.id) return [];
  const before = new Map(previous.deck.slides.map((slide) => [slide.id, JSON.stringify(slide)]));
  return next.deck.slides.filter((slide) => before.get(slide.id) !== JSON.stringify(slide)).map((slide) => slide.id);
}

function showCurrentSlide() {
  const slides = preview?.slides ?? [];
  const selected = slides[currentIndex];
  stage.replaceChildren();
  if (selected) {
    stage.innerHTML = selected.svg;
    stage.setAttribute("aria-label", `Slide ${currentIndex + 1}: ${selected.slideId}`);
  } else {
    stage.innerHTML = "<p>No visible slide.</p>";
  }
  pageStatus.textContent = `${slides.length === 0 ? 0 : currentIndex + 1} / ${slides.length}`;
  previousButton.disabled = currentIndex <= 0;
  nextButton.disabled = currentIndex >= slides.length - 1;
  for (const button of thumbnails.querySelectorAll<HTMLButtonElement>("button")) {
    button.setAttribute("aria-current", button.dataset.index === String(currentIndex) ? "page" : "false");
  }
}

function selectSlide(index: number) {
  if (!preview?.slides.length) return;
  currentIndex = Math.max(0, Math.min(index, preview.slides.length - 1));
  showCurrentSlide();
}

function showFindings(items: StructuralIssue[]) {
  issuesContainer.replaceChildren();
  if (items.length === 0) {
    issuesContainer.innerHTML = '<p class="success">No blocking findings.</p>';
    return;
  }
  for (const item of items) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `finding ${item.severity}`;
    const code = document.createElement("strong");
    code.textContent = item.code;
    const message = document.createElement("span");
    message.textContent = item.message;
    button.append(code, message);
    if (item.slideId) {
      button.addEventListener("click", () => {
        const index = preview?.slides.findIndex((slide) => slide.slideId === item.slideId) ?? -1;
        if (index >= 0) selectSlide(index);
      });
    } else {
      button.disabled = true;
    }
    issuesContainer.append(button);
  }
}

function renderThumbnails() {
  thumbnails.replaceChildren();
  for (const slide of preview?.slides ?? []) {
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.index = String(slide.index);
    button.setAttribute("aria-label", `Show slide ${slide.index + 1}: ${slide.slideId}`);
    button.innerHTML = `<span>${slide.index + 1}</span>${slide.svg}`;
    button.addEventListener("click", () => selectSlide(slide.index));
    thumbnails.append(button);
  }
}

function renderTransferProgress() {
  transferProgress.replaceChildren();
  if (transfers.length === 0) {
    transferProgress.innerHTML = "<p>No active transfers.</p>";
    return;
  }
  for (const item of transfers) {
    const row = document.createElement("p");
    row.dataset.transferId = item.transferId;
    row.className = `transfer-status ${item.status}`;
    row.textContent = `${item.kind}:${item.payloadId} · ${item.received.length}/${item.chunkCount} chunks · ${item.status}${item.error ? ` · ${item.error}` : ""}`;
    transferProgress.append(row);
  }
}

function recordTransfer(item: TransferProgress) {
  transfers = [...transfers.filter((candidate) => candidate.transferId !== item.transferId), item];
  renderTransferProgress();
}

function storedTransferProgress(item: StoredTransfer): TransferProgress {
  const received = [...item.received].sort((left, right) => left - right);
  const receivedSet = new Set(received);
  return {
    transferId: item.transferId,
    kind: item.kind,
    payloadId: item.payloadId,
    received,
    missing: Array.from({ length: item.chunkCount }, (_, index) => index).filter((index) => !receivedSet.has(index)),
    chunkCount: item.chunkCount,
    status: item.status,
    ...(item.error ? { error: item.error } : {}),
  };
}

async function refreshStoredTransfers() {
  const stored = (await listTransfers()).map(storedTransferProgress);
  const storedIds = new Set(stored.map((item) => item.transferId));
  transfers = [...transfers.filter((item) => item.status === "completed" && !storedIds.has(item.transferId)), ...stored];
  renderTransferProgress();
}

async function renderSession(nextSession: DeckSessionV1, changed: string[] = []) {
  const selectedSlideId = preview?.slides[currentIndex]?.slideId;
  revokeObjectUrls();
  const resolveAsset = await createAssetResolver(nextSession);
  presentation = authorPresentation(nextSession.deck, resolveAsset);
  const availableAssets = new Set(nextSession.assets.filter((asset) => resolveAsset(asset.id)).map((asset) => asset.id));
  const specIssues = validateDeckSpec(nextSession.deck, availableAssets);
  const diagnostics = validatePresentation(presentation);
  currentDiagnostics = diagnostics.map((diagnostic) => ({ severity: diagnostic.severity, code: diagnostic.code, message: diagnostic.message, path: diagnostic.path }));
  currentExportStatus = "not-run";
  const diagnosticIssues: StructuralIssue[] = diagnostics.map((diagnostic) => ({
    severity: diagnostic.severity === "error" ? "error" : "warning",
    code: diagnostic.code,
    message: diagnostic.message,
    ...(diagnostic.slideId ? { slideId: diagnostic.slideId } : {}),
    ...(diagnostic.elementId ? { elementId: diagnostic.elementId } : {}),
  }));
  const structural = diagnostics.some((diagnostic) => diagnostic.severity === "error") ? [] : inspectStructure(normalizePresentation(presentation));
  preview = await renderPresentationToSvg(presentation);
  findings = [
    ...specIssues,
    ...diagnosticIssues,
    ...structural,
    ...preview.warnings.map((warning) => ({ severity: "warning" as const, code: warning.code, message: warning.message, ...(warning.slideId ? { slideId: warning.slideId } : {}), ...(warning.elementId ? { elementId: warning.elementId } : {}) })),
  ];
  renderThumbnails();
  const retainedIndex = selectedSlideId ? preview.slides.findIndex((slide) => slide.slideId === selectedSlideId) : -1;
  currentIndex = retainedIndex >= 0 ? retainedIndex : 0;
  showCurrentSlide();
  showFindings(findings);
  deckTitle.textContent = nextSession.deck.brief.title;
  deckMeta.textContent = `${nextSession.deck.brief.themeId} · ${preview.slides.length} slides · revision ${nextSession.revision}`;
  const blocking = findings.filter((item) => item.severity === "error").length;
  downloadButton.disabled = blocking > 0;
  const changedText = changed.length > 0 ? ` Changed slides: ${changed.join(", ")}.` : "";
  status.textContent = `${persisted ? "Saved in this browser." : "Previewing in memory; browser storage unavailable."} ${blocking} blocking findings, ${findings.length - blocking} warnings.${changedText}`;
}

function download(bytes: Uint8Array, mimeType: string, filename: string) {
  const url = URL.createObjectURL(new Blob([bytes.slice().buffer], { type: mimeType }));
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

async function generateAndDownload() {
  if (!session || !presentation || findings.some((item) => item.severity === "error")) return;
  downloadButton.disabled = true;
  status.textContent = "Generating and inspecting PPTX in this browser…";
  try {
    const result = await generatePptx(presentation);
    const packageChecks = inspectPptxPackage(result.bytes);
    const exportIssues: StructuralIssue[] = result.warnings.map((warning) => ({ severity: "error", code: `export-${warning.code}`, message: warning.message, ...(warning.slideId ? { slideId: warning.slideId } : {}) }));
    if (!packageChecks.valid) exportIssues.push(...packageChecks.issues.map((message) => ({ severity: "error" as const, code: "invalid-package", message })));
    if (packageChecks.slideParts !== presentation.slides.length) exportIssues.push({ severity: "error", code: "slide-part-count", message: `Expected ${presentation.slides.length} slide parts, found ${packageChecks.slideParts}.` });
    findings = [...findings, ...exportIssues];
    showFindings(findings);
    const report: BuildReport = {
      runtime: "browser", sessionId: session.id, slideCount: result.slideCount, byteLength: result.byteLength,
      diagnostics: currentDiagnostics, exportWarnings: result.warnings, structuralIssues: findings, packageChecks,
      previewStatus: preview?.status ?? "failed", exportStatus: exportIssues.some((issue) => issue.severity === "error") ? "failed" : "generated",
      renderStatus: "not-run", generatedAt: new Date().toISOString(),
    };
    currentExportStatus = report.exportStatus;
    download(new TextEncoder().encode(`${JSON.stringify(report, null, 2)}\n`), "application/json", "build-report.json");
    if (report.exportStatus === "generated") {
      download(result.bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", `${session.id}.pptx`);
      status.textContent = `Generated ${result.slideCount} slides (${result.byteLength} bytes) and passed package inspection.`;
    } else {
      status.textContent = "PPTX failed package inspection. The build report was downloaded; the PPTX was withheld.";
    }
  } catch (error) {
    currentExportStatus = "failed";
    status.textContent = `PPTX generation failed: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    downloadButton.disabled = findings.some((item) => item.severity === "error");
  }
}

transferButton.addEventListener("click", async () => {
  transferError.textContent = "";
  transferButton.disabled = true;
  try {
    const result = await receiveTransferChunk(transferInput.value, session);
    recordTransfer(result);
    transferInput.value = "";
    if (result.session) {
      const changed = changedSlides(session, result.session);
      session = result.session;
      location.hash = encodeURIComponent(session.id);
      await renderSession(session, changed);
    } else if (result.completedAssetId && session) {
      await renderSession(session);
      status.textContent = `Asset ${result.completedAssetId} completed and the preview was refreshed.`;
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    transferError.textContent = message;
    await refreshStoredTransfers().catch(() => undefined);
  } finally {
    transferButton.disabled = false;
  }
});

previousButton.addEventListener("click", () => selectSlide(currentIndex - 1));
nextButton.addEventListener("click", () => selectSlide(currentIndex + 1));
downloadButton.addEventListener("click", () => void generateAndDownload());
document.addEventListener("keydown", (event) => {
  if (event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLInputElement) return;
  if (event.key === "ArrowLeft") selectSlide(currentIndex - 1);
  if (event.key === "ArrowRight") selectSlide(currentIndex + 1);
});

const requestedSession = decodeURIComponent(location.hash.slice(1));
Promise.all([
  requestedSession ? loadSession(requestedSession) : Promise.resolve(undefined),
  listTransfers(),
]).then(async ([stored, storedTransfers]) => {
  transfers = storedTransfers.map(storedTransferProgress);
  renderTransferProgress();
  if (!stored) return;
  session = stored;
  await renderSession(stored);
}).catch(() => { persisted = false; });

window.addEventListener("beforeunload", revokeObjectUrls);

Object.defineProperty(globalThis, "__pptkitPreviewBridge", {
  configurable: false,
  enumerable: true,
  writable: false,
  value: Object.freeze({
    protocol: PPTKIT_TRANSFER_PROTOCOL,
    maxChunkBytes: MAX_TRANSFER_CHUNK_BYTES,
    getState: () => ({ sessionId: session?.id, transfers: transfers.map((item) => ({ ...item, received: [...item.received], missing: [...item.missing] })) }),
  }),
});

Object.defineProperty(globalThis, "__pptkitPreviewState", {
  get: () => ({ sessionId: session?.id, slideCount: preview?.slides.length ?? 0, findings, exportStatus: currentExportStatus, packageStatus: currentExportStatus === "not-run" ? NOT_RUN_PACKAGE_CHECK.status : "checked", transfers }),
});
