import type { CustomDocumentProperty, NormalizedElement, NormalizedPresentation, NormalizedTextParagraph } from "@pptkit/core";
import type { XmlTextEdit } from "../preservation/xml-text-source.js";
import { PptxImportError, type ImportDiagnostic, type PptxSaveChange, type PptxSessionLimits } from "../types/import.js";
import type { ChartValueBinding, ChartValueProposal } from "./chart-value-bindings.js";
import type { ChartLabelBinding, ChartLabelEditProposal } from "./chart-label-bindings.js";
import type { ChartSeriesOrderProposal } from "./chart-series-order.js";
import type { TextRunBinding } from "./text-run-bindings.js";
import type { DrawingDescriptionBinding } from "./drawing-description-binding.js";
import { sourceIntegerAccepts, sourceRealAccepts } from "./project-custom-properties.js";
import { formatExactCurrency } from "../metadata/currency-text.js";

export interface TextEditGuard {
  readonly part?: string;
  readonly sourceObjectId?: string;
  readonly diagnostic: ImportDiagnostic;
}
export interface BoundXmlTextEdit extends XmlTextEdit {
  readonly rawRunElementIndex?: number;
  readonly rawParagraphElementIndex?: number;
  readonly rawMetadataRootElementIndex?: number;
  readonly rawExtendedRootNamespace?: string;
  readonly rawCustomPropertyElementIndex?: number;
  readonly rawCustomRootNamespace?: string;
}
export interface TextEditAnalysis {
  readonly changes: readonly PptxSaveChange[];
  readonly diagnostics: readonly ImportDiagnostic[];
  readonly parts: ReadonlyMap<string, readonly BoundXmlTextEdit[]>;
  readonly chartValues: readonly ChartValueProposal[];
  readonly chartLabels?: readonly ChartLabelEditProposal[];
  /** Internal candidate only; no public profile authorizes order changes yet. */
  readonly chartOrders?: readonly ChartSeriesOrderProposal[];
  readonly firstSlideNumber?: { readonly part: string; readonly elementIndex: number; readonly namespace: string; readonly expected?: string; readonly value?: number };
  readonly descriptions?: readonly { readonly binding: DrawingDescriptionBinding; readonly value?: string }[];
  readonly customNames?: readonly { readonly binding: CustomPropertyBinding; readonly value: string }[];
  readonly customRemovals?: readonly CustomPropertyBinding[];
  readonly customInsertions?: readonly { readonly source: CustomPropertyInsertionSource; readonly id: number; readonly property: CustomDocumentProperty; readonly index: number }[];
  readonly metadataInsertions?: readonly { readonly binding: CoreMetadataBinding; readonly value: string }[];
  readonly metadataRemovals?: readonly CoreMetadataBinding[];
  readonly slideVisibility?: readonly { readonly binding: SlideVisibilityBinding; readonly hidden: boolean }[];
}
export interface CoreMetadataBinding { readonly field: "title" | "author" | "subject" | "description" | "language" | "revision" | "keywords" | "company" | "createdAt" | "modifiedAt"; readonly part: string; readonly elementIndex: number; readonly rootElementIndex: number; readonly namespace: string; readonly local: string; readonly baseline: string | undefined; readonly lexical?: string; readonly insert?: true; readonly create?: { readonly dialect: "strict" | "transitional"; readonly relationshipId: string }; readonly extendedRootNamespace?: string; readonly removable?: true }
export interface CustomPropertyInsertionSource { readonly part: string; readonly rootNamespace: string; readonly valueNamespace: string; readonly usedIds: ReadonlySet<number>; readonly reservedNames: ReadonlySet<string>; readonly create?: { readonly dialect: "strict" | "transitional"; readonly relationshipId: string } }
export interface CustomPropertyBinding { readonly index: number; readonly part: string; readonly propertyElementIndex: number; readonly valueElementIndex: number; readonly rootNamespace: string; readonly valueNamespace: string; readonly local: string; readonly baseline: string | number | boolean; readonly lexical: string; readonly baselineName?: string; readonly lexicalName?: string; readonly reservedNames?: ReadonlySet<string>; readonly sourceObjectId: string }
export interface SlideVisibilityBinding { readonly slideId: string; readonly part: string; readonly elementIndex: number; readonly namespace: string; readonly expected?: string; readonly baseline: boolean }
export interface FirstSlideNumberBinding { readonly part: string; readonly elementIndex: number; readonly namespace: string; readonly expected?: string; readonly baseline?: number; readonly allowPresenceChange: boolean }
const runIdentity = (slideId: string, elementId: string | undefined, p: number, r: number, cell?: { rowIndex: number; cellIndex: number }): string => JSON.stringify([slideId, elementId, p, r, ...(cell ? [cell.rowIndex, cell.cellIndex] : [])]);
const identity = (binding: TextRunBinding): string => runIdentity(binding.slideId, binding.elementId, binding.paragraphIndex, binding.runIndex, binding.owner === "table-cell" ? binding : undefined);
const sourceKey = (binding: TextRunBinding): string => JSON.stringify([binding.source.part, binding.sourceElementIndex]);
const token = (value: string): string => value.replaceAll("~", "~0").replaceAll("/", "~1");

/** Pure snapshot analysis; no package access, mutation, hashing or writer calls. */
export async function analyzeTextEdits(
  baseline: NormalizedPresentation, candidate: NormalizedPresentation,
  bindings: readonly TextRunBinding[], guards: readonly TextEditGuard[],
  limits: Readonly<PptxSessionLimits>, checkpoint: () => Promise<void>, chartBindings: readonly ChartValueBinding[] = [], firstNumberBinding?: FirstSlideNumberBinding, slideVisibilityBindings: readonly SlideVisibilityBinding[] = [], metadataBindings: readonly CoreMetadataBinding[] = [], customPropertyBindings: readonly CustomPropertyBinding[] = [], chartLabelBindings: readonly ChartLabelBinding[] = [],
): Promise<TextEditAnalysis> {
  const maxWork = limits.maxSaveAnalysisWork!, maxChanges = limits.maxSaveChanges!;
  let work = 0;
  const charge = async (units: number) => {
    work += units;
    if (work > maxWork) throw new PptxImportError("save-limit-exceeded", "Save analysis exceeds maxSaveAnalysisWork.");
    await checkpoint();
  };
  const chartOwners = new Map<string, ChartValueBinding[]>(), chartPaths = new Map<string, ChartValueBinding>();
  for (const binding of chartBindings) { await charge(1); const key = JSON.stringify([binding.slideId, binding.elementId]), list = chartOwners.get(key) ?? []; list.push(binding); chartOwners.set(key, list); }
  const chartValues: ChartValueProposal[] = [];
  const labelOwners = new Map<string, ChartLabelBinding[]>(), labelPaths = new Map<string, ChartLabelBinding>(), chartLabels: ChartLabelEditProposal[] = [];
  for (const binding of chartLabelBindings) { await charge(1); const key = JSON.stringify([binding.slideId, binding.elementId]), list = labelOwners.get(key) ?? []; list.push(binding); labelOwners.set(key, list); }
  const byIdentity = new Map(bindings.map(binding => [identity(binding), binding]));
  const tableOwners = new Set(bindings.filter(binding => binding.owner === "table-cell").map(binding => JSON.stringify([binding.slideId, binding.elementId])));
  const runPaths = new Map<string, TextRunBinding>(), owners = new Map<string, { slideId: string; elementId?: string }>(), textOwners = new Set<string>();
  const recordRuns = async (content: readonly NormalizedTextParagraph[], path: string, slideId: string, elementId?: string, cell?: { rowIndex: number; cellIndex: number }) => {
    for (let p = 0; p < content.length; p++) for (let r = 0; r < content[p]!.runs.length; r++) {
      const runPath = `${path}/${p}/runs/${r}/text`; await charge(runPath.length + 1);
      const binding = byIdentity.get(runIdentity(slideId, elementId, p, r, cell));
      if (binding) runPaths.set(runPath, binding);
    }
  };
  const elements: { element: NormalizedElement; path: string; slideId: string }[] = [];
  for (let i = 0; i < baseline.slides.length; i++) {
    const slide = baseline.slides[i]!;
    owners.set(`/slides/${i}/notes`, { slideId: slide.id });
    await recordRuns(slide.notes, `/slides/${i}/notes`, slide.id);
    for (let j = 0; j < slide.elements.length; j++) elements.push({ element: slide.elements[j]!, path: `/slides/${i}/elements/${j}`, slideId: slide.id });
  }
  while (elements.length) {
    const { element, path, slideId } = elements.pop()!; await charge(path.length + 1);
    owners.set(path, { slideId, elementId: element.id });
    if (element.type === "chart") for (const binding of chartOwners.get(JSON.stringify([slideId, element.id])) ?? []) {
      const field = binding.coordinate === "x" ? "xValues" : binding.coordinate === "y" ? "yValues" : binding.coordinate === "size" ? "sizes" : "values";
      const valuePath = `${path}/series/${binding.seriesIndex}/${field}/${binding.pointIndex}`; await charge(valuePath.length + 1); chartPaths.set(valuePath, binding);
    }
    if (element.type === "chart") for (const binding of labelOwners.get(JSON.stringify([slideId, element.id])) ?? []) {
      const labelPath = binding.seriesIndex === undefined ? `${path}/categories/${binding.categoryIndex}` : `${path}/series/${binding.seriesIndex}/name`; await charge(labelPath.length + 1);
      if (labelPaths.has(labelPath)) throw new PptxImportError("save-candidate-invalid", "Duplicate chart label bindings have the same Core identity.");
      labelPaths.set(labelPath, binding);
    }
    if (element.type === "group") for (let i = 0; i < element.children.length; i++) elements.push({ element: element.children[i]!, path: `${path}/children/${i}`, slideId });
    if (element.type === "table" && tableOwners.has(JSON.stringify([slideId, element.id]))) {
      for (let rowIndex = 0; rowIndex < element.rows.length; rowIndex++) {
        const row = element.rows[rowIndex]!;
        for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
          await recordRuns(row.cells[cellIndex]!.content, `${path}/rows/${rowIndex}/cells/${cellIndex}/content`, slideId, element.id, { rowIndex, cellIndex });
        }
      }
    }
    const text = element.type === "text" ? element : element.type === "shape" ? element.text : undefined;
    if (!text) continue;
    const textPath = element.type === "text" ? path : `${path}/text`;
    textOwners.add(textPath);
    await recordRuns(text.content, `${textPath}/content`, slideId, element.id);
  }
  const metadataPaths = new Map<string, CoreMetadataBinding>();
  for (const binding of metadataBindings) { await charge(binding.part.length + binding.field.length + 1); metadataPaths.set(`/metadata/${binding.field}${binding.field === "keywords" ? "/0" : ""}`, binding); }
  const metadataEdits: { binding: CoreMetadataBinding; value: string }[] = [];
  const metadataRemovals: CoreMetadataBinding[] = [];
  const customPaths = new Map<string, CustomPropertyBinding>();
  for (const binding of customPropertyBindings) { await charge(binding.part.length + binding.lexical.length + 1); customPaths.set(`/metadata/customProperties/${binding.index}/value`, binding); }
  const customEdits: { binding: CustomPropertyBinding; value: string }[] = [];
  const customNamePaths = new Map<string, CustomPropertyBinding>();
  for (const binding of customPropertyBindings) if (binding.baselineName !== undefined) { await charge(binding.baselineName.length + 1); customNamePaths.set(`/metadata/customProperties/${binding.index}/name`, binding); }
  const customNames: { binding: CustomPropertyBinding; value: string }[] = [];
  const visibilityPaths = new Map<string, SlideVisibilityBinding>();
  const visibilityById = new Map<string, SlideVisibilityBinding>();
  for (const binding of slideVisibilityBindings) { await charge(binding.slideId.length + binding.part.length + 1); visibilityById.set(binding.slideId, binding); }
  for (const [index, slide] of baseline.slides.entries()) { const binding = visibilityById.get(slide.id); if (binding) visibilityPaths.set(`/slides/${index}/hidden`, binding); }
  const slideVisibility: { binding: SlideVisibilityBinding; hidden: boolean }[] = [];
  const changes: PptxSaveChange[] = [], diagnostics: ImportDiagnostic[] = [], proposals = new Map<string, string>();
  let firstSlideNumber: TextEditAnalysis["firstSlideNumber"];
  const values: { left: unknown; right: unknown; path: string; owner?: { slideId: string; elementId?: string } }[] = [{ left: baseline, right: candidate, path: "" }];
  const addDiagnostic = (diagnostic: ImportDiagnostic) => {
    if (diagnostics.length >= maxChanges) throw new PptxImportError("save-limit-exceeded", "Save diagnostics exceed maxSaveChanges.");
    diagnostics.push(diagnostic);
  };
  const emit = (left: unknown, right: unknown, path: string, owner: { slideId: string; elementId?: string } | undefined, arrayStructure = false) => {
    if (changes.length >= maxChanges) throw new PptxImportError("save-limit-exceeded", "Save changes exceed maxSaveChanges.");
    const chartBinding = chartPaths.get(path);
    const labelBinding = labelPaths.get(path);
    const visibilityBinding = visibilityPaths.get(path);
    const metadataBinding = metadataPaths.get(path);
    const customBinding = customPaths.get(path), customNameBinding = customNamePaths.get(path);
    const binding = runPaths.get(path), strings = typeof left === "string" && typeof right === "string";
    const kind: PptxSaveChange["kind"] = /^\/assets(?:\/|$)/.test(path) ? "asset" : arrayStructure ? "structure" : /\/customData(?:\/|$)/.test(path) ? "other" : binding || /\/(?:content|notes)(?:\/.*)?\/text$/.test(path) ? "text"
      : /^\/assets(?:\/|$)/.test(path) || /\/assetId$/.test(path) ? "asset"
      : /\/(?:box|transform|coordinateSize|size)(?:\/|$)/.test(path) || /\/(?:columns\/\d+|rows\/\d+\/height)$/.test(path) ? "geometry"
      : /^\/metadata(?:\/|$)/.test(path) ? "metadata"
      : /\/(?:style|frame|textStylePresets|theme)(?:\/|$)/.test(path) ? "style"
      : /\/(?:id|layoutId|type|hidden|section|tags|elements|children|slides|layouts|rowSpan|colSpan)$/.test(path) ? "structure" : "other";
    const change: PptxSaveChange = {
      kind, documentPath: path, ...owner, ...(binding ? { source: binding.source } : {}),
      ...(path === "/firstSlideNumber" && firstNumberBinding ? { source: { part: firstNumberBinding.part, elementIndex: firstNumberBinding.elementIndex, path: "presentation/@firstSlideNum" } } : {}),
      ...(path === "/firstSlideNumber" && firstNumberBinding && (left === undefined || typeof left === "number") && (right === undefined || typeof right === "number") ? { firstSlideNumber: { ...(left !== undefined ? { before: left } : {}), ...(right !== undefined ? { after: right } : {}) } } : {}),
      ...(customNameBinding ? { source: { part: customNameBinding.part, elementIndex: customNameBinding.propertyElementIndex, path: "Properties/property/@name", sourceObjectId: customNameBinding.sourceObjectId } } : {}),
      ...(customBinding ? { source: { part: customBinding.part, elementIndex: customBinding.valueElementIndex, path: "Properties/property/value", sourceObjectId: customBinding.sourceObjectId } } : {}),
      ...(metadataBinding ? { source: { part: metadataBinding.part, ...(!metadataBinding.insert ? { elementIndex: metadataBinding.elementIndex } : {}), path: `${metadataBinding.extendedRootNamespace ? "Properties" : "coreProperties"}/${metadataBinding.local}` }, ...(typeof left === "string" && typeof right === "string" ? { metadata: { before: left, after: right } } : {}) } : {}),
      ...(visibilityBinding ? { slideId: visibilityBinding.slideId, source: { part: visibilityBinding.part, elementIndex: visibilityBinding.elementIndex, path: "sld/@show" } } : {}),
      ...(chartBinding ? { source: chartBinding.source, ...(typeof left === "number" && typeof right === "number" ? { chartValue: { before: left, after: right, seriesIndex: chartBinding.seriesIndex, pointIndex: chartBinding.pointIndex, ...(chartBinding.coordinate ? { coordinate: chartBinding.coordinate } : {}) } } : {}) } : {}),
      ...(labelBinding && typeof left === "string" && typeof right === "string" ? labelBinding.seriesIndex === undefined ? { chartLabel: { before: left, after: right, categoryIndex: labelBinding.categoryIndex } } : { chartSeriesName: { before: left, after: right, seriesIndex: labelBinding.seriesIndex } } : {}),
      ...(kind === "text" && strings ? { text: { before: left, after: right } } : {}),
    };
    changes.push(change);
    if (path === "/firstSlideNumber" && firstNumberBinding && Object.is(left, firstNumberBinding.baseline) && ((right === undefined && firstNumberBinding.allowPresenceChange) || (typeof right === "number" && Number.isSafeInteger(right) && right >= 0 && right <= 9999))) firstSlideNumber = { ...firstNumberBinding, ...(right !== undefined ? { value: right } : {}) };
    else if (metadataBinding && typeof right === "string" && left === metadataBinding.baseline && (typeof left === "string" || metadataBinding.insert && left === undefined)) metadataEdits.push({ binding: metadataBinding, value: right });
    else if (metadataBinding?.removable && typeof left === "string" && left === metadataBinding.baseline && right === undefined) metadataRemovals.push(metadataBinding);
    else if (metadataBinding?.field === "revision" && typeof left === "number" && String(left) === metadataBinding.baseline && typeof right === "number" && Number.isSafeInteger(right) && right >= 0 && !Object.is(right, -0)) metadataEdits.push({ binding: metadataBinding, value: String(right) });
    else if (customNameBinding && left === customNameBinding.baselineName && typeof right === "string" && right.length > 0) {
      if (customNameBinding.reservedNames?.has(right)) addDiagnostic({ code: "save-edit-unsupported", severity: "error", documentPath: path, message: "Custom property name conflicts with an unprojected source property." });
      else if (/_x[0-9a-fA-F]{4}_x[0-9a-fA-F]{4}_/.test(right)) addDiagnostic({ code: "save-edit-unsupported", severity: "error", documentPath: path, message: "Overlapping custom-property escape-like sequences are unsupported." });
      else customNames.push({ binding: customNameBinding, value: right.replace(/_x[0-9a-fA-F]{4}_/g, match => `_x005F_${match.slice(1)}`) });
    }
    else if (customBinding && Object.is(left, customBinding.baseline) && typeof right === typeof left && (typeof right === "string" || typeof right === "boolean" || typeof right === "number" && Number.isFinite(right))) {
      const currencyText = customBinding.local === "cy" && typeof right === "number" ? formatExactCurrency(right) : undefined;
      const text = currencyText ?? (typeof right === "string" ? right.replace(/_x[0-9a-fA-F]{4}_/g, match => `_x005F_${match.slice(1)}`) : typeof right === "number" && Object.is(right, -0) ? "-0" : String(right));
      if (customBinding.local === "cy" && currencyText === undefined) addDiagnostic({ code: "save-edit-unsupported", severity: "error", documentPath: path, message: "Currency must be a nonnegative binary64 value exactly representable with four decimal places." });
      else if (typeof right === "number" && sourceIntegerAccepts(customBinding.local, right) === false) addDiagnostic({ code: "save-edit-unsupported", severity: "error", documentPath: path, message: "Custom integer exceeds its original source variant's supported range." });
      else if (typeof right === "number" && sourceRealAccepts(customBinding.local, right) === false) addDiagnostic({ code: "save-edit-unsupported", severity: "error", documentPath: path, message: "Custom real value cannot be represented exactly in its original source variant." });
      else if (typeof right === "string" && /_x[0-9a-fA-F]{4}_x[0-9a-fA-F]{4}_/.test(right)) addDiagnostic({ code: "save-edit-unsupported", severity: "error", documentPath: path, message: "Overlapping custom-property escape-like sequences are unsupported." });
      else customEdits.push({ binding: customBinding, value: text });
    }
    else if (visibilityBinding && typeof left === "boolean" && typeof right === "boolean" && left === visibilityBinding.baseline) slideVisibility.push({ binding: visibilityBinding, hidden: right });
    else if (chartBinding && typeof left === "number" && typeof right === "number" && Number.isFinite(right) && (chartBinding.coordinate !== "size" || right > 0) && chartBinding.status === "bound" && Object.is(left, chartBinding.baselineValue)) chartValues.push({ binding: chartBinding, value: right });
    else if (labelBinding && typeof left === "string" && typeof right === "string" && right.length > 0 && left === labelBinding.baselineValue && labelBinding.closure.status === "bound") {
      chartLabels.push({ binding: labelBinding, value: right });
    }
    else if (binding && strings && binding.status === "bound" && binding.sourceElementIndex && left === binding.baselineText) proposals.set(path, right);
    else addDiagnostic({
      code: "save-edit-unsupported", severity: "error", documentPath: path, ...owner, ...(binding ? { source: binding.source } : {}),
      message: labelBinding?.closure.reason ?? chartBinding?.reason ?? binding?.reason ?? (kind === "metadata" ? "The selected preservation profiles cannot represent this metadata change or its source is unavailable." : "The selected text-run profile admits only bound scalar replacements with unchanged document topology, styles, geometry and assets."),
    });
  };
  while (values.length) {
    const { left, right, path, owner: inherited } = values.pop()!; await charge(path.length + 1);
    if (Object.is(left, right)) continue;
    const owner = owners.get(path) ?? inherited;
    if (!left || !right || typeof left !== "object" || typeof right !== "object" || Array.isArray(left) !== Array.isArray(right)) { emit(left, right, path, owner); continue; }
    const l = left as Record<string, unknown>, r = right as Record<string, unknown>;
    let keys = Object.keys(l).sort(), rightKeys = Object.keys(r).sort();
    if (path === "" && firstNumberBinding?.allowPresenceChange) {
      values.push({ left: l.firstSlideNumber, right: r.firstSlideNumber, path: "/firstSlideNumber" });
      keys = keys.filter(key => key !== "firstSlideNumber"); rightKeys = rightKeys.filter(key => key !== "firstSlideNumber");
    }
    if (path === "/metadata") {
      for (const binding of metadataBindings) if (binding.insert || binding.removable && (r[binding.field] === undefined || binding.field === "keywords" && Array.isArray(r.keywords) && r.keywords.length === 0)) {
        if (binding.field === "keywords") {
          if (binding.insert && Array.isArray(l.keywords) && l.keywords.length === 0 && Array.isArray(r.keywords) && r.keywords.length === 1) values.push({ left: undefined, right: r.keywords[0], path: "/metadata/keywords/0" });
          else if (binding.removable && Array.isArray(l.keywords) && l.keywords.length === 1 && (r.keywords === undefined || Array.isArray(r.keywords) && r.keywords.length === 0)) values.push({ left: l.keywords[0], right: undefined, path: "/metadata/keywords/0" });
          else continue;
        } else values.push({ left: l[binding.field], right: r[binding.field], path: `/metadata/${binding.field}` });
        keys = keys.filter(key => key !== binding.field); rightKeys = rightKeys.filter(key => key !== binding.field);
      }
    }
    if ((Array.isArray(left) && left.length !== (right as unknown[]).length) || keys.length !== rightKeys.length || keys.some((key, i) => key !== rightKeys[i])) { emit(left, right, path, owner, Array.isArray(left)); continue; }
    for (let i = keys.length - 1; i >= 0; i--) {
      const key = keys[i]!;
      // Only the actual normalized text owner has this derived field exemption.
      if (key === "plainText" && textOwners.has(path)) continue;
      values.push({ left: l[key], right: r[key], path: `${path}/${token(key)}`, ...(owner ? { owner } : {}) });
    }
  }
  const targets = new Map<string, { text: string; binding: TextRunBinding; changed: boolean }>();
  for (const [path, binding] of runPaths) {
    await charge(path.length + 1);
    if (!binding.sourceElementIndex) continue;
    const key = sourceKey(binding), text = proposals.get(path) ?? binding.baselineText, changed = proposals.has(path), existing = targets.get(key);
    if (existing && existing.text !== text) {
      addDiagnostic({ code: "save-source-shared-conflict", severity: "error", source: binding.source, slideId: binding.slideId, ...(binding.elementId ? { elementId: binding.elementId } : {}), message: "Multiple Core views of one raw source text element propose different text; all views must agree." });
    } else targets.set(key, { text, binding, changed: changed || Boolean(existing?.changed) });
  }
  const guardKey = (part?: string, object?: string) => JSON.stringify([part ?? null, object ?? null]);
  const guardIndex = new Map<string, TextEditGuard[]>(), visibilityGuardIndex = new Map<string | undefined, TextEditGuard[]>();
  for (const guard of guards) {
    const key = guardKey(guard.part, guard.sourceObjectId); await charge(key.length + 1);
    let list = guardIndex.get(key); if (!list) { list = []; guardIndex.set(key, list); } list.push(guard);
    let visibilityGuards = visibilityGuardIndex.get(guard.part); if (!visibilityGuards) { visibilityGuards = []; visibilityGuardIndex.set(guard.part, visibilityGuards); } visibilityGuards.push(guard);
  }
  const parts = new Map<string, BoundXmlTextEdit[]>(), reportedGuards = new Set<TextEditGuard>();
  for (const { binding } of [...slideVisibility, ...metadataEdits, ...metadataRemovals.map(binding => ({ binding })), ...customEdits, ...customNames]) {
    for (const guard of [...(visibilityGuardIndex.get(undefined) ?? []), ...(visibilityGuardIndex.get(binding.part) ?? [])]) {
      await charge(1);
      if (!reportedGuards.has(guard)) { reportedGuards.add(guard); addDiagnostic(guard.diagnostic); }
    }
  }
  for (const { text, binding, changed } of targets.values()) {
    if (!changed) continue;
    const keys = new Set([guardKey(), guardKey(binding.source.part), guardKey(binding.source.part, binding.source.sourceObjectId), guardKey(undefined, binding.source.sourceObjectId)]);
    for (const key of keys) for (const guard of guardIndex.get(key) ?? []) {
      await charge(1);
      if (!reportedGuards.has(guard)) {
        reportedGuards.add(guard); addDiagnostic(guard.diagnostic);
      }
    }
    let edits = parts.get(binding.source.part);
    if (!edits) { edits = []; parts.set(binding.source.part, edits); }
    edits.push({ elementIndex: binding.sourceElementIndex!, namespaceUri: binding.namespaceUri, local: "t", expectedText: binding.baselineText, text,
      ...(binding.rawRunElementIndex !== undefined ? { rawRunElementIndex: binding.rawRunElementIndex } : {}),
      ...(binding.rawParagraphElementIndex !== undefined ? { rawParagraphElementIndex: binding.rawParagraphElementIndex } : {}) });
  }
  for (const { binding, value } of metadataEdits) {
    await charge(binding.part.length + (binding.baseline?.length ?? 0) + value.length + 1);
    if (binding.insert) continue;
    let edits = parts.get(binding.part); if (!edits) { edits = []; parts.set(binding.part, edits); }
    edits.push({ elementIndex: binding.elementIndex, namespaceUri: binding.namespace, local: binding.local, expectedText: binding.lexical ?? binding.baseline!, text: value, rawMetadataRootElementIndex: binding.rootElementIndex, ...(binding.extendedRootNamespace ? { rawExtendedRootNamespace: binding.extendedRootNamespace } : {}) });
  }
  for (const { binding, value } of customEdits) {
    await charge(binding.part.length + binding.lexical.length + value.length + 1);
    let edits = parts.get(binding.part); if (!edits) { edits = []; parts.set(binding.part, edits); }
    edits.push({ elementIndex: binding.valueElementIndex, namespaceUri: binding.valueNamespace, local: binding.local, expectedText: binding.lexical, text: value, rawCustomPropertyElementIndex: binding.propertyElementIndex, rawCustomRootNamespace: binding.rootNamespace });
  }
  changes.sort((a, b) => a.documentPath < b.documentPath ? -1 : a.documentPath > b.documentPath ? 1 : 0);
  const visibilityByPart = new Map<string, boolean>();
  for (const [index, old] of baseline.slides.entries()) {
    const binding = visibilityById.get(old.id), current = candidate.slides[index];
    if (!binding || !current || current.id !== old.id) continue;
    await charge(binding.part.length + 1);
    if (visibilityByPart.has(binding.part) && visibilityByPart.get(binding.part) !== current.hidden) addDiagnostic({ code: "save-source-shared-conflict", severity: "error", source: { part: binding.part, elementIndex: binding.elementIndex }, message: "Views of one source slide propose different visibility." });
    else visibilityByPart.set(binding.part, current.hidden);
  }
  const metadataInsertions = metadataEdits.filter(item => item.binding.insert);
  return { changes, diagnostics, parts, chartValues, ...(chartLabels.length ? { chartLabels } : {}), ...(metadataInsertions.length ? { metadataInsertions } : {}), ...(metadataRemovals.length ? { metadataRemovals } : {}), ...(customNames.length ? { customNames } : {}), ...(slideVisibility.length ? { slideVisibility } : {}), ...(firstSlideNumber ? { firstSlideNumber } : {}) };
}
