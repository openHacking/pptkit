import { PresentationValidationError, type NormalizedPresentation, type ManagedPresentationDocument } from "@pptkit/core";
import { PptxImportError, type ImportedPictureSource, type ImportedPresentation, type ImportDiagnostic, type ImportReport, type PptxSaveMode, type PptxSavePlan, type PptxSavePreparation, type PptxSaveReport, type PptxSessionLimits } from "../types/import.js";
import { equalState, freezeReport, notSerializable, snapshot } from "./state.js";
import { analyzeTextEdits, type FirstSlideNumberBinding, type SlideVisibilityBinding, type CoreMetadataBinding, type CustomPropertyBinding, type CustomPropertyInsertionSource, type TextEditGuard } from "./analyze-text-edits.js";
import { analyzeDescriptionEdit } from './analyze-description-edit.js';
import type { DrawingDescriptionBinding } from './drawing-description-binding.js';
import { prepareTextCandidate } from "./prepare-text-candidate.js";
import { analyzePictureEdit } from "./analyze-picture-edit.js";
import { PictureSourceUnsupportedError, preparePictureCandidate } from "./prepare-picture-candidate.js";
import { XmlSourceEditError } from "../preservation/xml-text-source.js";
import { PackageInspectionError } from "../types/archive.js";
import type { ChartValueBinding } from "./chart-value-bindings.js";
import type { TextRunBinding } from "./text-run-bindings.js";
import type { PptxInspectionOptions } from "../types/structure.js";
import { analyzeChartOrders, captureChartSeriesOrderRequests } from "./analyze-chart-order.js";
import type { ChartSeriesOrderBinding } from "./chart-series-order.js";

interface Slot { bytes?: Uint8Array; snapshot?: NormalizedPresentation; disposed: boolean; size: number }
interface EditState { readonly decimalSave?: boolean; readonly currencySave?: boolean; readonly allowCoreRemovals?: boolean; readonly allowCoreStringRemovals?: boolean; readonly allowCoreInsertions?: boolean; readonly allowCustomInsertions?: boolean; readonly customInsertionSource?: CustomPropertyInsertionSource; readonly allowCustomRemovals?: boolean; readonly customPropertyBindings?: readonly CustomPropertyBinding[]; readonly metadataBindings?: readonly CoreMetadataBinding[]; readonly slideVisibilityBindings?: readonly SlideVisibilityBinding[]; readonly descriptionBindings?: readonly DrawingDescriptionBinding[]; readonly descriptionBatch?: boolean; readonly descriptionMixed?: boolean; readonly firstNumberBinding?: FirstSlideNumberBinding; readonly pictureSources?: readonly ImportedPictureSource[]; readonly chartBindings?: readonly ChartValueBinding[]; readonly chartLabelBindings?: readonly import("./chart-label-bindings.js").ChartLabelBinding[]; readonly chartOrderBindings?: readonly ChartSeriesOrderBinding[]; readonly bindings: readonly TextRunBinding[]; readonly guards: readonly TextEditGuard[]; readonly options: PptxInspectionOptions; readonly diagramPreview?: "cache"; readonly textFontProfile?: "script-families-v1" | "script-families-v2" | "script-families-v3"; readonly notesTextProfile?: "master-style-v1"; readonly fieldTextProfile?: "slide-number-v1" | "date-and-number-v1" | "date-time-and-number-v1" }
export function createSession(document: ManagedPresentationDocument, report: ImportReport, source: Uint8Array | undefined, limits: Readonly<Required<PptxSessionLimits>>, editState?: EditState): ImportedPresentation {
  const currentSnapshot = () => snapshot(document, limits.maxSnapshotValues);
  const sameState = (a: unknown, b: unknown) => equalState(a, b, limits.maxSnapshotValues);
  let baseline: NormalizedPresentation | undefined = currentSnapshot(), disposed = false, retainedBytes = 0;
  const ownership = new WeakMap<object, Slot>(), active = new Set<Slot>(), operations = new Set<AbortController>();
  const abortLinks = new Map<AbortSignal, { controllers: Set<AbortController>; abort: () => void }>();
  const subscribeAbort = (signal: AbortSignal | undefined, controller: AbortController): (() => void) => {
    if (!signal) return () => {};
    let link = abortLinks.get(signal);
    if (!link) {
      const controllers = new Set<AbortController>();
      link = { controllers, abort: () => { for (const target of controllers) target.abort(); } };
      signal.addEventListener("abort", link.abort, { once: true }); abortLinks.set(signal, link);
    }
    link.controllers.add(controller);
    if (signal.aborted) controller.abort();
    const owned = link;
    return () => {
      owned.controllers.delete(controller);
      if (!owned.controllers.size) { signal.removeEventListener("abort", owned.abort); abortLinks.delete(signal); }
    };
  };
  const live = () => { if (disposed) throw new PptxImportError("session-disposed", "The import session has been disposed."); };
  const abort = (signal?: AbortSignal) => { live(); if (signal?.aborted) throw new PptxImportError("operation-aborted", "Save preparation was aborted."); };
  const yieldToCaller = async (signal?: AbortSignal) => { await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort(signal); };
  const release = (slot: Slot) => {
    if (active.delete(slot)) retainedBytes -= slot.size;
    slot.disposed = true; delete slot.bytes; delete slot.snapshot;
  };
  const blocked = (mode: PptxSaveMode, code: string, message: string, details: readonly ImportDiagnostic[] = []): PptxSavePreparation => freezeReport({ status: "blocked", report: { mode, assurance: "none", losses: [], diagnostics: [{ code, severity: "error", message }, ...details] } });
  const invalidDocument = (mode: PptxSaveMode, error: PptxImportError): PptxSavePreparation => {
    const details = error.cause instanceof PresentationValidationError ? error.cause.diagnostics.map((d): ImportDiagnostic => ({ code: d.code, severity: d.severity, message: d.message, documentPath: d.path, ...(d.slideId ? { slideId: d.slideId } : {}), ...(d.elementId ? { elementId: d.elementId } : {}) })) : [];
    return blocked(mode, error.code, error.message, details);
  };
  const currentPlan = (plan: PptxSavePlan, options: { readonly acceptLossIds?: readonly string[] } = {}): Slot => {
    live();
    const slot = plan && typeof plan === "object" ? ownership.get(plan) : undefined;
    if (!slot) throw new PptxImportError("plan-session-mismatch", "Save plan was not prepared by this session.");
    if (slot.disposed) throw new PptxImportError("plan-disposed", "The save plan has been disposed.");
    if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some((key) => key !== "acceptLossIds")) throw new PptxImportError("save-loss-consent-mismatch", "Consent must use the exact plan loss IDs, not a generic override.");
    const accepted = options.acceptLossIds;
    if (accepted !== undefined && (!Array.isArray(accepted) || accepted.length !== 0)) throw new PptxImportError("save-loss-consent-mismatch", "This loss-free plan accepts no loss IDs.");
    // Caller getters may synchronously edit or dispose the session/plan.
    live();
    if (slot.disposed) throw new PptxImportError("plan-disposed", "The save plan has been disposed.");
    if (!sameState(slot.snapshot, currentSnapshot())) throw new PptxImportError("save-snapshot-stale", "The current Core document differs from the prepared snapshot.");
    live();
    if (slot.disposed) throw new PptxImportError("plan-disposed", "The save plan has been disposed.");
    return slot;
  };
  const session: ImportedPresentation & { toJSON(): never } = {
    document, report,
    async prepareSave(options): Promise<PptxSavePreparation> {
      live();
      if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("prepareSave requires an explicit no-op, edit or rebuild mode.");
      const { mode, signal } = options;
      const orderRequests = captureChartSeriesOrderRequests(options.chartSeriesOrders, limits.maxSaveChanges, limits.maxSaveAnalysisWork);
      if (orderRequests.length && (mode !== "edit" || !editState?.chartOrderBindings)) throw new TypeError("Chart series order requires edit mode and existing-series-order-v6.");
      if (orderRequests.length) return blocked(mode, "save-chart-order-native-unsupported", "Series-order saves are disabled: changing c:order alone does not reorder the chart in the observed PowerPoint version.");
      const requestedRemovals = options.customPropertyRemovals;
      if (requestedRemovals !== undefined && (!Array.isArray(requestedRemovals) || Reflect.ownKeys(requestedRemovals).length !== requestedRemovals.length + 1)) throw new TypeError("Custom property removals require a dense array.");
      const removalIds = requestedRemovals ? Array.from(requestedRemovals) : [];
      if (removalIds.some(value => typeof value !== "string" || !value || value.length > 256) || new Set(removalIds).size !== removalIds.length) throw new TypeError("Custom property removals require unique bounded source identities.");
      if (removalIds.length && (mode !== "edit" || !editState?.allowCustomRemovals)) throw new TypeError("Custom property removal requires edit mode and remove-v3.");
      if (removalIds.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Custom property removals exceed maxSaveChanges.");
      const requestedAdditions = options.customPropertyAdditions;
      if (requestedAdditions !== undefined && (!Array.isArray(requestedAdditions) || Reflect.ownKeys(requestedAdditions).length !== requestedAdditions.length + 1)) throw new TypeError("Custom property additions require a dense array.");
      const additionIndexes = requestedAdditions ? Array.from(requestedAdditions) : [];
      if (additionIndexes.some(value => !Number.isSafeInteger(value) || value < 0) || new Set(additionIndexes).size !== additionIndexes.length) throw new TypeError("Custom property additions require unique nonnegative indexes.");
      if (additionIndexes.length && (mode !== "edit" || !editState?.allowCustomInsertions)) throw new TypeError("Custom property insertion requires edit mode and append-and-remove-v4.");
      if (removalIds.length + additionIndexes.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Custom property structural edits exceed maxSaveChanges.");
      const requestedCoreInsertions = options.coreMetadataInsertions;
      if (requestedCoreInsertions !== undefined && (!Array.isArray(requestedCoreInsertions) || Reflect.ownKeys(requestedCoreInsertions).length !== requestedCoreInsertions.length + 1)) throw new TypeError("Core metadata insertions require a dense array.");
      const coreFields: ("title" | "author" | "subject" | "description" | "language" | "revision")[] = requestedCoreInsertions ? Array.from(requestedCoreInsertions) : [];
      if (coreFields.some(field => !["title", "author", "subject", "description", "language", "revision"].includes(field)) || new Set(coreFields).size !== coreFields.length) throw new TypeError("Core metadata insertions require unique supported fields.");
      if (coreFields.length && (mode !== "edit" || !editState?.allowCoreInsertions)) throw new TypeError("Core metadata insertion requires edit mode and a supporting metadata profile.");
      if (removalIds.length + additionIndexes.length + coreFields.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Structural edits exceed maxSaveChanges.");
      const requestedCoreRemovals = options.coreMetadataRemovals;
      if (requestedCoreRemovals !== undefined && (!Array.isArray(requestedCoreRemovals) || Reflect.ownKeys(requestedCoreRemovals).length !== requestedCoreRemovals.length + 1)) throw new TypeError("Core metadata removals require a dense array.");
      const coreRemovals: ("title" | "author" | "language" | "revision")[] = requestedCoreRemovals ? Array.from(requestedCoreRemovals) : [];
      if (coreRemovals.some(field => !["revision", "title", "author", "language"].includes(field)) || new Set(coreRemovals).size !== coreRemovals.length || coreFields.some(field => coreRemovals.includes(field as typeof coreRemovals[number]))) throw new TypeError("Core metadata removals require unique supported fields and cannot overlap insertions.");
      if (coreRemovals.some(field => field !== "revision") && !editState?.allowCoreStringRemovals) throw new TypeError("String metadata removal requires a supporting metadata profile.");
      if (coreRemovals.length && (mode !== "edit" || !editState?.allowCoreRemovals)) throw new TypeError("Core metadata removal requires edit mode and a supporting metadata profile.");
      if (removalIds.length + additionIndexes.length + coreFields.length + coreRemovals.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Structural edits exceed maxSaveChanges.");
      if (!["no-op", "edit", "rebuild"].includes(mode)) throw new TypeError("prepareSave requires an explicit no-op, edit or rebuild mode.");
      abort(signal);
      let candidate: NormalizedPresentation;
      try { candidate = currentSnapshot(); } catch (e) { if (e instanceof PptxImportError && e.code === "save-invalid-document") return invalidDocument(mode, e); throw e; }
      await yieldToCaller(signal);
      try { if (!sameState(candidate, currentSnapshot())) return blocked(mode, "save-snapshot-stale", "The Core document changed while preparing the save."); }
      catch (e) { if (e instanceof PptxImportError && e.code === "save-invalid-document") return invalidDocument(mode, e); throw e; }
      abort(signal);
      if (mode !== "no-op" && !(mode === "edit" && editState)) return blocked(mode, "save-mode-unsupported", "Edit requires an explicit content preservation profile; rebuild and other preservation edits are not implemented.");
      if (mode === "no-op" && !sameState(candidate, baseline)) return blocked(mode, "save-document-changed", "No-op save requires the normalized document to equal its imported baseline.");
      if (mode === "edit" && !orderRequests.length && !coreFields.length && !coreRemovals.length && !removalIds.length && !additionIndexes.length && sameState(candidate, baseline)) return blocked(mode, "save-edit-no-changes", "The mapped document is unchanged; prepare a no-op plan to retain original ZIP bytes.");
      const coreBindings = coreFields.map(field => editState?.metadataBindings?.find(binding => binding.field === field && binding.insert));
      if (coreBindings.some(binding => !binding) || coreFields.some(field => field === "revision" ? !Number.isSafeInteger(candidate.metadata.revision) || candidate.metadata.revision < 0 || Object.is(candidate.metadata.revision, -0) : typeof candidate.metadata[field] !== "string")) return blocked(mode, "save-edit-unsupported", "Every requested core metadata insertion requires a missing bound source field and a Core string value.");
      const coreRemovalDefaults = { revision: 1, title: "PPTKit Presentation", author: "PPTKit", language: "en-US" } as const;
      const coreRemovalBindings = coreRemovals.map(field => editState?.metadataBindings?.find(binding => binding.field === field && binding.removable && !binding.insert));
      if (coreRemovalBindings.some(binding => !binding) || coreRemovals.some(field => candidate.metadata[field] !== coreRemovalDefaults[field])) return blocked(mode, "save-edit-unsupported", "Core metadata removal requires eligible bound fields and their normalized default values.");
      const selectedCoreRemovals = coreRemovalBindings.filter((binding): binding is CoreMetadataBinding => binding !== undefined);
      const coreWork = coreFields.length || coreRemovals.length ? (editState?.metadataBindings?.length ?? 0) * (coreFields.length + coreRemovals.length) + (editState?.guards.length ?? 0) + coreFields.reduce((sum, field) => sum + field.length + (String(candidate.metadata[field] ?? "").length) + 1, 0) : 0;
      const removalWork = coreWork + (removalIds.length ? removalIds.reduce((sum, id) => sum + id.length + 1, 0) + (editState?.customPropertyBindings?.length ?? 0) * (2 + Math.ceil(Math.log2(removalIds.length + 1))) : 0) + (additionIndexes.length ? (editState?.customInsertionSource?.usedIds.size ?? 0) + additionIndexes.length * 4 : 0);
      if (removalWork >= limits.maxSaveAnalysisWork) throw new PptxImportError("save-limit-exceeded", "Custom removal identity analysis exceeds maxSaveAnalysisWork.");
      const analysisLimits = removalWork ? { ...limits, maxSaveAnalysisWork: limits.maxSaveAnalysisWork - removalWork } : limits;
      const bindingsBySource = new Map(removalIds.length ? editState?.customPropertyBindings?.map(binding => [binding.sourceObjectId, binding]) ?? [] : []);
      const removalBindings: CustomPropertyBinding[] = [];
      for (const id of removalIds) {
        const binding = bindingsBySource.get(id);
        if (!binding) return blocked(mode, "save-edit-unsupported", "A requested custom property removal is not bound to the imported source.");
        removalBindings.push(binding);
      }
      const removedIndexes = new Set(removalBindings.map(binding => binding.index));
      if (removalBindings.length && !baseline!.metadata.customProperties) throw new PptxImportError("save-candidate-invalid", "Bound custom properties are unavailable in the baseline.");
      const survivors = removalBindings.length ? baseline!.metadata.customProperties!.filter((_, index) => !removedIndexes.has(index)) : baseline!.metadata.customProperties ?? (editState?.customInsertionSource?.create ? [] : undefined);
      const insertionSource = editState?.customInsertionSource;
      const customInsertions: NonNullable<import("./analyze-text-edits.js").TextEditAnalysis["customInsertions"]>[number][] = [];
      if (additionIndexes.length) {
        if (!insertionSource || !survivors || !candidate.metadata.customProperties) return blocked(mode, "save-edit-unsupported", "A bound existing custom property part is required for insertion.");
        const count = survivors.length;
        if (candidate.metadata.customProperties.length !== count + additionIndexes.length || additionIndexes.some((index, offset) => index !== count + offset)) return blocked(mode, "save-edit-unsupported", "Custom property additions must explicitly identify every appended Core entry in order.");
        const used = new Set(insertionSource.usedIds); let id = 2;
        for (const index of additionIndexes) {
          const property = candidate.metadata.customProperties[index]!;
          if (property.type === "decimal") return blocked(mode, "save-edit-unsupported", "Exact decimal custom property insertion is not supported by source-preserving profiles.");
          if (insertionSource.reservedNames.has(property.name)) return blocked(mode, "save-edit-unsupported", "A new custom property conflicts with an unprojected source name.");
          while (used.has(id) && id <= 2147483647) id++;
          if (id > 2147483647) return blocked(mode, "save-edit-unsupported", "No custom property ID is available.");
          used.add(id); customInsertions.push({ source: insertionSource, id, property, index }); id++;
        }
      }
      const analysisBaseline = removalBindings.length || customInsertions.length || selectedCoreRemovals.length ? { ...baseline!, metadata: { ...baseline!.metadata, ...(removalBindings.length || customInsertions.length ? { customProperties: [...(survivors ?? []), ...customInsertions.map(item => item.property)] } : {}), ...Object.fromEntries(coreRemovals.map(field => [field, coreRemovalDefaults[field]])) } } : baseline!;
      const sortedRemoved = [...removedIndexes].sort((a, b) => a - b);
      const countBefore = (index: number) => { let low = 0, high = sortedRemoved.length; while (low < high) { const mid = (low + high) >>> 1; if (sortedRemoved[mid]! < index) low = mid + 1; else high = mid; } return low; };
      const analysisCustomBindings = removalBindings.length ? editState?.customPropertyBindings?.filter(binding => !removedIndexes.has(binding.index)).map(binding => ({ ...binding, index: binding.index - countBefore(binding.index) })) : editState?.customPropertyBindings;
      // Edit reservations cover the full per-plan output budget before analysis
      // and candidate work. Ready plans shrink to their actual encoded size.
      const size = mode === "edit" ? limits.maxPreparedOutputBytes : source!.length;
      if (size > limits.maxPreparedOutputBytes || retainedBytes + size > limits.maxRetainedOutputBytes || active.size >= limits.maxLivePlans) throw new PptxImportError("save-limit-exceeded", "Prepared output exceeds the per-plan, retained-output or live-plan budget.");
      const slot: Slot = { size, disposed: false }; active.add(slot); retainedBytes += size;
      const operation = new AbortController(), callerSignal = signal;
      let unsubscribe = () => {};
      let ready = false;
      try {
        operations.add(operation);
        unsubscribe = subscribeAbort(callerSignal, operation);
        const signal = operation.signal;
        let bytes: Uint8Array, saveReport: PptxSaveReport;
        if (mode === "edit") {
          let work = 0;
          const checkpoint = async () => { abort(signal); if (++work % 64 === 0) await yieldToCaller(signal); };
          let numberOnly = false;
          if (editState!.pictureSources && editState!.firstNumberBinding && !Object.is(baseline!.firstSlideNumber, candidate.firstSlideNumber)) {
            const numberState = structuredClone(baseline!);
            if (candidate.firstSlideNumber === undefined) delete numberState.firstSlideNumber;
            else numberState.firstSlideNumber = candidate.firstSlideNumber;
            numberOnly = sameState(numberState, candidate);
          }
          if (editState!.pictureSources && !numberOnly) {
            // First isolate the number mutation. The picture analyzer must still
            // prove that every other normalized change is exactly one retarget.
            const numberChanged = !Object.is(baseline!.firstSlideNumber, candidate.firstSlideNumber);
            let numberState = baseline!, numberAnalysis: Awaited<ReturnType<typeof analyzeTextEdits>> | undefined;
            let pictureLimits = limits;
            if (numberChanged) {
              numberState = structuredClone(baseline!);
              if (candidate.firstSlideNumber === undefined) delete numberState.firstSlideNumber;
              else numberState.firstSlideNumber = candidate.firstSlideNumber;
              const analysisBudget = Math.floor(limits.maxSaveAnalysisWork / 3);
              if (analysisBudget < 1) throw new PptxImportError("save-limit-exceeded", "Combined picture and numbering analysis exceeds maxSaveAnalysisWork.");
              pictureLimits = { ...limits, maxSaveAnalysisWork: analysisBudget };
              numberAnalysis = await analyzeTextEdits(baseline!, numberState, [], [], pictureLimits, checkpoint, [], editState!.firstNumberBinding);
              if (numberAnalysis.diagnostics.length || !numberAnalysis.firstSlideNumber) return freezeReport({ status: "blocked", report: { mode, assurance: "none", losses: [], changes: numberAnalysis.changes, diagnostics: numberAnalysis.diagnostics } });
            }
            let picture: Awaited<ReturnType<typeof analyzePictureEdit>>;
            try { picture = await analyzePictureEdit(numberState, candidate, editState!.pictureSources, pictureLimits, checkpoint); }
            catch (error) { if (error instanceof PptxImportError && error.code === "save-candidate-invalid") return blocked(mode, "save-picture-edit-unsupported", error.message); throw error; }
            const guards = editState!.guards.filter(guard => (!guard.part || guard.part === picture.binding.source.part) && (!guard.sourceObjectId || guard.sourceObjectId === picture.binding.source.sourceObjectId));
            if (guards.length) return blocked(mode, "save-picture-source-guard", "Picture replacement is blocked by source package or object references.", guards.map(guard => guard.diagnostic));
            const changes = [...(numberAnalysis?.changes ?? []), ...picture.changes];
            if (changes.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Combined picture and numbering changes exceed maxSaveChanges.");
            try {
              let pictureInput = source!;
              if (numberAnalysis) {
                if (numberAnalysis.firstSlideNumber!.part === picture.binding.source.part) return blocked(mode, "save-picture-edit-unsupported", "Picture and numbering edits cannot share one source part.");
                const staged = await prepareTextCandidate(source!, numberState, numberAnalysis, { ...editState!.options, ...(signal ? { signal } : {}) }, pictureLimits, checkpoint);
                if ("diagnostics" in staged) return freezeReport({ status: "blocked", report: { mode, assurance: "none", losses: [], changes, diagnostics: staged.diagnostics } });
                const maxSaveValidationBytes = limits.maxSaveValidationBytes - staged.validationBytes, maxSaveValidationWork = limits.maxSaveValidationWork - staged.validationWork, maxPreparedOutputBytes = limits.maxPreparedOutputBytes - staged.bytes.length;
                if (maxSaveValidationBytes <= 0 || maxSaveValidationWork <= 0 || maxPreparedOutputBytes <= 0) throw new PptxImportError("save-limit-exceeded", "Combined candidate stages exhausted their shared budget.");
                pictureLimits = { ...pictureLimits, maxSaveValidationBytes, maxSaveValidationWork, maxPreparedOutputBytes };
                pictureInput = staged.bytes;
              }
              const prepared = await preparePictureCandidate(pictureInput, picture.binding, picture.bytes, picture.mimeType, { ...editState!.options, ...(signal ? { signal } : {}) }, pictureLimits, checkpoint);
              bytes = prepared.bytes;
            } catch (error) {
              if (error instanceof PictureSourceUnsupportedError) return blocked(mode, "save-picture-candidate-unsupported", error.message);
              if (error instanceof PptxImportError && error.code === "import-limit-exceeded") throw new PptxImportError("save-limit-exceeded", error.message, error.source, { cause: error });
              if (error instanceof PptxImportError && error.code.startsWith("import-")) return blocked(mode, "save-picture-candidate-unsupported", error.message);
              if (error instanceof XmlSourceEditError) {
                if (error.code === "xml-edit-limit-exceeded") throw new PptxImportError("save-limit-exceeded", error.message, undefined, { cause: error });
                throw new PptxImportError("save-candidate-invalid", error.message, undefined, { cause: error });
              }
              if (error instanceof PackageInspectionError) {
                if (error.code === "operation-aborted") throw new PptxImportError("operation-aborted", error.message, undefined, { cause: error });
                if (error.code.endsWith("limit-exceeded")) throw new PptxImportError("save-limit-exceeded", error.message, undefined, { cause: error });
                throw new PptxImportError("save-candidate-invalid", error.message, undefined, { cause: error });
              }
              throw error;
            }
            retainedBytes -= slot.size - bytes.length; slot.size = bytes.length;
            saveReport = freezeReport({ mode, assurance: "scoped-preservation", losses: [], changes, diagnostics: [] });
          } else {
          let mapped: Awaited<ReturnType<typeof analyzeChartOrders>> | undefined;
          let scalarLimits = analysisLimits;
          if (orderRequests.length) {
            const captureWork = orderRequests.reduce((sum, item) => sum + 1 + item.slideId.length + item.elementId.length + item.identities.length, 0);
            const guardWork = editState!.guards.length;
            const available = analysisLimits.maxSaveAnalysisWork - captureWork - guardWork;
            if (available <= 0) throw new PptxImportError("save-limit-exceeded", "Chart order capture and guards exhaust the analysis budget.");
            try { mapped = await analyzeChartOrders(analysisBaseline, candidate, editState!.chartOrderBindings!, orderRequests, available, checkpoint); }
            catch (error) { if (error instanceof PptxImportError && error.code === "save-candidate-invalid") return blocked(mode, "save-chart-order-unsupported", error.message); throw error; }
            const parts = new Set(mapped.proposals.flatMap(({ binding }) => [binding.part, binding.slidePart]));
            const guards = editState!.guards.filter(guard => !guard.part || parts.has(guard.part));
            if (guards.length) return blocked(mode, "save-chart-order-source-guard", "Chart series order is blocked by source package or slide references.", guards.map(guard => guard.diagnostic));
            const remaining = available - mapped.work;
            if (remaining <= 0 || mapped.changes.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Chart order exhausts the shared analysis budget.");
            scalarLimits = { ...analysisLimits, maxSaveAnalysisWork: remaining, maxSaveChanges: limits.maxSaveChanges - mapped.changes.length };
          }
          const scalarCandidate = mapped?.candidateForScalars ?? candidate;
          const descriptionAnalysis = editState!.descriptionBindings ? await analyzeDescriptionEdit(analysisBaseline, scalarCandidate, editState!.descriptionBindings, scalarLimits, checkpoint, editState!.descriptionBatch, editState!.descriptionMixed) : undefined;
          if (descriptionAnalysis?.diagnostics.length) return freezeReport({ status: 'blocked', report: { mode, assurance: 'none', losses: [], changes: descriptionAnalysis.changes, diagnostics: descriptionAnalysis.diagnostics } });
          let analysis: import("./analyze-text-edits.js").TextEditAnalysis;
          if (descriptionAnalysis && editState!.descriptionMixed) {
            const remainingWork = scalarLimits.maxSaveAnalysisWork - (descriptionAnalysis.work ?? 0) - descriptionAnalysis.descriptions.length - editState!.guards.length;
            if (remainingWork <= 0) throw new PptxImportError('save-limit-exceeded', 'Combined description and text analysis exceeds the shared work budget.');
            const textLimits = { ...scalarLimits, maxSaveAnalysisWork: remainingWork, maxSaveChanges: scalarLimits.maxSaveChanges - descriptionAnalysis.changes.length };
            const text = await analyzeTextEdits(analysisBaseline, descriptionAnalysis.remainder!, editState!.bindings, editState!.guards, textLimits, checkpoint, editState!.chartBindings, editState!.firstNumberBinding, editState!.slideVisibilityBindings, editState!.metadataBindings, analysisCustomBindings, editState!.chartLabelBindings);
            analysis = { ...text, descriptions: descriptionAnalysis.descriptions, changes: [...descriptionAnalysis.changes, ...text.changes].sort((a, b) => a.documentPath < b.documentPath ? -1 : a.documentPath > b.documentPath ? 1 : 0) };
          } else analysis = descriptionAnalysis ? { ...descriptionAnalysis, parts: new Map(), chartValues: [] } : await analyzeTextEdits(analysisBaseline, scalarCandidate, editState!.bindings, editState!.guards, scalarLimits, checkpoint, editState!.chartBindings, editState!.firstNumberBinding, editState!.slideVisibilityBindings, editState!.metadataBindings, analysisCustomBindings, editState!.chartLabelBindings);
          if (mapped) analysis = { ...analysis, chartOrders: mapped.proposals, changes: [...analysis.changes, ...mapped.changes].sort((a, b) => a.documentPath < b.documentPath ? -1 : a.documentPath > b.documentPath ? 1 : 0) };
          if (selectedCoreRemovals.length) {
            if (analysis.changes.length + selectedCoreRemovals.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Core metadata removals exceed maxSaveChanges.");
            const diagnostics = [...analysis.diagnostics], parts = new Set(selectedCoreRemovals.map(binding => binding.part));
            for (const guard of editState!.guards) { await checkpoint(); if ((!guard.part || parts.has(guard.part)) && !diagnostics.includes(guard.diagnostic)) { if (diagnostics.length >= limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Core metadata source guards exceed maxSaveChanges."); diagnostics.push(guard.diagnostic); } }
            const changes = selectedCoreRemovals.map(binding => ({ kind: "metadata" as const, documentPath: `/metadata/${binding.field}`, source: { part: binding.part, elementIndex: binding.elementIndex, path: `coreProperties/${binding.local}` } }));
            analysis = { ...analysis, diagnostics, metadataRemovals: [...(analysis.metadataRemovals ?? []), ...selectedCoreRemovals], changes: [...analysis.changes, ...changes].sort((a,b) => a.documentPath < b.documentPath ? -1 : a.documentPath > b.documentPath ? 1 : 0) };
          }
          if (coreFields.length) {
            const existing = new Set((analysis.metadataInsertions ?? []).map(item => item.binding.field));
            const forced = coreBindings.filter((binding): binding is CoreMetadataBinding => Boolean(binding) && !existing.has(binding!.field)).map(binding => ({ binding, value: String(candidate.metadata[binding.field as typeof coreFields[number]]) }));
            if (analysis.changes.length + forced.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Core metadata insertions exceed maxSaveChanges.");
            const parts = new Set(forced.map(item => item.binding.part));
            const diagnostics = [...analysis.diagnostics];
            for (const guard of editState!.guards) {
              await checkpoint();
              if ((!guard.part || parts.has(guard.part)) && !diagnostics.includes(guard.diagnostic)) {
                if (diagnostics.length >= limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Core metadata source guards exceed maxSaveChanges.");
                diagnostics.push(guard.diagnostic);
              }
            }
            analysis = { ...analysis, diagnostics, metadataInsertions: [...(analysis.metadataInsertions ?? []), ...forced], changes: [...analysis.changes, ...forced.map(({ binding }) => ({ kind: "metadata" as const, documentPath: `/metadata/${binding.field}`, source: { part: binding.part, path: `coreProperties/${binding.local}` } }))].sort((a, b) => a.documentPath < b.documentPath ? -1 : a.documentPath > b.documentPath ? 1 : 0) };
          }
          if (removalBindings.length || customInsertions.length) {
            if (analysis.changes.length + removalBindings.length + customInsertions.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Combined custom removals exceed maxSaveChanges.");
            const removalChanges = removalBindings.map(binding => ({ kind: "metadata" as const, documentPath: `/metadata/customProperties/${binding.index}`, source: { part: binding.part, elementIndex: binding.propertyElementIndex, path: "Properties/property", sourceObjectId: binding.sourceObjectId } }));
            const insertionChanges = customInsertions.map(item => ({ kind: "metadata" as const, documentPath: `/metadata/customProperties/${item.index}`, source: { part: item.source.part, path: "Properties/property", sourceObjectId: `{D5CDD505-2E9C-101B-9397-08002B2CF9AE}:${item.id}` } }));
            const removalParts = new Set([...removalBindings.map(binding => binding.part), ...customInsertions.map(item => item.source.part)]), removalDiagnostics: ImportDiagnostic[] = [];
            for (const guard of editState!.guards) {
              await checkpoint();
              if (guard.part && !removalParts.has(guard.part)) continue;
              if (removalDiagnostics.length + analysis.diagnostics.length >= limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Custom removal diagnostics exceed maxSaveChanges.");
              removalDiagnostics.push(guard.diagnostic);
            }
            analysis = { ...analysis, customRemovals: removalBindings, customInsertions, diagnostics: [...analysis.diagnostics, ...removalDiagnostics], changes: [...analysis.changes, ...removalChanges, ...insertionChanges].sort((a, b) => a.documentPath < b.documentPath ? -1 : a.documentPath > b.documentPath ? 1 : 0) };
          }
          abort(signal);
          if (analysis.diagnostics.length) return freezeReport({ status: "blocked", report: { mode, assurance: "none", losses: [], changes: analysis.changes, diagnostics: analysis.diagnostics } });
          if (descriptionAnalysis?.descriptions.length) {
            if ((descriptionAnalysis.work ?? 0) + descriptionAnalysis.descriptions.length + editState!.guards.length > limits.maxSaveAnalysisWork) throw new PptxImportError('save-limit-exceeded', 'Description source guards exceed the shared analysis budget.');
            const parts = new Set(descriptionAnalysis.descriptions.map(({ binding }) => binding.part));
            const objects = new Set(descriptionAnalysis.descriptions.map(({ binding }) => binding.sourceObjectId));
            const pairs = new Set(descriptionAnalysis.descriptions.map(({ binding }) => JSON.stringify([binding.part, binding.sourceObjectId])));
            const relevant: TextEditGuard[] = [];
            for (const guard of editState!.guards) {
              await checkpoint();
              if (guard.part && guard.sourceObjectId ? pairs.has(JSON.stringify([guard.part, guard.sourceObjectId])) : guard.part ? parts.has(guard.part) : guard.sourceObjectId ? objects.has(guard.sourceObjectId) : true) {
                if (relevant.length >= limits.maxSaveChanges) throw new PptxImportError('save-limit-exceeded', 'Description source guards exceed maxSaveChanges.');
                relevant.push(guard);
              }
            }
            if (relevant.length) return freezeReport({ status: 'blocked', report: { mode, assurance: 'none', losses: [], changes: analysis.changes, diagnostics: relevant.map(guard => guard.diagnostic) } });
          }
          const prepared = await prepareTextCandidate(source!, candidate, analysis, { ...editState!.options, ...(signal ? { signal } : {}) }, limits, checkpoint, editState!.diagramPreview, editState!.textFontProfile, editState!.notesTextProfile, editState!.fieldTextProfile, editState!.currencySave, editState!.decimalSave);
          abort(signal);
          if ("diagnostics" in prepared) return freezeReport({ status: "blocked", report: { mode, assurance: "none", losses: [], changes: analysis.changes, diagnostics: prepared.diagnostics } });
          bytes = prepared.bytes;
          retainedBytes -= slot.size - bytes.length; slot.size = bytes.length;
          saveReport = freezeReport({ mode, assurance: "scoped-preservation", losses: [], changes: analysis.changes, diagnostics: [] });
          }
        } else {
          bytes = new Uint8Array(size);
          for (let offset = 0; offset < size; offset += 1024 * 1024) {
            abort(signal); bytes.set(source!.subarray(offset, Math.min(size, offset + 1024 * 1024)), offset);
            await yieldToCaller(signal);
          }
          saveReport = freezeReport({ mode, assurance: "original-bytes", losses: [], diagnostics: [] });
        }
        if (!sameState(candidate, currentSnapshot())) return blocked(mode, "save-snapshot-stale", "The Core document changed during output preparation.");
        abort(signal);
        slot.bytes = bytes; slot.snapshot = candidate;
        const plan = Object.freeze({ report: saveReport, byteLength: bytes.length, dispose: () => release(slot), toJSON: notSerializable }) as unknown as PptxSavePlan;
        ownership.set(plan, slot); ready = true;
        return Object.freeze({ status: "ready", plan });
      } catch (e) {
        live();
        if (e instanceof PptxImportError && e.code === "save-invalid-document") return invalidDocument(mode, e);
        throw e;
      } finally {
        operations.delete(operation); unsubscribe();
        if (!ready) release(slot);
      }
    },
    assertSavePlanCurrent(plan, options = {}) { currentPlan(plan, options); },
    async save(plan, options = {}) {
      const slot = currentPlan(plan, options);
      return { bytes: Uint8Array.from(slot.bytes!), report: plan.report };
    },
    dispose() {
      if (disposed) return;
      disposed = true; source = undefined; baseline = undefined;
      editState = undefined;
      for (const operation of operations) operation.abort();
      // Detach caller-owned signals synchronously, even while child work is
      // unwinding. The per-operation finally blocks can unsubscribe again.
      for (const [signal, link] of abortLinks) {
        signal.removeEventListener("abort", link.abort); link.controllers.clear();
      }
      abortLinks.clear();
      for (const slot of active) release(slot);
    },
    toJSON: notSerializable,
  };
  return Object.freeze(session);
}
