import { captureSections } from "./sections.js";
import { captureSlideReadingOrder, reconcileSlideReadingOrder, remapSlideReadingOrder } from "../accessibility/reading-order.js";
import { assertCustomDocumentProperties, isCanonicalUtcDate } from "./custom-document-properties.js";
import { slideRebindData, stageLayoutRebind } from "./layout-rebind.js";
import { LayoutRebindError, type LayoutRebindPlan, type LayoutRebindApplyResult } from "../types/layout-rebind.js";
import { remapSlideEventTimeline } from "./slide-event-timeline.js";
import { remapTextActions, setCopiedReference } from "./copy-references.js";
import { collectSlideCopyDependencies } from "./slide-copy-dependencies.js";
import { fitCopiedGeometry } from "./copy-geometry.js";
import { equivalentDefinitionElements } from "./definition-equivalence.js";
import { DEFAULT_LAYOUT_ID, DEFAULT_PRESENTATION_SIZE } from "../constants/presentation.js";
import type { ElementAction, PlaceholderDefinitionInput, PresentationElement, PresentationElementInput, TextContentInput, TextElementInput } from "../types/element.js";
import type { AssetReference, AssetRegistrationOptions, AssetRegistrationRequest, AssetCopyOptions, AssetCopyMapping, AssetRetentionOptions, PresentationAssetReplacement, PresentationAsset, PresentationAssetInput } from "../types/asset.js";
import type { PresentationSize } from "../types/geometry.js";
import type {
  ManagedPresentationDocument,
  ManagedPresentationSlide,
  PresentationInit,
  PresentationDocument,
  PresentationMetadataInput,
  PresentationSection,
  PresentationSlide,
  PresentationSlideInput,
  SlideLayoutDefinition,
  SlideLayoutInput,
  SlideMasterDefinition,
  SlideMasterInput,
  SlideCopyOptions,
  SlideCopyResult,
  SlidesCopyOptions,
  SlidesCopyResult,
  LayoutCopyOptions,
  LayoutCopyResult,
  MasterCopyOptions,
  MasterCopyResult,
} from "../types/presentation.js";
import { PresentationValidationError } from "../types/diagnostic.js";
import { validatePresentation } from "../validation/validate-presentation.js";
import { assertDefinitionElements, assertDefinitionIdentity, captureDefinitionInput } from "./definition-input.js";
import type { PresentationThemeInput } from "../types/theme.js";
import type { TextStylePresetMap } from "../types/style.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import { defaultTheme, resolveTheme } from "../normalization/inheritance.js";
import { createDocumentId, IdAllocator } from "../utils/id.js";
import { normalizeSize } from "../validation/geometry.js";
import { collectAssetReferences } from "./asset-references.js";
import { AssetRegistry } from "./asset-registry.js";
import { ElementStore } from "./element-store.js";
import { PresentationSlideImpl } from "./presentation-slide.js";
import { captureEditOperation } from "./edit-operation.js";
import { PresentationEditError, type PresentationEditAction, type PresentationEditReadingOrder, type PresentationEditChange, type PresentationEditOperation, type PresentationEditResult } from "../types/edit-operation.js";

function assertInsertIndex(index: number, length: number, label: string): void {
  if (!Number.isInteger(index) || index < 0 || index > length) {
    throw new RangeError(`${label} index ${index} is outside 0..${length}.`);
  }
}

function assertCopyAssetFields(value: object, allowed: readonly string[]): void {
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== "string" || !allowed.includes(key) || !Object.getOwnPropertyDescriptor(value, key)?.enumerable) throw new TypeError("Source asset has unsupported fields or source type.");
  }
}

function isPlaceholderIndex(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= 4294967295;
}

type TextEdit = Extract<PresentationEditChange, { type: "text-run" | "table-cell-text-run" | "text-run-action" | "table-cell-text-run-action" | "text-string" | "table-cell-text-string" }>;
function editableText(element: PresentationElement, change: TextEdit): TextContentInput | undefined {
  if (change.type === "table-cell-text-run" || change.type === "table-cell-text-run-action" || change.type === "table-cell-text-string") return element.type === "table" ? element.rows?.[change.rowIndex]?.cells?.[change.cellIndex]?.content : undefined;
  return element.type === "text" ? element.content : element.type === "shape" ? element.text?.content : undefined;
}
function actionMatches(current: ElementAction | undefined, expected: PresentationEditAction): boolean {
  if (expected === null) return current === undefined;
  if (!current || current.type !== expected.type || Reflect.ownKeys(current).some(key => typeof key !== "string" || !["type", expected.type === "url" ? "url" : "slideId", "tooltip"].includes(key) || (current as unknown as Record<string, unknown>)[key] === undefined) || (current.tooltip ?? null) !== expected.tooltip || current.tooltip === null) return false;
  return current.type === "url" && expected.type === "url" ? current.url === expected.url : current.type === "slide" && expected.type === "slide" && current.slideId === expected.slideId;
}
function materializeAction(value: Exclude<PresentationEditAction, null>): ElementAction {
  return value.type === "url" ? { type: "url", url: value.url, ...(value.tooltip === null ? {} : { tooltip: value.tooltip }) } : { type: "slide", slideId: value.slideId, ...(value.tooltip === null ? {} : { tooltip: value.tooltip }) };
}

function setTextString(element: PresentationElement, change: TextEdit, value: string): void {
  if (change.type === "table-cell-text-string" && element.type === "table") element.rows[change.rowIndex]!.cells[change.cellIndex]!.content = value;
  else if (change.type === "text-string" && element.type === "text") element.content = value;
  else if (change.type === "text-string" && element.type === "shape" && element.text) element.text.content = value;
  else throw new Error("Invalid managed text target.");
}

function captureCopyMapping(value: unknown): Readonly<Record<string, string>> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Copy mappings must be plain records of nonempty string targets.");
  const captured = captureDefinitionInput(value, Object.keys(value)) as Record<string, unknown>;
  if (Object.values(captured).some(target => typeof target !== "string" || !target)) throw new TypeError("Copy mappings must be plain records of nonempty string targets.");
  return captured as Record<string, string>;
}

export class PresentationDocumentImpl implements ManagedPresentationDocument {
  readonly id: string;
  private sectionsValue: readonly PresentationSection[] | undefined;
  private firstSlideNumberValue: number | undefined = undefined;
  private metadataValue: Readonly<PresentationMetadataInput>;
  private readonly sizeValue: Readonly<PresentationSize>;
  private readonly themeValue: Readonly<PresentationThemeInput>;
  private readonly textStylePresetsValue: TextStylePresetMap;
  private readonly slideEntries: PresentationSlideImpl[] = [];
  private masterEntries: readonly SlideMasterDefinition[];
  private readonly layoutEntries: SlideLayoutDefinition[] = [];
  private readonly assetRegistry = new AssetRegistry();
  private readonly elementStore = new ElementStore();
  private readonly slideIds = new Set<string>();
  private slideAllocator = new IdAllocator("slide");
  private slideRevision = 0;
  private definitionRevision = 0;

  constructor(init: PresentationInit = {}) {
    this.setFirstSlideNumber(init.firstSlideNumber);
    if (init.masters !== undefined && (!Array.isArray(init.masters) || init.masters.some(master => !master || typeof master !== "object" || Array.isArray(master) || Reflect.ownKeys(master).some(key => typeof key !== "string" || !["id", "name", "theme", "background", "elements"].includes(key))))) throw new TypeError("Masters must be declarations with id, name, theme, background and elements only.");
    this.masterEntries = deepFreeze((init.masters ?? []).map((master: SlideMasterInput) => {
      if (master.elements !== undefined && !Array.isArray(master.elements)) throw new TypeError("Master elements must be an array.");
      assertDefinitionElements(master.elements ?? []);
      const { elements, ...definition } = deepClone(master);
      return { ...definition, ...(elements === undefined ? {} : { elements: elements.map((element: PresentationElementInput) => this.elementStore.materialize(element)) }) };
    }));
    this.id = init.id ?? createDocumentId();
    if (init.metadata?.customProperties !== undefined) assertCustomDocumentProperties(init.metadata.customProperties);
    if ([init.metadata?.createdAt, init.metadata?.modifiedAt].some(value => value !== undefined && !isCanonicalUtcDate(value))) throw new TypeError("Metadata dates must be canonical UTC timestamps.");
    const initialMetadata = deepClone(init.metadata ?? {});
    if (initialMetadata.customProperties !== undefined) assertCustomDocumentProperties(initialMetadata.customProperties);
    if ([initialMetadata.createdAt, initialMetadata.modifiedAt].some(value => value !== undefined && !isCanonicalUtcDate(value))) throw new TypeError("Metadata dates must be canonical UTC timestamps.");
    this.metadataValue = deepFreeze(initialMetadata);
    this.sizeValue = deepFreeze(normalizeSize(init.size, DEFAULT_PRESENTATION_SIZE));
    this.themeValue = deepFreeze(deepClone(init.theme ?? {}));
    this.textStylePresetsValue = deepFreeze(deepClone(init.textStylePresets ?? {}));
  }

  get masters(): readonly SlideMasterDefinition[] { return this.masterEntries; }

  get firstSlideNumber(): number | undefined { return this.firstSlideNumberValue; }
  setFirstSlideNumber(value?: number): void {
    if (value !== undefined && (!Number.isSafeInteger(value) || value < 0)) throw new RangeError("firstSlideNumber must be a nonnegative safe integer.");
    this.firstSlideNumberValue = value;
  }

  applyEditOperation(input: PresentationEditOperation): PresentationEditResult {
    const operation = captureEditOperation(input);
    if (operation.documentId !== this.id) throw new PresentationEditError("wrong-document", "Edit operation belongs to a different document.");
    const byId = new Map(this.slideEntries.map(slide => [slide.id, slide]));
    const staged: { slide: PresentationSlideImpl; value: boolean }[] = [];
    const textEdits = new Map<string, { slide: PresentationSlideImpl; original: PresentationElement; input: PresentationElement }>();
    const textTargets = new Map<PresentationSlideImpl, Map<string, { root: PresentationElement; element: PresentationElement; path: number[] }>>();
    const notesEdits = new Map<PresentationSlideImpl, TextContentInput>();
    const requestedOrders = new Map<PresentationSlideImpl, PresentationEditReadingOrder>();
    let nextNumber = this.firstSlideNumberValue;
    let nextOrder: PresentationSlideImpl[] | undefined;
    const inverse: PresentationEditChange[] = [];
    for (const change of operation.changes) {
      if (change.type === "chart-category" || change.type === "chart-series-name" || change.type === "chart-point" || change.type === "element-accessibility" || change.type === "element-action" || change.type === "element-transform" || change.type === "element-name" || change.type === "element-hidden" || change.type === "element-opacity" || change.type === "element-box" || change.type === "text-run" || change.type === "table-cell-text-run" || change.type === "text-run-action" || change.type === "table-cell-text-run-action" || change.type === "text-string" || change.type === "table-cell-text-string") {
        const slide = byId.get(change.slideId);
        if (!slide) throw new PresentationEditError("missing-target", "Element slide does not exist.");
        let targets = textTargets.get(slide);
        if (!targets) {
          targets = new Map();
          const stack = slide.elements.map(root => ({ root, element: root, path: [] as number[] }));
          while (stack.length) {
            const target = stack.pop()!;
            targets.set(target.element.id, target);
            if (target.element.type === "group") for (let i = 0; i < target.element.children.length; i++) stack.push({ root: target.root, element: target.element.children[i]!, path: [...target.path, i] });
          }
          textTargets.set(slide, targets);
        }
        const target = targets.get(change.elementId);
        if (!target) throw new PresentationEditError("missing-target", "Element target does not exist.");
        if (change.type === "chart-category" || change.type === "chart-series-name") {
          const chart = target.element;
          if (chart.type !== "chart") throw new PresentationEditError("missing-target", "Chart target does not exist.");
          let current: unknown;
          if (change.type === "chart-category") {
            if (chart.chartType === "scatter" || chart.chartType === "bubble" || !Array.isArray(chart.categories) || !Object.prototype.hasOwnProperty.call(chart.categories, change.categoryIndex)) throw new PresentationEditError("missing-target", "Chart category does not exist.");
            current = chart.categories[change.categoryIndex];
          } else current = chart.series?.[change.seriesIndex]?.name;
          if (typeof current !== "string") throw new PresentationEditError("missing-target", "Chart text target does not exist.");
          if (current !== change.expected) throw new PresentationEditError("edit-conflict", "Chart text has changed.");
        } else if (change.type === "chart-point") {
          if (target.element.type !== "chart") throw new PresentationEditError("missing-target", "Chart target does not exist.");
          const chart = target.element, numeric = chart.chartType === "scatter" || chart.chartType === "bubble";
          if (numeric ? change.coordinate === "values" || (chart.chartType === "scatter" && change.coordinate === "sizes") : change.coordinate !== "values") throw new PresentationEditError("missing-target", "Chart coordinate is not supported by this chart type.");
          const series = chart.series[change.seriesIndex] as unknown as Record<string, unknown> | undefined;
          const values = series?.[change.coordinate];
          if (!Array.isArray(values) || !Object.prototype.hasOwnProperty.call(values, change.pointIndex)) throw new PresentationEditError("missing-target", "Chart point does not exist.");
          const expected = change.expected === "-0" ? -0 : change.expected, next = change.value === "-0" ? -0 : change.value;
          if (!Object.is(values[change.pointIndex], expected)) throw new PresentationEditError("edit-conflict", "Chart point has changed.");
          if ((chart.chartType === "bubble" && (next === null || (change.coordinate === "sizes" && next <= 0))) ||
            ((chart.chartType === "pie" || chart.chartType === "doughnut") && next !== null && next < 0)) throw new PresentationEditError("invalid-operation", "Chart point value is invalid for this chart type.");
        } else if (change.type === "element-name" || change.type === "element-hidden" || change.type === "element-opacity") {
          const current = change.type === "element-name" ? target.element.name : change.type === "element-hidden" ? target.element.hidden : target.element.opacity;
          if (current === null || (current === undefined ? null : current) !== change.expected) throw new PresentationEditError("edit-conflict", "Element property has changed.");
        } else if (change.type === "element-accessibility") {
          const current = target.element.accessibility, expected = change.expected;
          if (expected === null ? current !== undefined : !current ||
            Reflect.ownKeys(current).some(key => typeof key !== "string" || !["description", "decorative"].includes(key) || current[key as "description" | "decorative"] === undefined) ||
            (current.description ?? null) !== expected.description || (current.decorative ?? null) !== expected.decorative || current.description === null || current.decorative === null) throw new PresentationEditError("edit-conflict", "Element accessibility has changed or has unsupported properties.");
        } else if (change.type === "element-action") {
          if (!actionMatches(target.element.action, change.expected)) throw new PresentationEditError("edit-conflict", "Element action has changed or has unsupported properties.");
          if (change.value?.type === "slide" && !byId.has(change.value.slideId)) throw new PresentationEditError("missing-target", "Action destination slide does not exist.");
        } else if (change.type === "element-transform") {
          const current = target.element.transform, expected = change.expected;
          if (expected === null ? current !== undefined : !current ||
            Reflect.ownKeys(current).some(key => typeof key !== "string" || !["rotation", "flipH", "flipV"].includes(key) || current[key as "rotation" | "flipH" | "flipV"] === undefined) ||
            (current.rotation ?? null) !== expected.rotation || (current.flipH ?? null) !== expected.flipH || (current.flipV ?? null) !== expected.flipV ||
            current.rotation === null || current.flipH === null || current.flipV === null) throw new PresentationEditError("edit-conflict", "Element transform has changed or has unsupported properties.");
        } else if (change.type === "element-box") {
          if (target.element.type === "connector") throw new PresentationEditError("missing-target", "Connectors use endpoints rather than an editable box.");
          if ((change.value?.height === null || change.expected?.height === null) && target.element.type !== "text") throw new PresentationEditError("invalid-operation", "Only text elements can have automatic box height.");
          const current = target.element.box, expected = change.expected;
          if (expected === null ? current !== undefined : current === undefined || current.x !== expected.x || current.y !== expected.y || current.width !== expected.width || (current.height === undefined ? null : current.height) !== expected.height || current.height === null) throw new PresentationEditError("edit-conflict", "Element box has changed.");
        } else {
          const content = editableText(target.element, change);
          if (change.type === "text-string" || change.type === "table-cell-text-string") {
            if (typeof content !== "string") throw new PresentationEditError("missing-target", "Text target does not contain string shorthand.");
            if (content !== change.expected) throw new PresentationEditError("edit-conflict", "Text string has changed.");
          } else {
            if (!Array.isArray(content)) throw new PresentationEditError("missing-target", "Text run target does not contain explicit paragraph/run content.");
            const paragraph = content[change.paragraphIndex];
            const run = paragraph && Array.isArray(paragraph.runs) ? paragraph.runs[change.runIndex] : undefined;
            if (!run || typeof run.text !== "string") throw new PresentationEditError("missing-target", "Text run target does not exist.");
            if (change.type === "text-run-action" || change.type === "table-cell-text-run-action") {
              if (!actionMatches(run.action, change.expected)) throw new PresentationEditError("edit-conflict", "Text run action has changed or has unsupported properties.");
              if (change.value?.type === "slide" && !byId.has(change.value.slideId)) throw new PresentationEditError("missing-target", "Action destination slide does not exist.");
            } else if (run.text !== change.expected) throw new PresentationEditError("edit-conflict", "Text run has changed.");
          }
        }
        let edit = textEdits.get(target.root.id);
        if (!edit) { edit = { slide, original: target.root, input: deepClone(target.root) }; textEdits.set(target.root.id, edit); }
        let edited = edit.input;
        for (const index of target.path) {
          if (edited.type !== "group") throw new Error("Invalid managed element path.");
          edited = edited.children[index]!;
        }
        if (change.type === "chart-category" || change.type === "chart-series-name") {
          if (edited.type !== "chart") throw new Error("Invalid managed chart target.");
          if (change.type === "chart-category") {
            if (edited.chartType === "scatter" || edited.chartType === "bubble") throw new Error("Invalid managed chart category.");
            edited.categories[change.categoryIndex] = change.value;
          } else edited.series[change.seriesIndex]!.name = change.value;
        } else if (change.type === "chart-point") {
          if (edited.type !== "chart") throw new Error("Invalid managed chart target.");
          const series = edited.series[change.seriesIndex] as unknown as Record<string, (number | null)[]>;
          series[change.coordinate]![change.pointIndex] = change.value === "-0" ? -0 : change.value;
        } else if (change.type === "element-accessibility") {
          if (change.value === null) delete edited.accessibility;
          else edited.accessibility = { ...(change.value.description === null ? {} : { description: change.value.description }), ...(change.value.decorative === null ? {} : { decorative: change.value.decorative }) };
        } else if (change.type === "element-action") {
          if (change.value === null) delete edited.action;
          else edited.action = materializeAction(change.value);
        } else if (change.type === "element-transform") {
          if (change.value === null) delete edited.transform;
          else edited.transform = { ...(change.value.rotation === null ? {} : { rotation: change.value.rotation }), ...(change.value.flipH === null ? {} : { flipH: change.value.flipH }), ...(change.value.flipV === null ? {} : { flipV: change.value.flipV }) };
        } else if (change.type === "element-name") {
          if (change.value === null) delete edited.name; else edited.name = change.value;
        } else if (change.type === "element-hidden") {
          if (change.value === null) delete edited.hidden; else edited.hidden = change.value;
        } else if (change.type === "element-opacity") {
          if (change.value === null) delete edited.opacity; else edited.opacity = change.value;
        } else if (change.type === "element-box") {
          if (change.value === null) delete edited.box;
          else if (change.value.height === null && edited.type === "text") (edited as TextElementInput).box = { x: change.value.x, y: change.value.y, width: change.value.width };
          else if (change.value.height !== null) edited.box = { ...change.value, height: change.value.height };
        } else if (change.type === "text-run-action" || change.type === "table-cell-text-run-action") {
          const run = (editableText(edited, change) as Exclude<TextContentInput, string>)[change.paragraphIndex]!.runs[change.runIndex]!;
          if (change.value === null) delete run.action; else run.action = materializeAction(change.value);
        } else if (change.type === "text-string" || change.type === "table-cell-text-string") setTextString(edited, change, change.value);
        else (editableText(edited, change) as Exclude<TextContentInput, string>)[change.paragraphIndex]!.runs[change.runIndex]!.text = change.value;
        inverse.push(Object.freeze({ ...change, expected: change.value, value: change.expected }) as PresentationEditChange);
      } else if (change.type === "notes-string" || change.type === "notes-run" || change.type === "notes-run-action") {
        const slide = byId.get(change.slideId);
        if (!slide) throw new PresentationEditError("missing-target", "Notes slide does not exist.");
        const original = slide.notes;
        if (change.type === "notes-string") {
          if (typeof original !== "string") throw new PresentationEditError("missing-target", "Notes do not contain string shorthand.");
          if (original !== change.expected) throw new PresentationEditError("edit-conflict", "Notes string has changed.");
          notesEdits.set(slide, change.value);
        } else {
          if (!Array.isArray(original)) throw new PresentationEditError("missing-target", "Notes do not contain explicit paragraph/run content.");
          const run = original[change.paragraphIndex]?.runs?.[change.runIndex];
          if (!run || typeof run.text !== "string") throw new PresentationEditError("missing-target", "Notes run does not exist.");
          if (change.type === "notes-run") {
            if (run.text !== change.expected) throw new PresentationEditError("edit-conflict", "Notes run has changed.");
          } else {
            if (!actionMatches(run.action, change.expected)) throw new PresentationEditError("edit-conflict", "Notes run action has changed or has unsupported properties.");
            if (change.value?.type === "slide" && !byId.has(change.value.slideId)) throw new PresentationEditError("missing-target", "Action destination slide does not exist.");
          }
          let edited = notesEdits.get(slide);
          if (edited === undefined) { edited = deepClone(original); notesEdits.set(slide, edited); }
          if (!Array.isArray(edited)) throw new PresentationEditError("invalid-operation", "Notes representation cannot change within a batch.");
          const next = edited[change.paragraphIndex]!.runs[change.runIndex]!;
          if (change.type === "notes-run") next.text = change.value;
          else if (change.value === null) delete next.action; else next.action = materializeAction(change.value);
        }
        inverse.push(Object.freeze({ ...change, expected: change.value, value: change.expected }) as PresentationEditChange);
      } else if (change.type === "slide-reading-order") {
        const slide = byId.get(change.slideId);
        if (!slide) throw new PresentationEditError("missing-target", `Unknown slide id "${change.slideId}".`);
        if (JSON.stringify(slide.readingOrder ?? null) !== JSON.stringify(change.expected)) throw new PresentationEditError("edit-conflict", "Slide reading order has changed.");
        requestedOrders.set(slide, change.value);
        inverse.push(Object.freeze({ type: "slide-reading-order", slideId: change.slideId, expected: change.value, value: change.expected }));
      } else if (change.type === "slide-hidden") {
        const slide = byId.get(change.slideId);
        if (!slide) throw new PresentationEditError("missing-target", `Unknown slide id "${change.slideId}".`);
        if (slide.hidden !== change.expected) throw new PresentationEditError("edit-conflict", `Slide "${change.slideId}" visibility has changed.`);
        staged.push({ slide, value: change.value });
        inverse.push(Object.freeze({ type: "slide-hidden", slideId: change.slideId, expected: change.value, value: change.expected }));
      } else if (change.type === "first-slide-number") {
        if ((this.firstSlideNumberValue ?? null) !== change.expected) throw new PresentationEditError("edit-conflict", "First slide number has changed.");
        nextNumber = change.value === null ? undefined : change.value;
        inverse.push(Object.freeze({ type: "first-slide-number", expected: change.value, value: change.expected }));
      } else {
        if (change.expected.length !== this.slideEntries.length || change.expected.some((id, index) => this.slideEntries[index]!.id !== id)) throw new PresentationEditError("edit-conflict", "Slide order has changed.");
        nextOrder = change.value.map(id => byId.get(id)!);
        inverse.push(Object.freeze({ type: "slide-order", expected: change.value, value: change.expected }));
      }
    }
    const finalOrders = new Map<PresentationSlideImpl, PresentationEditReadingOrder>();
    const stageOrders = (replacements: ReadonlyMap<string, PresentationElement>, editedSlides: ReadonlySet<PresentationSlideImpl>) => {
      for (const slide of new Set([...editedSlides, ...requestedOrders.keys()])) {
        const elements = slide.elements.map(element => replacements.get(element.id) ?? element);
        const original = slide.readingOrder ?? null;
        let next: PresentationEditReadingOrder;
        if (requestedOrders.has(slide)) {
          const requested = requestedOrders.get(slide)!;
          try { next = requested === null ? null : captureSlideReadingOrder(elements, requested); }
          catch { throw new PresentationEditError("invalid-operation", "Reading order does not match the resulting slide elements."); }
        } else next = original === null ? null : reconcileSlideReadingOrder(elements, original) ?? null;
        finalOrders.set(slide, next);
        if (!requestedOrders.has(slide) && JSON.stringify(original) !== JSON.stringify(next)) {
          inverse.push(Object.freeze({ type: "slide-reading-order", slideId: slide.id, expected: next, value: original }));
        }
      }
      if (inverse.length > 10000) throw new PresentationEditError("invalid-operation", "Operation inverse exceeds the 10000 change limit.");
    };
    if (textEdits.size) {
      const edits = [...textEdits.values()];
      const replacements = new Map<string, PresentationElement>();
      const editedSlides = new Set(edits.map(edit => edit.slide));
      this.elementStore.replaceMany(edits.map(edit => edit.original), edits.map(edit => edit.input as PresentationElementInput), elements => {
        for (const element of elements) replacements.set(element.id, element);
        for (const slide of editedSlides) PresentationSlideImpl.prototype.validateEditedElements.call(slide, replacements);
        stageOrders(replacements, editedSlides);
      });
      for (const slide of editedSlides) PresentationSlideImpl.prototype.commitEditedElements.call(slide, replacements);
    } else stageOrders(new Map(), new Set());
    for (const [slide, order] of finalOrders) PresentationSlideImpl.prototype.setReadingOrder.call(slide, order ?? undefined);
    for (const [slide, notes] of notesEdits) PresentationSlideImpl.prototype.setNotes.call(slide, notes);
    // The targets and primitive values have been fully checked; factory-owned setters do not call external code.
    for (const entry of staged) PresentationSlideImpl.prototype.setHidden.call(entry.slide, entry.value);
    this.firstSlideNumberValue = nextNumber;
    if (nextOrder && nextOrder.some((slide, index) => this.slideEntries[index] !== slide)) {
      for (let index = 0; index < nextOrder.length; index++) this.slideEntries[index] = nextOrder[index]!;
      this.slideRevision++;
    }
    const undo = Object.freeze({ schemaVersion: "pptkit-edit-v1" as const, documentId: this.id, changes: Object.freeze(inverse) });
    return Object.freeze({ inverse: undo, changes: operation.changes });
  }

  get metadata(): Readonly<PresentationMetadataInput> {
    return this.metadataValue;
  }

  setMetadata(metadata: PresentationMetadataInput): void {
    if (!metadata || typeof metadata !== "object" || Array.isArray(metadata) || ![Object.prototype, null].includes(Object.getPrototypeOf(metadata)) || Reflect.ownKeys(metadata).some(key => typeof key !== "string" || !Object.prototype.propertyIsEnumerable.call(metadata, key))) throw new TypeError("metadata must be a plain object with enumerable string fields.");
    if (metadata.customProperties !== undefined) assertCustomDocumentProperties(metadata.customProperties);
    const value = deepClone(metadata);
    if (value.customProperties !== undefined) assertCustomDocumentProperties(value.customProperties);
    const stringFields = ["title", "author", "company", "subject", "description", "language"] as const;
    if (Object.keys(value).some(key => ![...stringFields, "keywords", "revision", "customProperties", "createdAt", "modifiedAt"].includes(key as typeof stringFields[number] | "keywords" | "revision" | "customProperties" | "createdAt" | "modifiedAt")) || stringFields.some(key => value[key] !== undefined && typeof value[key] !== "string") || value.keywords !== undefined && (!Array.isArray(value.keywords) || Array.from(value.keywords).some(item => typeof item !== "string")) || value.revision !== undefined && (!Number.isSafeInteger(value.revision) || value.revision < 0) || [value.createdAt, value.modifiedAt].some(date => date !== undefined && !isCanonicalUtcDate(date))) throw new TypeError("metadata contains unsupported fields or values.");
    this.metadataValue = deepFreeze(value);
  }

  get size(): Readonly<PresentationSize> {
    return this.sizeValue;
  }

  get theme(): Readonly<PresentationThemeInput> {
    return this.themeValue;
  }

  get textStylePresets(): TextStylePresetMap {
    return this.textStylePresetsValue;
  }

  get sections(): readonly PresentationSection[] | undefined {
    if (!this.sectionsValue) return undefined;
    const positions = new Map(this.slideEntries.map((slide, index) => [slide.id, index]));
    return Object.freeze(this.sectionsValue.map(section => Object.freeze({ ...section, slideIds: Object.freeze([...section.slideIds].sort((a, b) => positions.get(a)! - positions.get(b)!)) })));
  }

  setSections(sections?: readonly PresentationSection[]): void {
    const revision = this.slideRevision, captured = captureSections(sections, this.slideIds);
    if (revision !== this.slideRevision) throw new Error("Slides or sections changed during section capture.");
    this.sectionsValue = captured; this.slideRevision++;
  }

  setSectionStructure(slideOrder: readonly string[], sections?: readonly PresentationSection[]): void {
    const revision = this.slideRevision;
    const byId = new Map(this.slideEntries.map(slide => [slide.id, slide]));
    if (!Array.isArray(slideOrder) || slideOrder.length !== byId.size || slideOrder.length > 100000 || Reflect.ownKeys(slideOrder).length !== slideOrder.length + 1) throw new TypeError("Slide order must be a complete dense array.");
    const seen = new Set<string>();
    const next: PresentationSlideImpl[] = [];
    for (let index = 0; index < byId.size; index++) {
      if (!Object.hasOwn(slideOrder, index)) throw new TypeError("Slide order must be a complete dense array.");
      const id = slideOrder[index], slide = byId.get(id!);
      if (!slide || seen.has(id!)) throw new TypeError("Slide order must contain each existing slide exactly once.");
      seen.add(id!); next.push(slide);
    }
    const captured = captureSections(sections, new Set(byId.keys()));
    if (revision !== this.slideRevision) throw new Error("Slides or sections changed during section structure capture.");
    for (let index = 0; index < next.length; index++) this.slideEntries[index] = next[index]!;
    this.sectionsValue = captured;
    this.slideRevision++;
  }

  get slides(): readonly ManagedPresentationSlide[] {
    return Object.freeze([...this.slideEntries]);
  }

  get assets(): readonly PresentationAsset[] {
    return this.assetRegistry.assets;
  }

  get layouts(): readonly SlideLayoutDefinition[] {
    return Object.freeze([...this.layoutEntries]);
  }

  addSlide(input: PresentationSlideInput = {}): ManagedPresentationSlide {
    return this.insertSlide(this.slideEntries.length, input);
  }

  insertSlide(index: number, input: PresentationSlideInput = {}): ManagedPresentationSlide {
    return this.commitStagedSlide(index, this.stageSlide(index, input));
  }

  private stageSlide(index: number, input: PresentationSlideInput, state?: { store: ElementStore; allocator: IdAllocator; revision: number }): { slide: PresentationSlideImpl; store: ElementStore; allocator: IdAllocator; revision: number } {
    assertInsertIndex(index, this.slideEntries.length, "Slide");
    const revision = state?.revision ?? this.slideRevision, allocator = state?.allocator ?? this.slideAllocator.fork(), staged = state?.store ?? this.elementStore.forkForAppend();
    const id = input.id ?? allocator.next((candidate) => this.slideIds.has(candidate));
    if (this.slideIds.has(id)) throw new Error(`Duplicate slide id "${id}" detected.`);
    const slide = new PresentationSlideImpl(id, input, staged);
    for (const element of input.elements ?? []) slide.addElement(element);
    if (input.readingOrder !== undefined) slide.setReadingOrder(input.readingOrder);
    if (input.eventTimeline !== undefined) slide.setEventTimeline(input.eventTimeline);
    return { slide, store: staged, allocator, revision };
  }

  private commitStagedSlide(index: number, candidate: { slide: PresentationSlideImpl; store: ElementStore; allocator: IdAllocator; revision: number }): ManagedPresentationSlide {
    const { slide, store, allocator, revision } = candidate;
    if (this.slideRevision !== revision || this.slideIds.has(slide.id)) throw new Error("Slides changed during append transaction.");
    this.elementStore.commitAppend(store, slide.elements);
    slide.bindElementStore(this.elementStore);
    this.slideAllocator = allocator;
    this.slideIds.add(slide.id);
    this.slideEntries.splice(index, 0, slide);
    this.slideRevision++;
    return slide;
  }

  moveSlide(slideId: string, toIndex: number): void {
    const fromIndex = this.slideEntries.findIndex((slide) => slide.id === slideId);
    if (fromIndex < 0) throw new Error(`Unknown slide id "${slideId}".`);
    if (!Number.isInteger(toIndex) || toIndex < 0 || toIndex >= this.slideEntries.length) {
      throw new RangeError(`Slide index ${toIndex} is outside 0..${Math.max(0, this.slideEntries.length - 1)}.`);
    }
    const [slide] = this.slideEntries.splice(fromIndex, 1);
    this.slideEntries.splice(toIndex, 0, slide!);
    this.slideRevision++;
  }

  removeSlide(slideId: string): PresentationSlide {
    const index = this.slideEntries.findIndex((slide) => slide.id === slideId);
    if (index < 0) throw new Error(`Unknown slide id "${slideId}".`);
    const [slide] = this.slideEntries.splice(index, 1);
    this.slideIds.delete(slideId);
    if (this.sectionsValue) this.sectionsValue = Object.freeze(this.sectionsValue.map(section => Object.freeze({ ...section, slideIds: Object.freeze(section.slideIds.filter(id => id !== slideId)) })));
    this.slideRevision++;
    slide!.releaseElements();
    return slide!;
  }

  duplicateSlide(slideId: string, toIndex?: number): ManagedPresentationSlide {
    const index = this.slideEntries.findIndex((slide) => slide.id === slideId);
    if (index < 0) throw new Error(`Unknown slide id "${slideId}".`);
    const source = this.slideEntries[index]!;
    const targetIndex = toIndex ?? index + 1;
    assertInsertIndex(targetIndex, this.slideEntries.length, "Slide");
    const state = { revision: this.slideRevision, allocator: this.slideAllocator.fork(), store: this.elementStore.forkForAppend() };
    const id = state.allocator.next(candidate => this.slideIds.has(candidate));
    const slideIds = new Map([[source.id, id]]);
    const copied = state.store.copyWithIdentityMap(source.elements, slideIds);
    const elements = copied.elements;
    const notes = source.notes === undefined ? undefined : deepClone(source.notes);
    remapTextActions(notes, slideIds);
    const duplicate = this.commitStagedSlide(targetIndex, this.stageSlide(targetIndex, {
      id,
      ...(source.showInheritedShapes !== undefined ? { showInheritedShapes: source.showInheritedShapes } : {}),
      ...(source.layoutId !== undefined ? { layoutId: source.layoutId } : {}),
      ...(source.background !== undefined ? { background: deepClone(source.background) } : {}),
      elements,
      ...(notes !== undefined ? { notes } : {}),
      ...(source.transition !== undefined ? { transition: deepClone(source.transition) } : {}),
      ...(source.readingOrder === undefined ? {} : { readingOrder: remapSlideReadingOrder(source.readingOrder, copied.elementIds) }),
      ...(source.eventTimeline !== undefined ? { eventTimeline: remapSlideEventTimeline(source.eventTimeline, id, copied.elementIds) } : {}),
      hidden: source.hidden,
      ...(source.section !== undefined ? { section: source.section } : {}),
      tags: [...source.tags],
      customData: deepClone(source.customData),
    }, state));
    if (this.sectionsValue) this.sectionsValue = Object.freeze(this.sectionsValue.map(section => section.slideIds.includes(slideId) ? Object.freeze({ ...section, slideIds: Object.freeze([...section.slideIds, duplicate.id]) }) : section));
    return duplicate;
  }

  registerAsset(input: PresentationAssetInput, options?: AssetRegistrationOptions): PresentationAsset {
    return this.assetRegistry.register(input, options);
  }

  registerAssets(requests: readonly AssetRegistrationRequest[]): readonly PresentationAsset[] {
    if (!Array.isArray(requests)) throw new TypeError("Asset registration batch must be an array.");
    if (!requests.length) return Object.freeze([]);
    const original = this.assetRegistry, revision = original.revision, staged = original.fork();
    const results: PresentationAsset[] = [];
    for (const request of requests) {
      if (!request || typeof request !== "object" || Array.isArray(request) || Object.keys(request).some(key => key !== "asset" && key !== "options")) throw new TypeError("Invalid asset registration request.");
      results.push(staged.register(request.asset, request.options));
    }
    original.commitFrom(staged, revision);
    return Object.freeze(results);
  }

  copyAssetsFrom(source: PresentationDocument, assetIds: readonly string[], options: AssetCopyOptions): readonly AssetCopyMapping[] {
    const revision = this.assetRegistry.revision;
    const staged = this.assetRegistry.fork();
    const copied = this.stageAssetCopy(source, assetIds, options, staged, revision);
    this.assetRegistry.commitFrom(staged, revision);
    return copied;
  }

  copySlideFrom(source: PresentationDocument, slideId: string, options: SlideCopyOptions): SlideCopyResult {
    const result = this.copySlidesFromInternal(source, [slideId], options, false);
    return Object.freeze({ slide: result.slides[0]!, elementIds: result.elementIds[slideId]!, assets: result.assets });
  }

  copySlidesFrom(source: PresentationDocument, sourceIds: readonly string[], options: SlidesCopyOptions): SlidesCopyResult {
    return this.copySlidesFromInternal(source, sourceIds, options, true);
  }

  private copySlidesFromInternal(source: PresentationDocument, sourceIds: readonly string[], options: SlidesCopyOptions, allowLayouts: boolean): SlidesCopyResult {
    const state = { revision: this.slideRevision, allocator: this.slideAllocator.fork(), store: this.elementStore.forkForAppend() };
    const definitionRevision = this.definitionRevision;
    const assetRevision = this.assetRegistry.revision, stagedAssets = this.assetRegistry.fork();
    if (!source || typeof source !== "object" || source === this || !Array.isArray(sourceIds) || !sourceIds.length) throw new TypeError("Slide batch copy requires a different source document and nonempty slide IDs.");
    const ids = Array.from(sourceIds);
    const selectedIds = new Set(ids);
    if (ids.some(id => typeof id !== "string" || !id) || selectedIds.size !== ids.length) throw new TypeError("Slide batch copy requires distinct nonempty slide IDs.");
    if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["index", "layoutId", ...(allowLayouts ? ["layouts", "perSlide"] : []), "theme", "geometry", "customData", "assets", "slideLinks", "connectorTargets", "placeholderRoles", "placeholderMatching", "placeholders", "textPresets"].includes(key))) throw new TypeError("Invalid slide copy options.");
    const { index = this.slideEntries.length, layoutId, theme, geometry, customData, assets, slideLinks: slideLinksInput = {}, connectorTargets: connectorTargetsInput = {}, placeholderRoles, placeholderMatching, placeholders: placeholdersInput = {}, textPresets: textPresetsInput = {} } = options;
    const slideLinks = captureCopyMapping(slideLinksInput), connectorTargets = captureCopyMapping(connectorTargetsInput), placeholders = captureCopyMapping(placeholdersInput), textPresets = captureCopyMapping(textPresetsInput);
    const layoutOverrides = allowLayouts ? options.layouts : undefined, pageOverrides = allowLayouts ? options.perSlide : undefined;
    const layouts = layoutOverrides === undefined ? {} : captureDefinitionInput<NonNullable<SlidesCopyOptions["layouts"]>>(layoutOverrides, ids);
    const perSlide = pageOverrides === undefined ? {} : captureDefinitionInput<NonNullable<SlidesCopyOptions["perSlide"]>>(pageOverrides, ids);
    if (theme !== "destination" || (geometry !== "preserve" && geometry !== "fit-explicit-geometry") || (customData !== "preserve" && customData !== "reject") || (layoutId !== null && (typeof layoutId !== "string" || !layoutId))) throw new TypeError("Slide batch copy requires explicit layout, destination theme, geometry and custom data policies.");
    if (placeholderRoles !== undefined && placeholderRoles !== "require-match" && placeholderRoles !== "require-match-and-index") throw new TypeError("Invalid placeholder role policy.");
    if (placeholderMatching !== undefined && (placeholderMatching !== "explicit-index" || placeholderRoles !== "require-match-and-index")) throw new TypeError("Explicit-index placeholder matching requires strict role-and-index policy.");
    assertInsertIndex(index, this.slideEntries.length, "Slide");
    if (!layouts || typeof layouts !== "object" || Array.isArray(layouts) || Object.values(layouts).some(value => value !== null && (typeof value !== "string" || !value))) throw new TypeError("Slide batch layout mappings must contain layout IDs or null.");
    if (!perSlide || typeof perSlide !== "object" || Array.isArray(perSlide)) throw new TypeError("Per-slide mappings must be a record.");
    for (const [id, overrides] of Object.entries(perSlide)) {
      if (!selectedIds.has(id) || !overrides || typeof overrides !== "object" || Array.isArray(overrides) || Object.keys(overrides).some(key => !["connectorTargets", "placeholders", "textPresets"].includes(key))) throw new TypeError("Invalid per-slide mapping.");
      for (const map of [overrides.connectorTargets, overrides.placeholders, overrides.textPresets]) if (map !== undefined && (!map || typeof map !== "object" || Array.isArray(map) || Object.values(map).some(value => typeof value !== "string" || !value))) throw new TypeError("Per-slide mappings must contain nonempty string targets.");
    }
    const sourceSlides = source.slides;
    if (!Array.isArray(sourceSlides)) throw new TypeError("Source slides must be an array.");
    const selected = new Map<string, PresentationSlide>();
    for (const slide of sourceSlides) {
      const id = slide.id;
      if (!selectedIds.has(id)) continue;
      if (selected.has(id)) throw new Error(`Expected one source slide with id "${id}".`);
      selected.set(id, slide);
    }
    const snapshots = ids.map(id => {
      const slide = selected.get(id);
      if (!slide) throw new Error(`Expected one source slide with id "${id}".`);
      const data = slideRebindData(slide);
      const rejectOpaque = () => {
        if (!data.customData || typeof data.customData !== "object" || Array.isArray(data.customData) || ![Object.prototype, null].includes(Object.getPrototypeOf(data.customData))) throw new TypeError("Slide copy requires a plain custom-data object.");
        if (Reflect.ownKeys(data.customData).length) throw new Error(`Slide copy rejected opaque custom data on source slide "${id}".`);
      };
      if (customData === "reject") rejectOpaque();
      const snapshot = deepClone(data);
      if (customData === "reject") rejectOpaque();
      if (snapshot.id !== id) throw new Error("Source slide identity changed during capture.");
      return snapshot;
    });
    const sourceSize = geometry === "fit-explicit-geometry" ? captureDefinitionInput<PresentationSize>(source.size, ["width", "height", "unit"]) : undefined;
    const dependencies = snapshots.map(snapshot => collectSlideCopyDependencies(snapshot));
    if (customData === "reject") {
      const opaque = dependencies.find(inventory => inventory.hasOpaqueCustomData);
      if (opaque) throw new Error(`Slide copy rejected opaque custom data on source slide "${opaque.slideId}".`);
    }
    const sourcePlaceholderRoles = new Map<string, ReadonlyMap<string, { kind: string; index?: number }>>();
    if (placeholderRoles !== undefined && dependencies.some(item => item.placeholders.length)) {
      const sourceLayouts = source.layouts;
      if (!Array.isArray(sourceLayouts)) throw new TypeError("Source layouts must be an array.");
      const sourceLayoutsById = new Map<string, SlideLayoutDefinition[]>();
      for (const layout of sourceLayouts) {
        if (!layout) throw new TypeError("Source layouts require nonempty identities.");
        const id = layout.id;
        if (typeof id !== "string" || !id) throw new TypeError("Source layouts require nonempty identities.");
        const matches = sourceLayoutsById.get(id) ?? [];
        matches.push(layout); sourceLayoutsById.set(id, matches);
      }
      for (const [position, snapshot] of snapshots.entries()) {
        if (!dependencies[position]!.placeholders.length) continue;
        const sourceLayoutId = snapshot.layoutId ?? DEFAULT_LAYOUT_ID;
        if (sourcePlaceholderRoles.has(sourceLayoutId)) continue;
        const matches = sourceLayoutsById.get(sourceLayoutId) ?? [];
        if (matches.length !== 1) throw new Error(`Expected one source layout with id "${sourceLayoutId}" for placeholder role matching.`);
        const captured = captureDefinitionInput<{ id: string; placeholders: readonly { key: string; kind: string; index?: number }[] }>({ id: matches[0]!.id, placeholders: matches[0]!.placeholders }, ["id", "placeholders"]);
        if (captured.id !== sourceLayoutId || !Array.isArray(captured.placeholders)) throw new Error("Source layout changed during placeholder capture.");
        const roles = new Map<string, { kind: string; index?: number }>(), indices = new Set<number>();
        for (const item of captured.placeholders) {
          if (!item || typeof item.key !== "string" || !item.key || !["title", "subtitle", "body", "image", "table", "footer", "slideNumber"].includes(item.kind) || roles.has(item.key)) throw new TypeError("Source layout has invalid or duplicate placeholders.");
          if (placeholderRoles === "require-match-and-index" && item.index !== undefined) {
            if (!isPlaceholderIndex(item.index) || indices.has(item.index)) throw new TypeError("Source layout has invalid or duplicate placeholder indexes.");
            indices.add(item.index);
          }
          roles.set(item.key, { kind: item.kind, ...(item.index === undefined ? {} : { index: item.index }) });
        }
        sourcePlaceholderRoles.set(sourceLayoutId, roles);
      }
    }
    const slideIds = new Map(ids.map(id => [id, state.allocator.next(candidate => this.slideIds.has(candidate))]));
    for (const [position, inventory] of dependencies.entries()) {
      const local = Object.hasOwn(perSlide, ids[position]!) ? perSlide[ids[position]!] : undefined;
      for (const ref of inventory.connectors) if (!ref.slideOwned && !Object.hasOwn(local?.connectorTargets ?? {}, ref.targetId) && !Object.hasOwn(connectorTargets, ref.targetId)) throw new Error(`Slide copy requires a destination mapping for external connector target "${ref.targetId}".`);
      for (const ref of inventory.slideActions) if (!slideIds.has(ref.targetId)) {
        const target = Object.hasOwn(slideLinks, ref.targetId) ? slideLinks[ref.targetId] : undefined;
        if (!target || !this.slideIds.has(target)) throw new Error(`Slide copy requires a destination mapping for slide "${ref.targetId}".`);
        slideIds.set(ref.targetId, target);
      }
    }
    const copiedAssets = this.stageAssetCopy(source, [...new Set(dependencies.flatMap(item => item.assets.map(ref => ref.id)))], assets, stagedAssets, assetRevision);
    const assetIds = new Map(copiedAssets.map(item => [item.sourceId, item.destinationId]));
    const targetLayouts = new Map(this.layoutEntries.map(layout => [layout.id, layout]));
    const targetMasters = new Map(this.masterEntries.map(master => [master.id, master]));
    const targetScopes = new Map<string, ReadonlySet<string>>();
    const targetPlaceholderIndexes = new Map<string, { byKey: Map<string, PlaceholderDefinitionInput[]>; byIndex: Map<number, PlaceholderDefinitionInput[]> }>();
    const entries = snapshots.map((snapshot, position) => {
      const inventory = dependencies[position]!, chosen = Object.hasOwn(layouts, snapshot.id) ? layouts[snapshot.id]! : layoutId;
      const local = Object.hasOwn(perSlide, snapshot.id) ? perSlide[snapshot.id] : undefined;
      if (chosen !== null && (typeof chosen !== "string" || !chosen)) throw new TypeError("Slide batch layout mappings must contain layout IDs or null.");
      const layout = targetLayouts.get(chosen ?? DEFAULT_LAYOUT_ID);
      if (chosen !== null && !layout) throw new Error(`Unknown destination layout id "${chosen}".`);
      const master = layout?.masterId === undefined ? undefined : targetMasters.get(layout.masterId);
      const scopeId = chosen ?? DEFAULT_LAYOUT_ID;
      let sharedTargets = targetScopes.get(scopeId);
      if (!sharedTargets) { sharedTargets = new Set([...(layout?.elements ?? []), ...(master?.elements ?? [])].map(element => element.id)); targetScopes.set(scopeId, sharedTargets); }
      const copied = state.store.copyWithIdentityMap(snapshot.elements, slideIds);
      if (sourceSize) fitCopiedGeometry(copied.elements, sourceSize, this.sizeValue);
      const root = { elements: copied.elements, notes: snapshot.notes === undefined ? undefined : deepClone(snapshot.notes) };
      remapTextActions(root.notes, slideIds);
      for (const ref of inventory.assets) setCopiedReference(root, ref.path, assetIds.get(ref.id)!);
      for (const ref of inventory.connectors) if (!ref.slideOwned) {
        const target = Object.hasOwn(local?.connectorTargets ?? {}, ref.targetId) ? local!.connectorTargets![ref.targetId] : connectorTargets[ref.targetId];
        if (!target || !sharedTargets.has(target)) throw new Error(`Slide copy connector target "${ref.targetId}" must map to a top-level destination layout or master object.`);
        setCopiedReference(root, `${ref.path}.elementId`, target);
      }
      let targetPlaceholderIndex = targetPlaceholderIndexes.get(scopeId);
      if (inventory.placeholders.length && !targetPlaceholderIndex) {
        targetPlaceholderIndex = { byKey: new Map(), byIndex: new Map() };
        for (const placeholder of layout?.placeholders ?? []) {
          const sameKey = targetPlaceholderIndex.byKey.get(placeholder.key) ?? [];
          sameKey.push(placeholder); targetPlaceholderIndex.byKey.set(placeholder.key, sameKey);
          if (placeholder.index !== undefined) {
            const sameIndex = targetPlaceholderIndex.byIndex.get(placeholder.index) ?? [];
            sameIndex.push(placeholder); targetPlaceholderIndex.byIndex.set(placeholder.index, sameIndex);
          }
        }
        targetPlaceholderIndexes.set(scopeId, targetPlaceholderIndex);
      }
      for (const ref of inventory.placeholders) {
        const sourceRole = sourcePlaceholderRoles.get(snapshot.layoutId ?? DEFAULT_LAYOUT_ID)?.get(ref.key);
        let target: string;
        if (Object.hasOwn(local?.placeholders ?? {}, ref.key)) target = local!.placeholders![ref.key]!;
        else if (Object.hasOwn(placeholders, ref.key)) target = placeholders[ref.key]!;
        else if (placeholderMatching === "explicit-index") {
          if (!sourceRole) throw new Error(`Slide copy requires a source placeholder for "${ref.key}".`);
          if (!isPlaceholderIndex(sourceRole.index)) throw new Error(`Slide copy requires an explicit valid source placeholder index for "${ref.key}".`);
          const matches = targetPlaceholderIndex!.byIndex.get(sourceRole.index) ?? [];
          if (matches.length !== 1) throw new Error(`Slide copy requires one destination placeholder with index ${sourceRole.index} for "${ref.key}".`);
          if (typeof matches[0]!.key !== "string" || !matches[0]!.key.trim()) throw new TypeError("Destination placeholder key must be a nonempty string.");
          target = matches[0]!.key;
        } else target = ref.key;
        const targetMatches = targetPlaceholderIndex!.byKey.get(target) ?? [];
        const destinationPlaceholder = targetMatches[0];
        if (!destinationPlaceholder) throw new Error(`Slide copy requires a destination placeholder for "${ref.key}".`);
        if (placeholderRoles !== undefined) {
          if (targetMatches.length !== 1) throw new Error(`Slide copy requires one destination placeholder for "${target}".`);
          if (!sourceRole) throw new Error(`Slide copy requires a source placeholder for "${ref.key}".`);
          if (sourceRole.kind !== destinationPlaceholder.kind) throw new Error(`Slide copy placeholder role mismatch for "${ref.key}": ${sourceRole.kind} to ${destinationPlaceholder.kind}.`);
          if (placeholderRoles === "require-match-and-index") {
            if (!isPlaceholderIndex(sourceRole.index) || !isPlaceholderIndex(destinationPlaceholder.index)) throw new Error(`Slide copy requires explicit valid placeholder indexes for "${ref.key}".`);
            if ((targetPlaceholderIndex!.byIndex.get(destinationPlaceholder.index)?.length ?? 0) !== 1) throw new Error(`Slide copy requires a unique destination placeholder index for "${target}".`);
            if (sourceRole.index !== destinationPlaceholder.index) throw new Error(`Slide copy placeholder index mismatch for "${ref.key}": ${sourceRole.index} to ${destinationPlaceholder.index}.`);
          }
        }
        setCopiedReference(root, ref.path, target);
      }
      for (const ref of inventory.textPresets) {
        const target = Object.hasOwn(local?.textPresets ?? {}, ref.name) ? local!.textPresets![ref.name]! : Object.hasOwn(textPresets, ref.name) ? textPresets[ref.name]! : ref.name;
        if (!Object.hasOwn(this.textStylePresetsValue, target)) throw new Error(`Slide copy requires a destination text preset for "${ref.name}".`);
        setCopiedReference(root, ref.path, target);
      }
      const id = slideIds.get(snapshot.id)!;
      const candidate = this.stageSlide(index, { id, ...(chosen === null ? {} : { layoutId: chosen }), ...(snapshot.showInheritedShapes !== undefined ? { showInheritedShapes: snapshot.showInheritedShapes } : {}), ...(snapshot.background !== undefined ? { background: snapshot.background } : {}), elements: copied.elements, ...(root.notes !== undefined ? { notes: root.notes } : {}), ...(snapshot.transition !== undefined ? { transition: snapshot.transition } : {}), ...(snapshot.readingOrder === undefined ? {} : { readingOrder: remapSlideReadingOrder(snapshot.readingOrder, copied.elementIds) }), ...(snapshot.eventTimeline !== undefined ? { eventTimeline: remapSlideEventTimeline(snapshot.eventTimeline, id, copied.elementIds) } : {}), hidden: snapshot.hidden, ...(snapshot.section !== undefined ? { section: snapshot.section } : {}), tags: [...snapshot.tags], customData: deepClone(snapshot.customData) }, state);
      return { slide: candidate.slide, elementIds: Object.freeze(Object.fromEntries(copied.elementIds)) };
    });
    const candidateIds = new Set(entries.map(entry => entry.slide.id)), copiedAssetIds = new Set(copiedAssets.map(item => item.destinationId));
    const diagnostics = validatePresentation({ masters: this.masterEntries, layouts: this.layoutEntries, slides: [...this.slideEntries, ...entries.map(entry => entry.slide)], assets: stagedAssets.assets, textStylePresets: this.textStylePresetsValue, theme: this.themeValue }).filter(diagnostic => diagnostic.severity === "error" && ((diagnostic.slideId !== undefined && candidateIds.has(diagnostic.slideId)) || (diagnostic.assetId !== undefined && copiedAssetIds.has(diagnostic.assetId))));
    if (diagnostics.length) throw new PresentationValidationError(diagnostics);
    if (this.slideRevision !== state.revision || this.definitionRevision !== definitionRevision || entries.some(entry => this.slideIds.has(entry.slide.id))) throw new Error("Document changed during slide copy.");
    if (this.assetRegistry.revision !== assetRevision) throw new Error("Asset registry changed during slide copy.");
    const trailing = this.slideEntries.slice(index);
    this.elementStore.commitAppend(state.store, entries.flatMap(entry => [...entry.slide.elements]));
    this.assetRegistry.commitFrom(stagedAssets, assetRevision);
    this.slideAllocator = state.allocator;
    this.slideEntries.length = index;
    for (const entry of entries) { entry.slide.bindElementStore(this.elementStore); this.slideIds.add(entry.slide.id); this.slideEntries.push(entry.slide); }
    for (const slide of trailing) this.slideEntries.push(slide);
    this.slideRevision++;
    return Object.freeze({ slides: Object.freeze(entries.map(entry => entry.slide)), slideIds: Object.freeze(Object.fromEntries(ids.map(id => [id, slideIds.get(id)!]))), elementIds: Object.freeze(Object.fromEntries(ids.map((id, position) => [id, entries[position]!.elementIds]))), assets: copiedAssets });
  }

  copySlideLayoutFrom(source: PresentationDocument, layoutId: string, options: LayoutCopyOptions): LayoutCopyResult {
    const slideRevision = this.slideRevision, stagedStore = this.elementStore.forkForAppend();
    const definitionRevision = this.definitionRevision;
    const assetRevision = this.assetRegistry.revision, stagedAssets = this.assetRegistry.fork();
    if (!source || typeof source !== "object" || source === this || typeof layoutId !== "string" || !layoutId) throw new TypeError("Layout copy requires a different source document and a layout ID.");
    if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["id", "masterId", "theme", "assets", "slideLinks", "connectorTargets", "textPresets", "reuseEquivalent"].includes(key))) throw new TypeError("Invalid layout copy options.");
    const { id, masterId, theme, assets, reuseEquivalent = false, slideLinks: slideLinksInput = {}, connectorTargets: connectorTargetsInput = {}, textPresets: textPresetsInput = {} } = options;
    if (typeof reuseEquivalent !== "boolean") throw new TypeError("Layout reuse policy must be a boolean.");
    const slideLinks = captureCopyMapping(slideLinksInput), connectorTargets = captureCopyMapping(connectorTargetsInput), textPresets = captureCopyMapping(textPresetsInput);
    if (typeof id !== "string" || !id.trim() || theme !== "destination" || (masterId !== null && (typeof masterId !== "string" || !masterId))) throw new TypeError("Layout copy requires an explicit ID, destination theme and master policy.");
    if (this.layoutEntries.some(layout => layout.id === id)) throw new Error(`Duplicate slide layout id "${id}" detected.`);
    if (masterId !== null && !this.masterEntries.some(master => master.id === masterId)) throw new Error(`Unknown destination master id "${masterId}".`);
    const master = masterId === null ? undefined : this.masterEntries.find(item => item.id === masterId);
    const masterTargets = new Set((master?.elements ?? []).map(element => element.id));
    const sourceLayouts = source.layouts;
    if (!Array.isArray(sourceLayouts)) throw new TypeError("Source layouts must be an array.");
    const matches = sourceLayouts.filter(layout => layout.id === layoutId);
    if (matches.length !== 1) throw new Error(`Expected one source layout with id "${layoutId}".`);
    const snapshot = captureDefinitionInput<SlideLayoutDefinition>(matches[0]!, ["id", "name", "masterId", "background", "elements", "placeholders", "showMasterShapes"]);
    if (snapshot.id !== layoutId) throw new Error("Source layout identity changed during capture.");
    assertDefinitionIdentity(id, snapshot.name);
    const dependencies = collectSlideCopyDependencies({ id: layoutId, layoutId: undefined, showInheritedShapes: undefined, background: undefined, elements: snapshot.elements, notes: undefined, transition: undefined, eventTimeline: undefined, hidden: false, section: undefined, tags: [], customData: {} });
    for (const ref of dependencies.connectors) if (!ref.slideOwned && !Object.hasOwn(connectorTargets, ref.targetId)) throw new Error(`Layout copy requires a destination mapping for external connector target "${ref.targetId}".`);
    const slideIds = new Map<string, string>();
    for (const ref of dependencies.slideActions) {
      const target = Object.hasOwn(slideLinks, ref.targetId) ? slideLinks[ref.targetId] : undefined;
      if (!target || !this.slideIds.has(target)) throw new Error(`Layout copy requires a destination mapping for slide "${ref.targetId}".`);
      slideIds.set(ref.targetId, target);
    }
    const copiedAssets = this.stageAssetCopy(source, [...new Set(dependencies.assets.map(ref => ref.id))], assets, stagedAssets, assetRevision);
    const assetIds = new Map(copiedAssets.map(item => [item.sourceId, item.destinationId]));
    const copied = stagedStore.copyWithIdentityMap(snapshot.elements, slideIds), root = { elements: copied.elements };
    for (const ref of dependencies.assets) setCopiedReference(root, ref.path, assetIds.get(ref.id)!);
    for (const ref of dependencies.connectors) if (!ref.slideOwned) {
      const target = connectorTargets[ref.targetId];
      if (!target || !masterTargets.has(target)) throw new Error(`Layout copy connector target "${ref.targetId}" must map to a top-level destination master object.`);
      setCopiedReference(root, `${ref.path}.elementId`, target);
    }
    for (const ref of dependencies.textPresets) {
      const target = Object.hasOwn(textPresets, ref.name) ? textPresets[ref.name]! : ref.name;
      if (!Object.hasOwn(this.textStylePresetsValue, target)) throw new Error(`Layout copy requires a destination text preset for "${ref.name}".`);
      setCopiedReference(root, ref.path, target);
    }
    const elements = copied.elements.map(element => stagedStore.materialize(element));
    const copiedPlaceholders = snapshot.placeholders.map(placeholder => {
      if (placeholder.textStylePreset === undefined) return placeholder;
      const name = placeholder.textStylePreset;
      const target = Object.hasOwn(textPresets, name) ? textPresets[name]! : name;
      if (!Object.hasOwn(this.textStylePresetsValue, target)) throw new Error(`Layout copy requires a destination text preset for "${name}".`);
      return { ...placeholder, textStylePreset: target };
    });
    const candidate: SlideLayoutDefinition = deepFreeze({ id, name: snapshot.name, ...(masterId === null ? {} : { masterId }), ...(snapshot.showMasterShapes === undefined ? {} : { showMasterShapes: snapshot.showMasterShapes }), ...(snapshot.background === undefined ? {} : { background: snapshot.background }), elements, placeholders: copiedPlaceholders });
    const copiedAssetIds = new Set(copiedAssets.map(item => item.destinationId)), index = this.layoutEntries.length;
    const affectedSlides = new Set(this.slideEntries.filter(slide => (slide.layoutId ?? DEFAULT_LAYOUT_ID) === id).map(slide => slide.id));
    const diagnostics = validatePresentation({ masters: this.masterEntries, layouts: [...this.layoutEntries, candidate], slides: this.slideEntries, assets: stagedAssets.assets, textStylePresets: this.textStylePresetsValue, theme: this.themeValue }).filter(diagnostic => diagnostic.severity === "error" && (diagnostic.path?.startsWith(`layouts.${index}.`) || (diagnostic.slideId !== undefined && affectedSlides.has(diagnostic.slideId)) || (diagnostic.assetId !== undefined && copiedAssetIds.has(diagnostic.assetId))));
    if (diagnostics.length) throw new PresentationValidationError(diagnostics);
    if (this.slideRevision !== slideRevision || this.definitionRevision !== definitionRevision || this.layoutEntries.some(layout => layout.id === id) || (masterId !== null && !this.masterEntries.some(master => master.id === masterId))) throw new Error("Document changed during layout copy.");
    if (this.assetRegistry.revision !== assetRevision) throw new Error("Asset registry changed during layout copy.");
    if (reuseEquivalent && copiedAssets.every(asset => asset.reused)) {
      for (const existing of this.layoutEntries) {
        const mapped = equivalentDefinitionElements(candidate, existing);
        if (mapped) {
          const elementIds = Object.fromEntries([...copied.elementIds].map(([sourceId, candidateId]) => [sourceId, mapped.get(candidateId)!]));
          return Object.freeze({ layout: existing, elementIds: Object.freeze(elementIds), assets: copiedAssets });
        }
      }
    }
    this.elementStore.commitAppend(stagedStore, elements);
    this.assetRegistry.commitFrom(stagedAssets, assetRevision);
    this.layoutEntries.push(candidate);
    this.definitionRevision++;
    return Object.freeze({ layout: candidate, elementIds: Object.freeze(Object.fromEntries(copied.elementIds)), assets: copiedAssets });
  }

  copySlideMasterFrom(source: PresentationDocument, masterId: string, options: MasterCopyOptions): MasterCopyResult {
    const slideRevision = this.slideRevision, stagedStore = this.elementStore.forkForAppend();
    const definitionRevision = this.definitionRevision;
    const assetRevision = this.assetRegistry.revision, stagedAssets = this.assetRegistry.fork();
    if (!source || typeof source !== "object" || source === this || typeof masterId !== "string" || !masterId) throw new TypeError("Master copy requires a different source document and a master ID.");
    if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => !["id", "theme", "assets", "slideLinks", "textPresets", "reuseEquivalent"].includes(key))) throw new TypeError("Invalid master copy options.");
    const { id, theme, assets, reuseEquivalent = false, slideLinks: slideLinksInput = {}, textPresets: textPresetsInput = {} } = options;
    if (typeof reuseEquivalent !== "boolean") throw new TypeError("Master reuse policy must be a boolean.");
    const slideLinks = captureCopyMapping(slideLinksInput), textPresets = captureCopyMapping(textPresetsInput);
    if (typeof id !== "string" || !id.trim() || (theme !== "destination" && theme !== "source-declarations" && theme !== "source-effective")) throw new TypeError("Master copy requires an explicit ID and theme policy.");
    if (this.masterEntries.some(master => master.id === id)) throw new Error(`Duplicate master id "${id}" detected.`);
    const sourceMasters = source.masters;
    if (!Array.isArray(sourceMasters)) throw new TypeError("Source masters must be an array.");
    const matches = sourceMasters.filter(master => master.id === masterId);
    if (matches.length !== 1) throw new Error(`Expected one source master with id "${masterId}".`);
    const snapshot = captureDefinitionInput<SlideMasterDefinition>(matches[0]!, ["id", "name", "theme", "background", "elements"]);
    if (snapshot.id !== masterId) throw new Error("Source master identity changed during capture.");
    const sourceElements = snapshot.elements ?? [];
    const dependencies = collectSlideCopyDependencies({ id: masterId, layoutId: undefined, showInheritedShapes: undefined, background: undefined, elements: sourceElements, notes: undefined, transition: undefined, eventTimeline: undefined, hidden: false, section: undefined, tags: [], customData: {} });
    if (dependencies.connectors.some(ref => !ref.slideOwned)) throw new Error("Master copy has an external connector target.");
    const slideIds = new Map<string, string>();
    for (const ref of dependencies.slideActions) {
      const target = Object.hasOwn(slideLinks, ref.targetId) ? slideLinks[ref.targetId] : undefined;
      if (!target || !this.slideIds.has(target)) throw new Error(`Master copy requires a destination mapping for slide "${ref.targetId}".`);
      slideIds.set(ref.targetId, target);
    }
    const copiedAssets = this.stageAssetCopy(source, [...new Set(dependencies.assets.map(ref => ref.id))], assets, stagedAssets, assetRevision);
    const assetIds = new Map(copiedAssets.map(item => [item.sourceId, item.destinationId]));
    const copied = stagedStore.copyWithIdentityMap(sourceElements, slideIds), root = { elements: copied.elements };
    for (const ref of dependencies.assets) setCopiedReference(root, ref.path, assetIds.get(ref.id)!);
    for (const ref of dependencies.textPresets) {
      const target = Object.hasOwn(textPresets, ref.name) ? textPresets[ref.name]! : ref.name;
      if (!Object.hasOwn(this.textStylePresetsValue, target)) throw new Error(`Master copy requires a destination text preset for "${ref.name}".`);
      setCopiedReference(root, ref.path, target);
    }
    const elements = copied.elements.map(element => stagedStore.materialize(element));
    const effectiveTheme = theme === "source-effective" ? resolveTheme(snapshot.theme, resolveTheme(captureDefinitionInput<PresentationThemeInput>(source.theme, ["name", "colors", "fonts"]), defaultTheme(), { kind: "document", id: source.id, path: "theme" }), { kind: "master", id: masterId, path: "theme" }).value : undefined;
    const candidate: SlideMasterDefinition = deepFreeze({ id, name: snapshot.name, ...(effectiveTheme !== undefined ? { theme: effectiveTheme } : theme === "source-declarations" && snapshot.theme !== undefined ? { theme: snapshot.theme } : {}), ...(snapshot.background === undefined ? {} : { background: snapshot.background }), ...(snapshot.elements === undefined ? {} : { elements }) });
    const copiedAssetIds = new Set(copiedAssets.map(item => item.destinationId)), index = this.masterEntries.length;
    const affectedLayouts = new Set(this.layoutEntries.filter(layout => layout.masterId === id).map(layout => layout.id));
    const affectedSlides = new Set(this.slideEntries.filter(slide => affectedLayouts.has(slide.layoutId ?? DEFAULT_LAYOUT_ID)).map(slide => slide.id));
    const diagnostics = validatePresentation({ masters: [...this.masterEntries, candidate], layouts: this.layoutEntries, slides: this.slideEntries, assets: stagedAssets.assets, textStylePresets: this.textStylePresetsValue, theme: this.themeValue }).filter(diagnostic => diagnostic.severity === "error" && (diagnostic.path?.startsWith(`masters.${index}.`) || (diagnostic.layoutId !== undefined && affectedLayouts.has(diagnostic.layoutId)) || (diagnostic.slideId !== undefined && affectedSlides.has(diagnostic.slideId)) || (diagnostic.assetId !== undefined && copiedAssetIds.has(diagnostic.assetId))));
    if (diagnostics.length) throw new PresentationValidationError(diagnostics);
    if (this.slideRevision !== slideRevision || this.definitionRevision !== definitionRevision || this.masterEntries.some(master => master.id === id)) throw new Error("Document changed during master copy.");
    if (this.assetRegistry.revision !== assetRevision) throw new Error("Asset registry changed during master copy.");
    if (reuseEquivalent && copiedAssets.every(asset => asset.reused)) {
      for (const existing of this.masterEntries) {
        const mapped = equivalentDefinitionElements(candidate, existing);
        if (mapped) {
          const elementIds = Object.fromEntries([...copied.elementIds].map(([sourceId, candidateId]) => [sourceId, mapped.get(candidateId)!]));
          return Object.freeze({ master: existing, elementIds: Object.freeze(elementIds), assets: copiedAssets });
        }
      }
    }
    this.elementStore.commitAppend(stagedStore, elements);
    this.assetRegistry.commitFrom(stagedAssets, assetRevision);
    this.masterEntries = Object.freeze([...this.masterEntries, candidate]);
    this.definitionRevision++;
    return Object.freeze({ master: candidate, elementIds: Object.freeze(Object.fromEntries(copied.elementIds)), assets: copiedAssets });
  }

  private stageAssetCopy(source: PresentationDocument, assetIds: readonly string[], options: AssetCopyOptions, staged: AssetRegistry, revision: number): readonly AssetCopyMapping[] {
    if (!source || typeof source !== "object" || source === this) throw new TypeError("Asset copy requires a different source document.");
    if (!Array.isArray(assetIds)) throw new TypeError("Asset copy requires distinct nonempty source IDs.");
    const ids = Array.from(assetIds);
    if (ids.some(id => typeof id !== "string" || !id) || new Set(ids).size !== ids.length) throw new TypeError("Asset copy requires distinct nonempty source IDs.");
    if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => key !== "deduplicate" && key !== "external")) throw new TypeError("Asset copy requires explicit deduplication and external-source policies.");
    const { deduplicate, external } = options;
    if ((deduplicate !== "registry" && deduplicate !== "content") || (external !== "reject" && external !== "preserve")) throw new TypeError("Invalid asset copy policy.");
    const assets = source.assets;
    if (!Array.isArray(assets)) throw new TypeError("Source assets must be an array.");
    const byId = new Map<string, PresentationAsset>();
    for (const asset of assets) {
      if (!asset || typeof asset !== "object") throw new TypeError("Source assets must be objects.");
      const id = asset.id;
      if (typeof id !== "string" || !id || byId.has(id)) throw new TypeError("Source asset identities must be unique nonempty strings.");
      byId.set(id, asset);
    }
    const requests: AssetRegistrationRequest[] = [];
    for (const id of ids) {
      const found = byId.get(id);
      if (!found) throw new Error(`Unknown source asset id "${id}".`);
      assertCopyAssetFields(found, ["id", "kind", "source", "mimeType", "width", "height", "accessibility", "dedupeKey"]);
      const fields = { ...found };
      if (fields.source && typeof fields.source === "object") assertCopyAssetFields(fields.source, ["type", "value", "encoding"]);
      if (fields.accessibility && typeof fields.accessibility === "object") assertCopyAssetFields(fields.accessibility, ["description", "decorative"]);
      const asset = deepClone(fields);
      // Getters read during capture can add fields that structuredClone omits.
      assertCopyAssetFields(found, ["id", "kind", "source", "mimeType", "width", "height", "accessibility", "dedupeKey"]);
      if (fields.source && typeof fields.source === "object") assertCopyAssetFields(fields.source, ["type", "value", "encoding"]);
      if (fields.accessibility && typeof fields.accessibility === "object") assertCopyAssetFields(fields.accessibility, ["description", "decorative"]);
      if (asset.id !== id) throw new Error("Source asset identity changed during capture.");
      if (Object.keys(asset).some(key => !["id", "kind", "source", "mimeType", "width", "height", "accessibility", "dedupeKey"].includes(key)) || !["image", "audio", "video"].includes(asset.kind) || !asset.source || typeof asset.source !== "object" || Object.keys(asset.source).some(key => !["type", "value", "encoding"].includes(key)) || !["bytes", "path", "url"].includes(asset.source.type) || typeof asset.source.value !== "string" || !asset.source.value.trim() || (asset.source.type !== "bytes" && "encoding" in asset.source) || (asset.accessibility !== undefined && (!asset.accessibility || typeof asset.accessibility !== "object" || Object.keys(asset.accessibility).some(key => key !== "description" && key !== "decorative")))) throw new TypeError("Source asset has unsupported fields or source type.");
      if (asset.source.type !== "bytes" && external === "reject") throw new Error(`External source for asset "${id}" requires explicit preserve policy.`);
      const { id: _sourceId, ...input } = asset;
      requests.push({ asset: input, ...(deduplicate === "content" ? { options: { deduplicate: "content" } } : {}) });
    }
    if (this.assetRegistry.revision !== revision) throw new Error("Asset registry changed while capturing source assets.");
    const seen = new Set(staged.assets.map(asset => asset.id));
    const copied = requests.map(request => staged.register(request.asset, request.options));
    return Object.freeze(copied.map((asset, index) => {
      const reused = seen.has(asset.id); seen.add(asset.id);
      return Object.freeze({ sourceId: ids[index]!, destinationId: asset.id, reused });
    }));
  }

  replaceAsset(assetId: string, replacement: PresentationAssetReplacement): PresentationAsset {
    return this.assetRegistry.replace(assetId, replacement);
  }

  getAssetReferences(assetId: string): readonly AssetReference[] {
    return Object.freeze(collectAssetReferences(this.slideEntries, this.layoutEntries, this.masterEntries).get(assetId) ?? []);
  }

  private retainedAssets(options: AssetRetentionOptions = {}): Set<string> {
    if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => key !== "retainAssetIds")) throw new TypeError("Invalid asset retention options.");
    const ids = options.retainAssetIds;
    if (ids !== undefined && (!Array.isArray(ids) || ids.some(id => typeof id !== "string"))) throw new TypeError("retainAssetIds must be an array of asset IDs.");
    const retained = new Set(ids ?? []);
    for (const id of retained) if (!this.assetRegistry.get(id)) throw new Error(`Unknown retained asset id "${id}".`);
    return retained;
  }

  removeAsset(assetId: string, options?: AssetRetentionOptions): PresentationAsset {
    const retained = this.retainedAssets(options);
    if (!this.assetRegistry.get(assetId)) throw new Error(`Unknown asset id "${assetId}".`);
    if (retained.has(assetId)) throw new Error(`Asset "${assetId}" is explicitly retained.`);
    const references = this.getAssetReferences(assetId);
    if (references.length) throw new Error(`Asset "${assetId}" is referenced at ${references[0]!.path}; remove its uses before deleting it.`);
    return this.assetRegistry.remove(assetId);
  }

  removeUnusedAssets(options?: AssetRetentionOptions): readonly PresentationAsset[] {
    const retained = this.retainedAssets(options);
    const referenced = collectAssetReferences(this.slideEntries, this.layoutEntries, this.masterEntries);
    for (const id of referenced.keys()) retained.add(id);
    return this.assetRegistry.removeExcept(retained);
  }

  getAsset(assetId: string): PresentationAsset | undefined {
    return this.assetRegistry.get(assetId);
  }

  applyLayoutRebind(plan: LayoutRebindPlan): LayoutRebindApplyResult {
    if (!plan || typeof plan !== "object" || Array.isArray(plan) || ![Object.prototype, null].includes(Object.getPrototypeOf(plan))) throw new LayoutRebindError("invalid-plan", "A layout rebind value plan is required.");
    const descriptors = Object.getOwnPropertyDescriptors(plan as object);
    for (const key of ["sourceSnapshot", "requests", "options"]) if (!descriptors[key] || !("value" in descriptors[key]!)) throw new LayoutRebindError("invalid-plan", "Plan inputs must be data properties.");
    if (typeof descriptors.sourceSnapshot!.value !== "string") throw new LayoutRebindError("invalid-plan", "A source snapshot is required.");
    const stage = stageLayoutRebind(this, descriptors.requests!.value, descriptors.options!.value, descriptors.sourceSnapshot!.value);
    if (stage.plan.sourceSnapshot !== descriptors.sourceSnapshot!.value) throw new LayoutRebindError("stale-plan", "Authoring data changed after the layout rebind was prepared.");
    if (stage.plan.status !== "ready") return Object.freeze({ status: "needs-review", plan: stage.plan });
    const changed = new Set(stage.plan.changedSlideIds);
    const selected = this.slideEntries.filter(slide => changed.has(slide.id));
    if (!selected.length) return Object.freeze({ status: "applied", changedSlideIds: stage.plan.changedSlideIds });
    const staged = new Map(stage.slides.map(item => [item.slideId, item]));
    let replacements = new Map<string, readonly import("../types/element.js").PresentationElement[]>();
    this.elementStore.replaceMany(selected.flatMap(slide => [...slide.elements]), selected.flatMap(slide => [...staged.get(slide.id)!.elements]), elements => {
      let offset = 0;
      replacements = new Map(selected.map(slide => { const values = elements.slice(offset, offset + slide.elements.length); offset += slide.elements.length; return [slide.id, values]; }));
      const candidate = this.slideEntries.map(slide => {
        const values = replacements.get(slide.id), item = staged.get(slide.id);
        return values && item ? { ...slideRebindData(slide), layoutId: item.layoutId, background: item.background, elements: values } : slide;
      });
      const diagnostics = validatePresentation({ masters: this.masterEntries, layouts: this.layoutEntries, slides: candidate, assets: this.assets, textStylePresets: this.textStylePresets, theme: this.theme });
      if (diagnostics.some(diagnostic => diagnostic.severity === "error")) throw new PresentationValidationError(diagnostics);
    });
    for (const slide of selected) { const item = staged.get(slide.id)!; slide.commitLayoutRebind(item.layoutId, item.background, replacements.get(slide.id)!); }
    return Object.freeze({ status: "applied", changedSlideIds: stage.plan.changedSlideIds });
  }

  defineSlideMaster(input: SlideMasterInput): SlideMasterDefinition {
    const captured = captureDefinitionInput(input, ["id", "name", "theme", "background", "elements"]);
    assertDefinitionIdentity(captured.id, captured.name);
    if (this.masterEntries.some(master => master.id === captured.id)) throw new Error(`Duplicate master id "${captured.id}".`);
    return this.commitMaster(captured);
  }

  replaceSlideMaster(masterId: string, input: Omit<SlideMasterInput, "id"> & { id?: string }): SlideMasterDefinition {
    const index = this.masterEntries.findIndex(master => master.id === masterId);
    if (index < 0) throw new Error(`Unknown master id "${masterId}".`);
    const captured = captureDefinitionInput(input, ["id", "name", "theme", "background", "elements"]);
    if (captured.id !== undefined && captured.id !== masterId) throw new Error("replaceSlideMaster must retain the master ID.");
    assertDefinitionIdentity(masterId, captured.name);
    return this.commitMaster({ ...captured, id: masterId }, index);
  }

  private commitMaster(input: SlideMasterInput, index?: number): SlideMasterDefinition {
    if (input.elements !== undefined && !Array.isArray(input.elements)) throw new TypeError("Master elements must be an array.");
    assertDefinitionElements(input.elements ?? []);
    const original = index === undefined ? undefined : this.masterEntries[index];
    let replacement: SlideMasterDefinition | undefined;
    this.elementStore.replaceMany(original?.elements ?? [], input.elements ?? [], elements => {
      const { elements: _inputElements, ...definition } = input;
      const master = deepFreeze({ ...definition, ...(input.elements === undefined ? {} : { elements }) });
      const masters = index === undefined ? [...this.masterEntries, master] : this.masterEntries.map((entry, i) => i === index ? master : entry);
      const layouts = new Set(this.layoutEntries.filter(layout => layout.masterId === input.id).map(layout => layout.id));
      const slides = new Set(this.slideEntries.filter(slide => layouts.has(slide.layoutId ?? DEFAULT_LAYOUT_ID)).map(slide => slide.id));
      const diagnostics = validatePresentation({ masters, layouts: this.layoutEntries, slides: this.slideEntries, assets: this.assets, textStylePresets: this.textStylePresets, theme: this.theme })
        .filter(diagnostic => diagnostic.masterId === input.id || (diagnostic.layoutId !== undefined && layouts.has(diagnostic.layoutId)) || (diagnostic.slideId !== undefined && slides.has(diagnostic.slideId)));
      if (diagnostics.some(diagnostic => diagnostic.severity === "error")) throw new PresentationValidationError(diagnostics);
      replacement = master;
    });
    this.masterEntries = Object.freeze(index === undefined ? [...this.masterEntries, replacement!] : this.masterEntries.map((entry, i) => i === index ? replacement! : entry));
    this.definitionRevision++;
    return replacement!;
  }

  replaceSlideLayout(layoutId: string, input: Omit<SlideLayoutInput, "id"> & { id?: string }): SlideLayoutDefinition {
    const index = this.layoutEntries.findIndex(layout => layout.id === layoutId);
    if (index < 0) throw new Error(`Unknown slide layout id "${layoutId}".`);
    const captured = captureDefinitionInput(input, ["id", "name", "masterId", "background", "elements", "placeholders", "showMasterShapes"]);
    if (captured.id !== undefined && captured.id !== layoutId) throw new Error("replaceSlideLayout must retain the layout ID.");
    assertDefinitionIdentity(layoutId, captured.name);
    if ((captured.elements !== undefined && !Array.isArray(captured.elements)) || (captured.placeholders !== undefined && !Array.isArray(captured.placeholders))) throw new TypeError("Layout elements and placeholders must be arrays.");
    assertDefinitionElements(captured.elements ?? []);
    let replacement: SlideLayoutDefinition | undefined;
    this.elementStore.replaceMany(this.layoutEntries[index]!.elements, captured.elements ?? [], elements => {
      const candidate = deepFreeze({ ...captured, id: layoutId, elements, placeholders: captured.placeholders ?? [] });
      const layouts = this.layoutEntries.map((layout, i) => i === index ? candidate : layout);
      const affectedSlides = new Set(this.slideEntries.filter(slide => (slide.layoutId ?? DEFAULT_LAYOUT_ID) === layoutId).map(slide => slide.id));
      const diagnostics = validatePresentation({ layouts, slides: this.slideEntries, masters: this.masterEntries, assets: this.assets, textStylePresets: this.textStylePresets, theme: this.theme })
        .filter(diagnostic => diagnostic.layoutId === layoutId || (diagnostic.slideId !== undefined && affectedSlides.has(diagnostic.slideId)));
      if (diagnostics.some(diagnostic => diagnostic.severity === "error")) throw new PresentationValidationError(diagnostics);
      replacement = candidate;
    });
    this.layoutEntries[index] = replacement!;
    this.definitionRevision++;
    return replacement!;
  }

  defineSlideLayout(input: SlideLayoutInput): SlideLayoutDefinition {
    if (this.layoutEntries.some((layout) => layout.id === input.id)) {
      throw new Error(`Duplicate slide layout id "${input.id}" detected.`);
    }
    const elements: ReturnType<ElementStore["materialize"]>[] = [];
    try {
      for (const element of input.elements ?? []) elements.push(this.elementStore.materialize(element));
      const placeholders = deepFreeze(deepClone(input.placeholders ?? []));
      const layout = deepFreeze({
        id: input.id,
        name: input.name,
        ...(input.showMasterShapes !== undefined ? { showMasterShapes: input.showMasterShapes } : {}),
        ...(input.masterId !== undefined ? { masterId: input.masterId } : {}),
        ...(input.background !== undefined ? { background: deepClone(input.background) } : {}),
        elements,
        placeholders,
      }) as SlideLayoutDefinition;
      this.layoutEntries.push(layout);
      this.definitionRevision++;
      return layout;
    } catch (error) {
      for (const element of elements) this.elementStore.release(element);
      throw error;
    }
  }
}
