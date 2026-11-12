import { captureSlideReadingOrder, reconcileSlideReadingOrder } from "../accessibility/reading-order.js";
import type { EventTimelinePlan, SlideEventTimelineInput } from "../types/event-timeline.js";
import { captureSlideEventTimeline, compileSlideEventTimeline } from "./slide-event-timeline.js";
import type { GroupElement, PresentationElement, PresentationElementInput, TextContentInput } from "../types/element.js";
import type { JsonValue, ManagedPresentationSlide, PresentationSlideInput, SlideReadingOrder, SlideTransitionInput } from "../types/presentation.js";
import type { BackgroundPaintInput } from "../types/style.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import type { ElementStore } from "./element-store.js";

function assertInsertIndex(index: number, length: number, label: string): void {
  if (!Number.isInteger(index) || index < 0 || index > length) {
    throw new RangeError(`${label} index ${index} is outside 0..${length}.`);
  }
}

function resolveElementOrder(original: readonly PresentationElement[], elementIds: readonly string[], scope: string): PresentationElement[] {
  if (!Array.isArray(elementIds)) throw new TypeError("Element order must be an array of IDs.");
  if (elementIds.length !== original.length) throw new Error(`Element order must contain every ${scope} ID exactly once.`);
  const byId = new Map(original.map(element => [element.id, element]));
  const ordered: PresentationElement[] = [], seen = new Set<string>();
  for (let index = 0; index < original.length; index++) {
    const id: unknown = elementIds[index];
    if (typeof id !== "string" || seen.has(id) || !byId.has(id)) throw new Error(`Element order must contain every ${scope} ID exactly once.`);
    seen.add(id); ordered.push(byId.get(id)!);
  }
  return ordered;
}

export class PresentationSlideImpl implements ManagedPresentationSlide {
  private inheritedShapesValue: boolean | undefined;
  private readonly entries: PresentationElement[] = [];
  private backgroundValue: BackgroundPaintInput | undefined;
  private notesValue: TextContentInput | undefined;
  private readingOrderValue: readonly SlideReadingOrder[] | undefined;
  private eventTimelineValue: SlideEventTimelineInput | undefined;
  private transitionValue: SlideTransitionInput | undefined;
  private readonly sectionValue: string | undefined;
  private layoutIdValue: string | undefined;
  private hiddenValue: boolean;
  readonly tags: readonly string[];
  readonly customData: Readonly<Record<string, JsonValue>>;

  constructor(
    readonly id: string,
    input: PresentationSlideInput,
    private elementStore: ElementStore,
  ) {
    this.inheritedShapesValue = input.showInheritedShapes;
    this.layoutIdValue = input.layoutId;
    this.backgroundValue = input.background === undefined ? undefined : deepFreeze(deepClone(input.background));
    this.notesValue = input.notes === undefined ? undefined : deepFreeze(deepClone(input.notes));
    this.transitionValue = input.transition === undefined ? undefined : deepFreeze(deepClone(input.transition));
    this.hiddenValue = input.hidden ?? false;
    Object.defineProperty(this, "hidden", { enumerable: true, configurable: false, get: () => this.hiddenValue });
    this.sectionValue = input.section;
    this.tags = deepFreeze([...(input.tags ?? [])]);
    this.customData = deepFreeze(deepClone(input.customData ?? {}));
  }

  /** Internal document append transaction: switch staged ownership to the live store. */
  bindElementStore(store: ElementStore): void { this.elementStore = store; }

  /** Internal document transaction commit; identities were staged in the shared store. */
  commitLayoutRebind(layoutId: string | undefined, background: BackgroundPaintInput | undefined, elements: readonly PresentationElement[]): void {
    this.layoutIdValue = layoutId;
    this.backgroundValue = background;
    this.entries.length = 0;
    for (const element of elements) this.entries.push(element);
    this.reconcileReadingOrder();
  }

  get hidden(): boolean { return this.hiddenValue; }

  setHidden(value: boolean): void {
    if (typeof value !== "boolean") throw new TypeError("hidden must be boolean.");
    this.hiddenValue = value;
  }

  get showInheritedShapes(): boolean | undefined { return this.inheritedShapesValue; }

  setShowInheritedShapes(value?: boolean): void {
    if (value !== undefined && typeof value !== "boolean") throw new TypeError("showInheritedShapes must be boolean or undefined.");
    this.inheritedShapesValue = value;
  }

  get layoutId(): string | undefined {
    return this.layoutIdValue;
  }

  get background(): BackgroundPaintInput | undefined {
    return this.backgroundValue;
  }

  setBackground(value?: BackgroundPaintInput): void {
    this.backgroundValue = value === undefined ? undefined : deepFreeze(deepClone(value));
  }

  get elements(): readonly PresentationElement[] {
    return Object.freeze([...this.entries]);
  }

  get notes(): TextContentInput | undefined {
    return this.notesValue;
  }

  get transition(): SlideTransitionInput | undefined {
    return this.transitionValue;
  }

  get readingOrder(): readonly SlideReadingOrder[] | undefined { return this.readingOrderValue; }

  setReadingOrder(order?: readonly SlideReadingOrder[]): void {
    if (order === undefined) { this.readingOrderValue = undefined; return; }
    const original = this.entries.slice(), previous = this.readingOrderValue;
    const captured = captureSlideReadingOrder(original, order);
    if (this.readingOrderValue !== previous || this.entries.length !== original.length || this.entries.some((element, index) => element !== original[index])) {
      throw new Error("Slide changed while capturing its reading order.");
    }
    this.readingOrderValue = captured.length ? captured : undefined;
  }

  private reconcileReadingOrder(): void {
    if (this.readingOrderValue) this.readingOrderValue = reconcileSlideReadingOrder(this.entries, this.readingOrderValue);
  }

  get eventTimeline(): SlideEventTimelineInput | undefined { return this.eventTimelineValue; }

  setEventTimeline(timeline?: SlideEventTimelineInput): void {
    const captured = timeline === undefined ? undefined : captureSlideEventTimeline(this.id, this.entries, timeline);
    this.eventTimelineValue = captured;
  }

  getEventTimelinePlan(): EventTimelinePlan | undefined {
    return this.eventTimelineValue === undefined ? undefined : compileSlideEventTimeline(this.id, this.entries, this.eventTimelineValue);
  }

  private validateTimeline(elements: readonly PresentationElement[]): void {
    if (this.eventTimelineValue !== undefined) compileSlideEventTimeline(this.id, elements, this.eventTimelineValue);
  }

  /** Internal transaction support: validate all replacements together before committing. */
  validateEditedElements(replacements: ReadonlyMap<string, PresentationElement>): void {
    if (this.eventTimelineValue !== undefined) this.validateTimeline(this.entries.map(element => replacements.get(element.id) ?? element));
  }

  commitEditedElements(replacements: ReadonlyMap<string, PresentationElement>): void {
    for (let i = 0; i < this.entries.length; i++) {
      const replacement = replacements.get(this.entries[i]!.id);
      if (replacement) this.entries[i] = replacement;
    }
    this.reconcileReadingOrder();
  }

  replaceElement(elementId: string, input: PresentationElementInput): PresentationElement {
    const index = this.entries.findIndex(element => element.id === elementId);
    if (index < 0) throw new Error(`Unknown element id "${elementId}".`);
    const element = this.elementStore.replace(this.entries[index]!, input, candidate => {
      if (this.eventTimelineValue !== undefined) this.validateTimeline(this.entries.map((current, i) => i === index ? candidate : current));
    });
    this.entries[index] = element;
    this.reconcileReadingOrder();
    return element;
  }

  get section(): string | undefined {
    return this.sectionValue;
  }

  setNotes(notes?: TextContentInput): void {
    this.notesValue = notes === undefined ? undefined : deepFreeze(deepClone(notes));
  }

  setTransition(transition?: SlideTransitionInput): void {
    this.transitionValue = transition === undefined ? undefined : deepFreeze(deepClone(transition));
  }

  addElement(input: PresentationElementInput): PresentationElement {
    return this.insertElement(this.entries.length, input);
  }

  insertElement(index: number, input: PresentationElementInput): PresentationElement {
    assertInsertIndex(index, this.entries.length, "Element");
    const validate = this.eventTimelineValue === undefined ? undefined : (candidate: PresentationElement) => this.validateTimeline([...this.entries.slice(0, index), candidate, ...this.entries.slice(index)]);
    const element = this.elementStore.materialize(input, validate);
    this.entries.splice(index, 0, element);
    this.reconcileReadingOrder();
    return element;
  }

  moveElement(elementId: string, toIndex: number): void {
    const fromIndex = this.entries.findIndex((element) => element.id === elementId);
    if (fromIndex < 0) throw new Error(`Unknown element id "${elementId}".`);
    if (!Number.isInteger(toIndex) || toIndex < 0 || toIndex >= this.entries.length) {
      throw new RangeError(`Element index ${toIndex} is outside 0..${Math.max(0, this.entries.length - 1)}.`);
    }
    const [element] = this.entries.splice(fromIndex, 1);
    this.entries.splice(toIndex, 0, element!);
  }

  reorderElements(elementIds: readonly string[]): void {
    const original = this.entries.slice();
    const ordered = resolveElementOrder(original, elementIds, "top-level");
    if (this.entries.length !== original.length || this.entries.some((element, index) => element !== original[index])) throw new Error("Slide elements changed during reorder.");
    for (let index = 0; index < ordered.length; index++) this.entries[index] = ordered[index]!;
  }

  reorderGroupElements(groupId: string, elementIds: readonly string[]): void {
    const original = this.entries.slice();
    type Ancestor = { group: GroupElement; index: number; parent?: Ancestor };
    type Search = { element: PresentationElement; rootIndex: number; parent?: Ancestor };
    const stack: Search[] = original.map((element, rootIndex) => ({ element, rootIndex })).reverse();
    let found: Search | undefined;
    while (stack.length) {
      const item = stack.pop()!;
      if (item.element.id === groupId) { found = item; break; }
      if (item.element.type === "group") for (let index = item.element.children.length - 1; index >= 0; index--) stack.push({ element: item.element.children[index]!, rootIndex: item.rootIndex, parent: { group: item.element, index, ...(item.parent ? { parent: item.parent } : {}) } });
    }
    if (!found || found.element.type !== "group") throw new Error(`Unknown group id "${groupId}".`);
    const ordered = resolveElementOrder(found.element.children, elementIds, "direct child");
    if (this.entries.length !== original.length || this.entries.some((element, index) => element !== original[index])) throw new Error("Slide elements changed during reorder.");
    let replacement: PresentationElementInput = { ...found.element, children: ordered as PresentationElementInput[] };
    for (let parent = found.parent; parent; parent = parent.parent) {
      const { group, index: childIndex } = parent;
      const children = [...group.children] as PresentationElementInput[]; children[childIndex] = replacement;
      replacement = { ...group, children };
    }
    this.replaceElement(original[found.rootIndex]!.id, replacement);
  }

  removeElement(elementId: string): PresentationElement {
    const index = this.entries.findIndex((element) => element.id === elementId);
    if (index < 0) throw new Error(`Unknown element id "${elementId}".`);
    if (this.eventTimelineValue !== undefined) this.validateTimeline(this.entries.filter((_, i) => i !== index));
    const [element] = this.entries.splice(index, 1);
    this.elementStore.release(element!);
    this.reconcileReadingOrder();
    return element!;
  }

  duplicateElement(elementId: string, toIndex?: number): PresentationElement {
    const index = this.entries.findIndex((element) => element.id === elementId);
    if (index < 0) throw new Error(`Unknown element id "${elementId}".`);
    const targetIndex = toIndex ?? index + 1;
    assertInsertIndex(targetIndex, this.entries.length, "Element");
    return this.insertElement(targetIndex, this.elementStore.copyInputs([this.entries[index]!])[0]!);
  }

  releaseElements(): void {
    for (const element of this.entries) this.elementStore.release(element);
  }
}
