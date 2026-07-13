import type { PresentationElement, PresentationElementInput, TextContentInput } from "../types/element.js";
import type { JsonValue, PresentationSlide, PresentationSlideInput } from "../types/presentation.js";
import type { PaintInput } from "../types/style.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import type { ElementStore } from "./element-store.js";

function assertInsertIndex(index: number, length: number, label: string): void {
  if (!Number.isInteger(index) || index < 0 || index > length) {
    throw new RangeError(`${label} index ${index} is outside 0..${length}.`);
  }
}

export class PresentationSlideImpl implements PresentationSlide {
  private readonly entries: PresentationElement[] = [];
  private readonly backgroundValue: PaintInput | undefined;
  private readonly notesValue: TextContentInput | undefined;
  private readonly sectionValue: string | undefined;
  private readonly layoutIdValue: string | undefined;
  readonly hidden: boolean;
  readonly tags: readonly string[];
  readonly customData: Readonly<Record<string, JsonValue>>;

  constructor(
    readonly id: string,
    input: PresentationSlideInput,
    private readonly elementStore: ElementStore,
  ) {
    this.layoutIdValue = input.layoutId;
    this.backgroundValue = input.background === undefined ? undefined : deepFreeze(deepClone(input.background));
    this.notesValue = input.notes === undefined ? undefined : deepFreeze(deepClone(input.notes));
    this.hidden = input.hidden ?? false;
    this.sectionValue = input.section;
    this.tags = deepFreeze([...(input.tags ?? [])]);
    this.customData = deepFreeze(deepClone(input.customData ?? {}));
  }

  get layoutId(): string | undefined {
    return this.layoutIdValue;
  }

  get background(): PaintInput | undefined {
    return this.backgroundValue;
  }

  get elements(): readonly PresentationElement[] {
    return Object.freeze([...this.entries]);
  }

  get notes(): TextContentInput | undefined {
    return this.notesValue;
  }

  get section(): string | undefined {
    return this.sectionValue;
  }

  addElement(input: PresentationElementInput): PresentationElement {
    return this.insertElement(this.entries.length, input);
  }

  insertElement(index: number, input: PresentationElementInput): PresentationElement {
    assertInsertIndex(index, this.entries.length, "Element");
    const element = this.elementStore.materialize(input);
    this.entries.splice(index, 0, element);
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

  removeElement(elementId: string): PresentationElement {
    const index = this.entries.findIndex((element) => element.id === elementId);
    if (index < 0) throw new Error(`Unknown element id "${elementId}".`);
    const [element] = this.entries.splice(index, 1);
    this.elementStore.release(element!);
    return element!;
  }

  duplicateElement(elementId: string, toIndex?: number): PresentationElement {
    const index = this.entries.findIndex((element) => element.id === elementId);
    if (index < 0) throw new Error(`Unknown element id "${elementId}".`);
    return this.insertElement(toIndex ?? index + 1, this.elementStore.withoutIds(this.entries[index]!));
  }

  releaseElements(): void {
    for (const element of this.entries) this.elementStore.release(element);
  }
}
