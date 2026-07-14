import { DEFAULT_PRESENTATION_SIZE } from "../constants/presentation.js";
import type { PresentationAsset, PresentationAssetInput } from "../types/asset.js";
import type { PresentationElementInput } from "../types/element.js";
import type { PresentationSize } from "../types/geometry.js";
import type {
  PresentationDocument,
  PresentationInit,
  PresentationMetadataInput,
  PresentationSlide,
  PresentationSlideInput,
  SlideLayoutDefinition,
  SlideLayoutInput,
} from "../types/presentation.js";
import type { PresentationThemeInput } from "../types/theme.js";
import type { TextStylePresetMap } from "../types/style.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import { createDocumentId, IdAllocator } from "../utils/id.js";
import { normalizeSize } from "../validation/geometry.js";
import { AssetRegistry } from "./asset-registry.js";
import { ElementStore } from "./element-store.js";
import { PresentationSlideImpl } from "./presentation-slide.js";

function assertInsertIndex(index: number, length: number, label: string): void {
  if (!Number.isInteger(index) || index < 0 || index > length) {
    throw new RangeError(`${label} index ${index} is outside 0..${length}.`);
  }
}

export class PresentationDocumentImpl implements PresentationDocument {
  readonly id: string;
  private readonly metadataValue: Readonly<PresentationMetadataInput>;
  private readonly sizeValue: Readonly<PresentationSize>;
  private readonly themeValue: Readonly<PresentationThemeInput>;
  private readonly textStylePresetsValue: TextStylePresetMap;
  private readonly slideEntries: PresentationSlideImpl[] = [];
  private readonly layoutEntries: SlideLayoutDefinition[] = [];
  private readonly assetRegistry = new AssetRegistry();
  private readonly elementStore = new ElementStore();
  private readonly slideIds = new Set<string>();
  private readonly slideAllocator = new IdAllocator("slide");

  constructor(init: PresentationInit = {}) {
    this.id = init.id ?? createDocumentId();
    this.metadataValue = deepFreeze(deepClone(init.metadata ?? {}));
    this.sizeValue = deepFreeze(normalizeSize(init.size, DEFAULT_PRESENTATION_SIZE));
    this.themeValue = deepFreeze(deepClone(init.theme ?? {}));
    this.textStylePresetsValue = deepFreeze(deepClone(init.textStylePresets ?? {}));
  }

  get metadata(): Readonly<PresentationMetadataInput> {
    return this.metadataValue;
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

  get slides(): readonly PresentationSlide[] {
    return Object.freeze([...this.slideEntries]);
  }

  get assets(): readonly PresentationAsset[] {
    return this.assetRegistry.assets;
  }

  get layouts(): readonly SlideLayoutDefinition[] {
    return Object.freeze([...this.layoutEntries]);
  }

  addSlide(input: PresentationSlideInput = {}): PresentationSlide {
    return this.insertSlide(this.slideEntries.length, input);
  }

  insertSlide(index: number, input: PresentationSlideInput = {}): PresentationSlide {
    assertInsertIndex(index, this.slideEntries.length, "Slide");
    const id = input.id ?? this.slideAllocator.next((candidate) => this.slideIds.has(candidate));
    if (this.slideIds.has(id)) throw new Error(`Duplicate slide id "${id}" detected.`);
    const slide = new PresentationSlideImpl(id, input, this.elementStore);
    try {
      for (const element of input.elements ?? []) slide.addElement(element);
    } catch (error) {
      slide.releaseElements();
      throw error;
    }
    this.slideIds.add(id);
    this.slideEntries.splice(index, 0, slide);
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
  }

  removeSlide(slideId: string): PresentationSlide {
    const index = this.slideEntries.findIndex((slide) => slide.id === slideId);
    if (index < 0) throw new Error(`Unknown slide id "${slideId}".`);
    const [slide] = this.slideEntries.splice(index, 1);
    this.slideIds.delete(slideId);
    slide!.releaseElements();
    return slide!;
  }

  duplicateSlide(slideId: string, toIndex?: number): PresentationSlide {
    const index = this.slideEntries.findIndex((slide) => slide.id === slideId);
    if (index < 0) throw new Error(`Unknown slide id "${slideId}".`);
    const source = this.slideEntries[index]!;
    const elements: PresentationElementInput[] = source.elements.map((element) => this.elementStore.withoutIds(element));
    return this.insertSlide(toIndex ?? index + 1, {
      ...(source.layoutId !== undefined ? { layoutId: source.layoutId } : {}),
      ...(source.background !== undefined ? { background: deepClone(source.background) } : {}),
      elements,
      ...(source.notes !== undefined ? { notes: deepClone(source.notes) } : {}),
      hidden: source.hidden,
      ...(source.section !== undefined ? { section: source.section } : {}),
      tags: [...source.tags],
      customData: deepClone(source.customData),
    });
  }

  registerAsset(input: PresentationAssetInput): PresentationAsset {
    return this.assetRegistry.register(input);
  }

  getAsset(assetId: string): PresentationAsset | undefined {
    return this.assetRegistry.get(assetId);
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
        ...(input.background !== undefined ? { background: deepClone(input.background) } : {}),
        elements,
        placeholders,
      }) as SlideLayoutDefinition;
      this.layoutEntries.push(layout);
      return layout;
    } catch (error) {
      for (const element of elements) this.elementStore.release(element);
      throw error;
    }
  }
}
