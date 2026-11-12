import type { AssetReference } from "../types/asset.js";
import type { PresentationElement } from "../types/element.js";
import type { PresentationSlide, SlideLayoutDefinition, SlideMasterDefinition } from "../types/presentation.js";

/** Known authoring references only; no custom-data guessing or format preservation lookup. */
export function collectAssetReferences(slides: readonly PresentationSlide[], layouts: readonly SlideLayoutDefinition[], masters: readonly SlideMasterDefinition[] = []): Map<string, AssetReference[]> {
  const result = new Map<string, AssetReference[]>();
  const collect = (elements: readonly PresentationElement[], path: string, owner: { slideId: string } | { layoutId: string } | { masterId: string }) => {
    const stack: { elements: readonly PresentationElement[]; path: string; index: number }[] = [{ elements, path, index: 0 }];
    while (stack.length) {
      const frame = stack[stack.length - 1]!;
      if (frame.index === frame.elements.length) { stack.pop(); continue; }
      const index = frame.index++, element = frame.elements[index]!, elementPath = `${frame.path}.${index}`;
      if (element.type === "group") stack.push({ elements: element.children, path: `${elementPath}.children`, index: 0 });
      if (element.type === "image") {
        let references = result.get(element.assetId);
        if (!references) { references = []; result.set(element.assetId, references); }
        references.push(Object.freeze({ assetId: element.assetId, elementId: element.id, path: `${elementPath}.assetId`, ...owner }));
      }
    }
  };
  slides.forEach((slide, index) => collect(slide.elements, `slides.${index}.elements`, { slideId: slide.id }));
  layouts.forEach((layout, index) => collect(layout.elements, `layouts.${index}.elements`, { layoutId: layout.id }));
  masters.forEach((master, index) => collect(master.elements ?? [], `masters.${index}.elements`, { masterId: master.id }));
  return result;
}
