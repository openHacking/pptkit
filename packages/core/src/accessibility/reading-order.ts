import type { SlideReadingOrder } from "../types/presentation.js";
import type { PresentationElement } from "../types/element.js";
import { deepFreeze } from "../utils/clone.js";

function visible(element: PresentationElement): boolean {
  return !element.hidden && element.accessibility?.decorative !== true;
}

function indexGroups(elements: readonly PresentationElement[]): Map<string, readonly PresentationElement[]> {
  const groups = new Map<string, readonly PresentationElement[]>(), stack = [...elements];
  while (stack.length) {
    const element = stack.pop()!;
    if (!visible(element) || element.type !== "group") continue;
    groups.set(element.id, element.children);
    for (const child of element.children) stack.push(child);
  }
  return groups;
}

export function captureSlideReadingOrder(elements: readonly PresentationElement[], input: readonly SlideReadingOrder[]): readonly SlideReadingOrder[] {
  if (!Array.isArray(input)) throw new TypeError("Slide readingOrder must be an array.");
  const groups = indexGroups(elements), owners = new Set<string | undefined>();
  return deepFreeze(Array.from(input, (scope: SlideReadingOrder) => {
    if (!scope || (scope.groupId !== undefined && typeof scope.groupId !== "string") || !Array.isArray(scope.elementIds) || owners.has(scope.groupId)) {
      throw new TypeError("Slide readingOrder requires unique owners and element ID arrays.");
    }
    owners.add(scope.groupId);
    const children = scope.groupId === undefined ? elements : groups.get(scope.groupId);
    if (!children) throw new RangeError("Slide readingOrder group must be visible, non-decorative and owned by the slide.");
    const ids = children.filter(visible).map(element => element.id), order = Array.from(scope.elementIds as readonly string[]);
    const expected = new Set(ids);
    if (order.length === 0 || expected.size !== ids.length || order.length !== ids.length || new Set(order).size !== order.length || order.some(id => !expected.has(id))) {
      throw new RangeError("Slide readingOrder must list each visible, non-decorative direct child exactly once.");
    }
    return { ...(scope.groupId === undefined ? {} : { groupId: scope.groupId }), elementIds: order };
  }));
}

/** Retain prior relative order after edits; append newly eligible children in drawing order. */
export function reconcileSlideReadingOrder(elements: readonly PresentationElement[], input: readonly SlideReadingOrder[]): readonly SlideReadingOrder[] | undefined {
  const groups = indexGroups(elements);
  const output: SlideReadingOrder[] = [];
  for (const scope of input) {
    const children = scope.groupId === undefined ? elements : groups.get(scope.groupId);
    if (!children) continue;
    const ids = new Set(children.filter(visible).map(element => element.id));
    const retained = scope.elementIds.filter(id => ids.delete(id));
    const elementIds = [...retained, ...ids];
    if (elementIds.length) output.push({ ...(scope.groupId === undefined ? {} : { groupId: scope.groupId }), elementIds });
  }
  return output.length ? deepFreeze(output) : undefined;
}

export function remapSlideReadingOrder(input: readonly SlideReadingOrder[], identities: ReadonlyMap<string, string>): readonly SlideReadingOrder[] {
  const remap = (id: string): string => {
    const mapped = identities.get(id);
    if (mapped === undefined) throw new RangeError("Cannot copy a slide readingOrder with an unmapped object.");
    return mapped;
  };
  return input.map(scope => ({ ...(scope.groupId === undefined ? {} : { groupId: remap(scope.groupId) }), elementIds: scope.elementIds.map(remap) }));
}
