import type { LayoutElement, LayoutResult } from "@pptkit/layout";
import type { SvgSlideReadingOrder } from "../types/public.js";

function hasUrlAction(elements: readonly LayoutElement[]): boolean {
  const pending = [...elements];
  while (pending.length) {
    const element = pending.pop()!;
    if (element.hidden) continue;
    // The outer object action is exposed even on a decorative object.
    if (element.action?.type === "url") return true;
    if (element.accessibility.decorative) continue;
    if (element.type === "group") for (const child of element.children) pending.push(child);
    const paragraphs = element.type === "text" ? element.content : element.type === "shape" ? element.text?.content : undefined;
    if (paragraphs?.some(p => p.runs.some(run => run.action?.type === "url"))) return true;
    if (element.type === "table" && element.rows.some(row => row.cells.some(cell => cell.content.some(p => p.runs.some(run => run.action?.type === "url"))))) return true;
  }
  return false;
}

type OrderIds = ReadonlyMap<string, string>;

export function readingOrderIds(elementIds: readonly string[], prefix: string): OrderIds {
  return new Map(elementIds.map((id, position) => [id, `${prefix}-${position}`]));
}

export function wrapReadingOrder(element: LayoutElement, rendered: string, ids: OrderIds | undefined): string {
  const id = element.hidden || element.accessibility.decorative ? undefined : ids?.get(element.id);
  return id === undefined ? rendered : `<g id="${id}">${rendered}</g>`;
}

export function readingOrderOwnership(ids: OrderIds | undefined): string {
  return ids === undefined ? "" : ` aria-owns="${[...ids.values()].join(" ")}"`;
}

/** Capture independent order before rendering or invoking user callbacks. */
export function captureReadingOrders(layout: LayoutResult, input: readonly SvgSlideReadingOrder[] | undefined): ReadonlyMap<string, readonly SvgSlideReadingOrder[]> {
  const orders = new Map<string, SvgSlideReadingOrder[]>();
  if (input === undefined) input = layout.slides.flatMap(slide => slide.hidden ? [] : (slide.readingOrder ?? []).map(scope => ({ ...scope, slideId: slide.id })));
  if (!Array.isArray(input)) throw new TypeError("readingOrder must be an array of slide orders.");
  if (input.length === 0) return orders;
  const slides = new Map(layout.slides.map(slide => [slide.id, slide]));
  const groupIndexes = new Map<string, Map<string, Extract<LayoutElement, { type: "group" }> | null>>();
  const owners = new Map<string, Set<string | undefined>>();
  for (const entry of input as readonly SvgSlideReadingOrder[]) {
    if (!entry || typeof entry.slideId !== "string" || !Array.isArray(entry.elementIds) || (entry.groupId !== undefined && typeof entry.groupId !== "string")) throw new TypeError("readingOrder requires slide IDs, optional group IDs and element ID arrays.");
    const slide = slides.get(entry.slideId);
    if (!slide || slide.hidden) throw new RangeError("readingOrder requires an existing visible slide.");
    const scopes = owners.get(slide.id) ?? new Set();
    if (scopes.has(entry.groupId)) throw new TypeError("readingOrder requires distinct slide IDs for root orders and distinct group IDs within a slide.");
    scopes.add(entry.groupId);
    owners.set(slide.id, scopes);
    let children = slide.elements;
    if (entry.groupId !== undefined) {
      let groups = groupIndexes.get(slide.id);
      if (!groups) {
        groups = new Map();
        const pending = [...slide.elements];
        while (pending.length) {
          const element = pending.pop()!;
          if (element.type !== "group" || element.hidden || element.accessibility.decorative) continue;
          groups.set(element.id, groups.has(element.id) ? null : element);
          for (const child of element.children) pending.push(child);
        }
        groupIndexes.set(slide.id, groups);
      }
      const group = groups.get(entry.groupId);
      if (!group) throw new RangeError("readingOrder requires an unambiguous visible, non-decorative slide-owned group.");
      children = group.children;
    }
    const elements = children.filter(element => !element.hidden && !element.accessibility.decorative);
    const ids = new Set(elements.map(element => element.id));
    const order = Array.from(entry.elementIds as readonly string[]);
    if (ids.size !== elements.length || !order.length || order.length !== elements.length || new Set(order).size !== order.length || order.some(id => !ids.has(id))) {
      const scope = entry.groupId === undefined ? "slide-owned root" : "direct group child";
      throw new RangeError(`readingOrder must list every visible non-decorative ${scope} exactly once.`);
    }
    if (hasUrlAction(elements)) throw new RangeError("readingOrder does not reorder artwork with URL actions; its keyboard focus order would remain unchanged.");
    const list = orders.get(slide.id) ?? [];
    list.push({ slideId: slide.id, ...(entry.groupId === undefined ? {} : { groupId: entry.groupId }), elementIds: order });
    orders.set(slide.id, list);
  }
  // Equivalent scope lists yield identical wrapper IDs and incremental dependencies.
  for (const list of orders.values()) list.sort((a, b) => {
    if (a.groupId === b.groupId) return 0;
    if (a.groupId === undefined) return -1;
    if (b.groupId === undefined) return 1;
    return a.groupId < b.groupId ? -1 : 1;
  });
  return orders;
}
