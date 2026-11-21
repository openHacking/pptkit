import type { LayoutElement, LayoutGroupElement, LayoutResult } from '@pptkit/layout';

type MediaElement = Extract<LayoutElement, { type: 'audio' | 'video' }>;
interface ParentLink { readonly group: LayoutGroupElement; readonly parent: ParentLink | undefined }

/** Bounded, identity-safe slide-local lookup shared by media mount and overlay. */
export function findLayoutMedia(layout: LayoutResult, slideId: string, elementId: string): { target: MediaElement; parents: readonly LayoutGroupElement[] } {
  if (!layout || layout.status !== 'resolved' || !Array.isArray(layout.slides) || !Array.isArray(layout.assets)) throw new TypeError('A resolved Layout result is required.');
  if (layout.slides.length > 100000 || layout.assets.length > 100000) throw new RangeError('Layout media search exceeds its work limit.');
  const slides = layout.slides.filter(slide => slide.id === slideId);
  if (slides.length !== 1) throw new TypeError('slideId must identify exactly one slide.');
  const slide = slides[0]!;
  if (!Array.isArray(slide.elements) || slide.elements.length > 100000) throw new RangeError('Slide media search exceeds its work limit.');
  const stack: { element: LayoutElement; hidden: boolean; parent: ParentLink | undefined }[] = slide.elements.map(element => ({ element, hidden: false, parent: undefined }));
  const seen = new Set<LayoutElement>(); let target: LayoutElement | undefined, hidden = false, parent: ParentLink | undefined, work = 0;
  while (stack.length) {
    if (++work > 100000) throw new RangeError('Slide media search exceeds its work limit.');
    const item = stack.pop()!, element = item.element;
    if (!element || typeof element !== 'object' || seen.has(element)) throw new TypeError('Slide element structure is invalid or cyclic.');
    seen.add(element);
    const invisible = item.hidden || element.hidden;
    if (element.id === elementId) { if (target) throw new TypeError('elementId is ambiguous on this slide.'); target = element; hidden = invisible; parent = item.parent; }
    if (element.type === 'group') {
      if (!Array.isArray(element.children)) throw new TypeError('Group children are invalid.');
      if (work + stack.length + element.children.length > 100000) throw new RangeError('Slide media search exceeds its work limit.');
      const link = { group: element, parent: item.parent };
      for (const child of element.children) stack.push({ element: child, hidden: invisible, parent: link });
    }
  }
  if (!target || (target.type !== 'audio' && target.type !== 'video')) throw new TypeError('elementId must identify an audio or video object on the slide.');
  if (hidden) throw new TypeError('Hidden media or media in a hidden group cannot be mounted.');
  const parents: LayoutGroupElement[] = []; while (parent) { parents.push(parent.group); parent = parent.parent; } parents.reverse();
  return { target, parents };
}
