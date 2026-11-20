import { captureSlideReadingOrder } from "../accessibility/reading-order.js";
import type { ElementAction, PresentationElement, TextContentInput } from '../types/element.js';
import type { PresentationSlideData } from '../types/presentation.js';
import { compileSlideEventTimeline } from './slide-event-timeline.js';
import { deepFreeze } from '../utils/clone.js';

/** Internal source inventory for cross-document slide transfer; no mutation or write authority. */
export interface SlideCopyDependencies {
  readonly slideId: string;
  readonly layoutId?: string;
  readonly elementIds: readonly string[];
  readonly assets: readonly { readonly id: string; readonly path: string }[];
  readonly connectors: readonly { readonly targetId: string; readonly path: string; readonly slideOwned: boolean }[];
  readonly slideActions: readonly { readonly targetId: string; readonly path: string; readonly self: boolean }[];
  readonly urlActions: readonly { readonly path: string }[];
  readonly placeholders: readonly { readonly key: string; readonly path: string }[];
  readonly textPresets: readonly { readonly name: string; readonly path: string }[];
  readonly fields: readonly { readonly id: string; readonly kind: string; readonly path: string }[];
  readonly timelineTargets: readonly { readonly stageId: string; readonly targetId: string; readonly triggerId?: string }[];
  readonly hasTimeline: boolean;
  readonly hasOpaqueCustomData: boolean;
}

export function collectSlideCopyDependencies(slide: PresentationSlideData, maxWork = 100000): SlideCopyDependencies {
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError('Slide copy work limit must be a positive safe integer.');
  if (!slide || typeof slide.id !== 'string' || !slide.id || !Array.isArray(slide.elements)) throw new TypeError('Slide copy requires a valid slide snapshot.');
  let work = 0;
  const charge = () => { if (++work > maxWork) throw new RangeError('Slide copy dependency work limit exceeded.'); };
  const ids: string[] = [], owned = new Set<string>(), elements: { element: PresentationElement; path: string }[] = [];
  if (slide.elements.length > maxWork) throw new RangeError('Slide copy dependency work limit exceeded.');
  const stack = slide.elements.map((element, index) => ({ element, path: `elements.${index}` })).reverse();
  while (stack.length) {
    charge(); const entry = stack.pop()!, element = entry.element;
    if (!element || typeof element.id !== 'string' || !element.id || owned.has(element.id)) throw new TypeError('Slide element identities must be unique nonempty strings.');
    if (!['text', 'shape', 'image', 'audio', 'video', 'connector', 'group', 'table', 'chart'].includes(element.type)) throw new TypeError('Slide copy encountered an unsupported element type.');
    ids.push(element.id); owned.add(element.id); elements.push(entry);
    if (element.type === 'group') {
      if (!Array.isArray(element.children)) throw new TypeError('Group children must be an array.');
      if (element.children.length > maxWork - work - stack.length) throw new RangeError('Slide copy dependency work limit exceeded.');
      for (let i = element.children.length - 1; i >= 0; i--) stack.push({ element: element.children[i]!, path: `${entry.path}.children.${i}` });
    }
  }
  const assets: { id: string; path: string }[] = [], connectors: { targetId: string; path: string; slideOwned: boolean }[] = [];
  const slideActions: { targetId: string; path: string; self: boolean }[] = [], urlActions: { path: string }[] = [];
  const placeholders: { key: string; path: string }[] = [], textPresets: { name: string; path: string }[] = [], fields: { id: string; kind: string; path: string }[] = [];
  const action = (value: ElementAction | undefined, path: string) => {
    if (value === undefined) return; charge();
    if (value.type === 'slide' && typeof value.slideId === 'string' && value.slideId) slideActions.push({ targetId: value.slideId, path, self: value.slideId === slide.id });
    else if (value.type === 'url' && typeof value.url === 'string') urlActions.push({ path });
    else throw new TypeError('Slide copy encountered an unsupported action.');
  };
  const text = (content: TextContentInput | undefined, path: string) => {
    if (content === undefined || typeof content === 'string') return;
    if (!Array.isArray(content)) throw new TypeError('Slide copy requires valid text content.');
    for (let pi = 0; pi < content.length; pi++) {
      const paragraph = content[pi];
      charge(); if (!paragraph || !Array.isArray(paragraph.runs)) throw new TypeError('Slide copy requires valid text runs.');
      for (let ri = 0; ri < paragraph.runs.length; ri++) {
        const run = paragraph.runs[ri];
        charge(); if (!run || typeof run !== 'object') throw new TypeError('Slide copy requires valid text runs.');
        const at = `${path}.${pi}.runs.${ri}`; action(run.action, `${at}.action`);
        if (run.field !== undefined) { if (!run.field || typeof run.field.id !== 'string' || !run.field.id || !['slide-number', 'dynamic-date'].includes(run.field.kind)) throw new TypeError('Slide copy encountered an unsupported field.'); fields.push({ id: run.field.id, kind: run.field.kind, path: `${at}.field` }); }
      }
    }
  };
  const preset = (name: string | undefined, path: string) => { if (name !== undefined) { charge(); if (typeof name !== 'string' || !name) throw new TypeError('Slide copy requires nonempty preset names.'); textPresets.push({ name, path }); } };
  for (const { element, path } of elements) {
    action(element.action, `${path}.action`);
    if (element.placeholderKey !== undefined) { charge(); if (typeof element.placeholderKey !== 'string' || !element.placeholderKey) throw new TypeError('Slide copy requires nonempty placeholder keys.'); placeholders.push({ key: element.placeholderKey, path: `${path}.placeholderKey` }); }
    if (element.type === 'image') { charge(); if (typeof element.assetId !== 'string' || !element.assetId) throw new TypeError('Slide copy requires nonempty image asset IDs.'); assets.push({ id: element.assetId, path: `${path}.assetId` }); }
    else if (element.type === 'audio') for (const key of ['audioAssetId', 'posterAssetId'] as const) { const id = element[key]; if (key === 'posterAssetId' && id === undefined) continue; charge(); if (typeof id !== 'string' || !id) throw new TypeError('Slide copy requires nonempty audio and poster asset IDs.'); assets.push({ id, path: `${path}.${key}` }); }
    else if (element.type === 'video') for (const key of ['videoAssetId', 'posterAssetId'] as const) { const id = element[key]; if (key === 'posterAssetId' && id === undefined) continue; charge(); if (typeof id !== 'string' || !id) throw new TypeError('Slide copy requires nonempty video and poster asset IDs.'); assets.push({ id, path: `${path}.${key}` }); }
    else if (element.type === 'connector') for (const key of ['start', 'end'] as const) {
      const endpoint = element[key]; if (!endpoint || typeof endpoint !== 'object') throw new TypeError('Slide copy requires valid connector endpoints.');
      if ('elementId' in endpoint) { charge(); if (typeof endpoint.elementId !== 'string' || !endpoint.elementId) throw new TypeError('Slide copy requires nonempty connector target IDs.'); connectors.push({ targetId: endpoint.elementId, path: `${path}.${key}`, slideOwned: owned.has(endpoint.elementId) }); }
    }
    else if (element.type === 'text') { preset(element.textStylePreset, `${path}.textStylePreset`); text(element.content, `${path}.content`); }
    else if (element.type === 'shape' && element.text !== undefined) { preset(element.text.textStylePreset, `${path}.text.textStylePreset`); text(element.text.content, `${path}.text.content`); }
    else if (element.type === 'table') {
      if (!Array.isArray(element.rows)) throw new TypeError('Slide copy requires valid table rows.');
      for (let ri = 0; ri < element.rows.length; ri++) {
        const row = element.rows[ri]; charge(); if (!row || !Array.isArray(row.cells)) throw new TypeError('Slide copy requires valid table cells.');
        for (let ci = 0; ci < row.cells.length; ci++) {
          const cell = row.cells[ci]; charge(); if (!cell || typeof cell !== 'object') throw new TypeError('Slide copy requires valid table cells.');
          const at = `${path}.rows.${ri}.cells.${ci}`; preset(cell.textStylePreset, `${at}.textStylePreset`); text(cell.content, `${at}.content`);
        }
      }
    }
  }
  if (slide.readingOrder !== undefined) {
    if (!Array.isArray(slide.readingOrder)) throw new TypeError("Slide readingOrder must be an array.");
    for (const scope of slide.readingOrder) {
      charge();
      if (!scope || !Array.isArray(scope.elementIds)) throw new TypeError("Slide readingOrder requires element ID arrays.");
      for (const _id of scope.elementIds) charge();
    }
    captureSlideReadingOrder(slide.elements, slide.readingOrder);
  }
  text(slide.notes, 'notes');
  const timelineTargets: { stageId: string; targetId: string; triggerId?: string }[] = [];
  if (slide.eventTimeline !== undefined) {
    charge();
    if (Array.isArray(slide.eventTimeline.stages) && slide.eventTimeline.stages.length > maxWork - work) throw new RangeError('Slide copy dependency work limit exceeded.');
    compileSlideEventTimeline(slide.id, slide.elements, slide.eventTimeline);
    for (const stage of slide.eventTimeline.stages) { charge(); timelineTargets.push({ stageId: stage.id, targetId: stage.target.elementId, ...(stage.trigger.type === 'object-click' ? { triggerId: stage.trigger.elementId } : {}) }); }
  }
  return deepFreeze({ slideId: slide.id, ...(slide.layoutId !== undefined ? { layoutId: slide.layoutId } : {}), elementIds: ids, assets, connectors, slideActions, urlActions, placeholders, textPresets, fields, timelineTargets, hasTimeline: slide.eventTimeline !== undefined, hasOpaqueCustomData: Object.keys(slide.customData).length > 0 });
}
