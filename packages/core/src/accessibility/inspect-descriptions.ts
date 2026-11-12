import { PresentationValidationError } from '../types/diagnostic.js';
import { DEFAULT_LAYOUT_ID } from '../constants/presentation.js';
import type { PresentationElement } from '../types/element.js';
import type { PresentationData } from '../types/presentation.js';
import { validatePresentation } from '../validation/validate-presentation.js';

export interface MissingDescriptionFinding {
  readonly code: 'missing-description';
  readonly elementType: 'image' | 'chart' | 'table';
  readonly elementId: string;
  readonly path: string;
  readonly slideId?: string;
  readonly layoutId?: string;
  readonly masterId?: string;
}
export interface DescriptionInspection {
  readonly scope: 'authored-definitions';
  readonly findings: readonly MissingDescriptionFinding[];
}

export interface EffectiveDescriptionFinding {
  readonly code: 'missing-description';
  readonly elementType: 'image' | 'chart' | 'table';
  readonly elementId: string;
  readonly path: string;
  readonly slideId: string;
  readonly source: Readonly<{ kind: 'slide' | 'layout' | 'master'; id: string }>;
}
export interface EffectiveDescriptionInspection {
  readonly scope: 'modeled-slide-occurrences';
  readonly findings: readonly EffectiveDescriptionFinding[];
}
export class EffectiveDescriptionLimitError extends Error {
  constructor(readonly maxFindings: number) { super('Effective description finding limit exceeded.'); this.name = 'EffectiveDescriptionLimitError'; }
}

/** Review explicit object descriptions; this does not compute reading order or contrast. */
export function inspectMissingDescriptions(document: PresentationData): DescriptionInspection {
  const diagnostics = validatePresentation(document);
  if (diagnostics.some(item => item.severity === 'error')) throw new PresentationValidationError(diagnostics);
  const findings: MissingDescriptionFinding[] = [];
  const walk = (elements: readonly PresentationElement[], path: string, owner: Pick<MissingDescriptionFinding, 'slideId' | 'layoutId' | 'masterId'>, suppressed = false): void => {
    elements.forEach((element, index) => {
      const current = `${path}.${index}`, skip = suppressed || element.hidden === true || element.accessibility?.decorative === true;
      if (element.type === 'group') walk(element.children, `${current}.children`, owner, skip);
      else if (!skip && (element.type === 'image' || element.type === 'chart' || element.type === 'table') && !element.accessibility?.description?.trim()) {
        findings.push(Object.freeze({ code: 'missing-description', elementType: element.type, elementId: element.id, path: `${current}.accessibility.description`, ...owner }));
      }
    });
  };
  document.slides.forEach((slide, index) => { if (!slide.hidden) walk(slide.elements, `slides.${index}.elements`, { slideId: slide.id }); });
  document.layouts.forEach((layout, index) => walk(layout.elements, `layouts.${index}.elements`, { layoutId: layout.id }));
  (document.masters ?? []).forEach((master, index) => walk(master.elements ?? [], `masters.${index}.elements`, { masterId: master.id }));
  return Object.freeze({ scope: 'authored-definitions', findings: Object.freeze(findings) });
}

/** Expand authored findings to modeled visible slide occurrences, not native reading order. */
export function inspectEffectiveMissingDescriptions(document: PresentationData, options: { maxFindings?: number } = {}): EffectiveDescriptionInspection {
  if (!options || typeof options !== 'object' || Array.isArray(options) || options.maxFindings !== undefined && (!Number.isSafeInteger(options.maxFindings) || options.maxFindings < 0)) throw new TypeError('maxFindings must be a nonnegative safe integer.');
  const authored = inspectMissingDescriptions(document);
  const byOwner = new Map<string, MissingDescriptionFinding[]>();
  for (const item of authored.findings) {
    const key = item.slideId !== undefined ? `slide:${item.slideId}` : item.layoutId !== undefined ? `layout:${item.layoutId}` : `master:${item.masterId}`;
    const entries = byOwner.get(key) ?? []; entries.push(item); byOwner.set(key, entries);
  }
  const layouts = new Map(document.layouts.map(layout => [layout.id, layout]));
  const findings: EffectiveDescriptionFinding[] = [];
  const add = (slideId: string, kind: 'slide' | 'layout' | 'master', id: string): void => {
    for (const item of byOwner.get(`${kind}:${id}`) ?? []) {
      if (options.maxFindings !== undefined && findings.length >= options.maxFindings) throw new EffectiveDescriptionLimitError(options.maxFindings);
      findings.push(Object.freeze({ code: item.code, elementType: item.elementType, elementId: item.elementId, path: item.path, slideId, source: Object.freeze({ kind, id }) }));
    }
  };
  for (const slide of document.slides) {
    if (slide.hidden) continue;
    const layout = layouts.get(slide.layoutId ?? DEFAULT_LAYOUT_ID);
    if (slide.showInheritedShapes !== false && layout) {
      if (layout.showMasterShapes !== false && layout.masterId !== undefined) add(slide.id, 'master', layout.masterId);
      add(slide.id, 'layout', layout.id);
    }
    add(slide.id, 'slide', slide.id);
  }
  return Object.freeze({ scope: 'modeled-slide-occurrences', findings: Object.freeze(findings) });
}
