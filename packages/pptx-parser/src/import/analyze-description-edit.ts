import type { NormalizedElement, NormalizedPresentation } from '@pptkit/core';
import { PptxImportError, type ImportDiagnostic, type PptxSaveChange, type PptxSessionLimits } from '../types/import.js';
import { equalState } from './state.js';
import type { DrawingDescriptionBinding } from './drawing-description-binding.js';

export interface DescriptionEditAnalysis {
  readonly changes: readonly PptxSaveChange[];
  readonly diagnostics: readonly ImportDiagnostic[];
  readonly descriptions: readonly { readonly binding: DrawingDescriptionBinding; readonly value?: string }[];
  readonly work?: number;
  readonly remainder?: NormalizedPresentation;
}

/** Isolate bound description changes; reject conflicts and all other snapshot changes. */
export async function analyzeDescriptionEdit(baseline: NormalizedPresentation, candidate: NormalizedPresentation, bindings: readonly DrawingDescriptionBinding[], limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>, allowBatch = false, allowOtherChanges = false): Promise<DescriptionEditAnalysis> {
  let work = 0;
  const charge = async (amount: number) => {
    work += amount;
    if (work > limits.maxSaveAnalysisWork) throw new PptxImportError('save-limit-exceeded', 'Description analysis exceeds maxSaveAnalysisWork.');
    await checkpoint();
  };
  const key = (slide: string, id: string) => JSON.stringify([slide, id]);
  const index = async (state: NormalizedPresentation) => {
    const result = new Map<string, { element: NormalizedElement; path: string }>();
    const stack: { slide: string; element: NormalizedElement; path: string }[] = [];
    state.slides.forEach((slide, i) => slide.elements.forEach((element, j) => stack.push({ slide: slide.id, element, path: `/slides/${i}/elements/${j}` })));
    while (stack.length) {
      const { slide, element, path } = stack.pop()!; await charge(path.length + 1);
      result.set(key(slide, element.id), { element, path });
      if (element.type === 'group') element.children.forEach((child, i) => stack.push({ slide, element: child, path: `${path}/children/${i}` }));
    }
    return result;
  };
  const expected = structuredClone(baseline), original = await index(expected), next = await index(candidate);
  const selected: { binding: DrawingDescriptionBinding; value?: string; path: string }[] = [];
  const sourceValues = new Map<string, string | undefined>();
  for (const binding of bindings) {
    await charge(binding.part.length + binding.elementId.length + 1);
    if (binding.status !== 'bound' || binding.elementIndex === undefined) continue;
    const before = original.get(key(binding.slideId, binding.elementId)), after = next.get(key(binding.slideId, binding.elementId));
    if (!before || !after || before.path !== after.path || (binding.property === 'decorative' ? before.element.accessibility.decorative !== binding.expectedBoolean : (binding.property === 'name' ? before.element.name : before.element.accessibility.description) !== binding.expected)) continue;
    const value = binding.property === 'decorative' ? after.element.accessibility.decorative ? '1' : '0' : binding.property === 'name' ? after.element.name : after.element.accessibility.description;
    const sourceKey = JSON.stringify([binding.part, binding.elementIndex, binding.property ?? 'descr']);
    if (sourceValues.has(sourceKey) && sourceValues.get(sourceKey) !== value) return { changes: [], descriptions: [], diagnostics: [{ code: 'save-source-shared-conflict', severity: 'error', source: { part: binding.part, elementIndex: binding.elementIndex }, message: 'Views of one source drawing property disagree.' }] };
    sourceValues.set(sourceKey, value);
    if ((binding.property === 'decorative' ? after.element.accessibility.decorative === binding.expectedBoolean : value === binding.expected) || value === undefined && !binding.allowPresenceChange || value !== undefined && typeof value !== 'string') continue;
    await charge((value?.length ?? 0) + 1);
    if (selected.length && !allowBatch) return { changes: [], descriptions: [], diagnostics: [{ code: 'save-description-edit-unsupported', severity: 'error', message: 'The selected description profile admits exactly one changed description.' }] };
    if (selected.length >= limits.maxSaveChanges) throw new PptxImportError('save-limit-exceeded', 'Description edits exceed maxSaveChanges.');
    if (binding.property === 'decorative') before.element.accessibility.decorative = value === '1'; else if (binding.property === 'name') { if (value === undefined) continue; before.element.name = value; } else if (value === undefined) delete before.element.accessibility.description; else before.element.accessibility.description = value;
    selected.push({ binding, ...(value !== undefined ? { value } : {}), path: binding.property === 'name' ? `${before.path}/name` : `${before.path}/accessibility/${binding.property === 'decorative' ? 'decorative' : 'description'}` });
  }
  if (!allowOtherChanges && (!selected.length || !equalState(expected, candidate, limits.maxSnapshotValues))) return { changes: [], descriptions: [], diagnostics: [{ code: bindings.some(binding => binding.property === 'decorative') ? 'save-decorative-edit-unsupported' : bindings.some(binding => binding.property === 'name') ? 'save-name-edit-unsupported' : 'save-description-edit-unsupported', severity: 'error', message: 'The selected drawing property profile admits only bound changes with all other normalized state unchanged.' }] };
  selected.sort((a, b) => a.path < b.path ? -1 : a.path > b.path ? 1 : 0);
  const changes: PptxSaveChange[] = selected.map(({ binding, path }) => ({ kind: 'metadata', documentPath: path, slideId: binding.slideId, elementId: binding.elementId, source: { part: binding.part, elementIndex: binding.elementIndex!, sourceObjectId: binding.sourceObjectId } }));
  const unique = new Map<string, { binding: DrawingDescriptionBinding; value?: string }>();
  for (const { binding, value } of selected) unique.set(JSON.stringify([binding.part, binding.elementIndex, binding.property ?? 'descr']), { binding, ...(value !== undefined ? { value } : {}) });
  let remainder: NormalizedPresentation | undefined;
  if (allowOtherChanges) {
    remainder = structuredClone(candidate);
    const remainderIndex = await index(remainder);
    for (const { binding } of selected) {
      await charge(binding.elementId.length + 1);
      const target = remainderIndex.get(key(binding.slideId, binding.elementId));
      if (!target) throw new PptxImportError('save-candidate-invalid', 'Description owner disappeared while isolating the edit.');
      if (binding.property === 'decorative') target.element.accessibility.decorative = binding.expectedBoolean!; else if (binding.property === 'name') target.element.name = binding.expected!; else if (binding.expected === undefined) delete target.element.accessibility.description; else target.element.accessibility.description = binding.expected;
    }
  }
  return { changes, descriptions: [...unique.values()], diagnostics: [], work, ...(remainder ? { remainder } : {}) };
}
