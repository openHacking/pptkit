import type { PresentationElement, PresentationElementInput } from '../types/element.js';
import type { SlideLayoutDefinition, SlideMasterDefinition } from '../types/presentation.js';
import { deepClone } from '../utils/clone.js';
import { remapCopiedReferences } from './copy-references.js';

function equal(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object' || Array.isArray(left) !== Array.isArray(right)) return false;
  const a = Object.keys(left).sort(), b = Object.keys(right).sort();
  if (a.length !== b.length || a.some((key, index) => key !== b[index])) return false;
  return a.every(key => equal((left as Record<string, unknown>)[key], (right as Record<string, unknown>)[key]));
}

/** Exact ordered structural equivalence, allowing only owned element identities to differ. */
export function equivalentDefinitionElements(candidate: SlideLayoutDefinition | SlideMasterDefinition, existing: SlideLayoutDefinition | SlideMasterDefinition): Map<string, string> | undefined {
  const { id: _a, elements: left, ...a } = candidate;
  const { id: _b, elements: right, ...b } = existing;
  if (!equal(a, b) || (left === undefined) !== (right === undefined)) return undefined;
  if (left === undefined || right === undefined) return new Map();
  const ids = new Map<string, string>();
  function pair(x: readonly PresentationElement[], y: readonly PresentationElement[]): boolean {
    if (x.length !== y.length) return false;
    for (let i = 0; i < x.length; i++) {
      const p = x[i]!, q = y[i]!;
      if (p.type !== q.type) return false;
      ids.set(p.id, q.id);
      if (p.type === 'group' && (q.type !== 'group' || !pair(p.children, q.children))) return false;
    }
    return true;
  }
  if (!pair(left, right)) return undefined;
  const mapped = deepClone(left) as PresentationElementInput[];
  remapCopiedReferences(mapped, ids, new Map());
  function remapIds(elements: PresentationElementInput[]): void {
    for (const element of elements) {
      element.id = ids.get(element.id!)!;
      if (element.type === 'group') remapIds(element.children);
    }
  }
  remapIds(mapped);
  return equal(mapped, right) ? ids : undefined;
}
