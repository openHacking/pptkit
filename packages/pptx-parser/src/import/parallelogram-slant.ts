import { attribute, children } from '../presentation/structure-xml.js';
import type { XmlElement } from '../xml/read-xml.js';

type Result = { supported: true; value?: number; list?: XmlElement; guide?: XmlElement } | { supported: false };
const xmlns = 'http://www.w3.org/2000/xmlns/';
const whitespaceOnly = (node: XmlElement) => node.children.every(child => child.kind === 'element' || (child.kind === 'text' && !child.value.trim()));
const noForeignAttributes = (node: XmlElement, allowed: readonly string[]) => node.attributes.every(a => a.namespaceUri === xmlns || (a.namespaceUri === '' && allowed.includes(a.local)));

/** Bounded literal slant profile; formulas and additional/unknown adjustment content remain source-owned. */
export function readParallelogramSlant(preset: XmlElement, namespace: string): Result {
  const lists = children(preset).filter(node => node.namespaceUri === namespace && node.local === 'avLst');
  if (lists.length > 1) return { supported: false };
  const list = lists[0];
  if (!list) return { supported: true };
  if (!noForeignAttributes(list, []) || !whitespaceOnly(list)) return { supported: false };
  const guides = children(list);
  if (!guides.length) return { supported: true, list };
  if (guides.length !== 1) return { supported: false };
  const guide = guides[0]!;
  if (guide.namespaceUri !== namespace || guide.local !== 'gd' || children(guide).length || !whitespaceOnly(guide) || !noForeignAttributes(guide, ['name', 'fmla']) || attribute(guide, 'name') !== 'adj') return { supported: false };
  const formula = attribute(guide, 'fmla')?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, '').match(/^val[\t\r\n ]+([+-]?\d+)$/);
  if (!formula) return { supported: false };
  const value = Number(formula[1]);
  if (!Number.isSafeInteger(value) || value < 0 || value > 100000) return { supported: false };
  return { supported: true, value: value === 0 ? 0 : value / 100000, list, guide };
}
