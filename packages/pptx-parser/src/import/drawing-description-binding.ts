import type { structureContext } from '../presentation/structure-context.js';
import type { XmlElement } from '../xml/read-xml.js';

/** Private source binding evidence; package guards and candidate validation still gate saving. */
export interface DrawingDescriptionBinding {
  readonly property?: 'name' | 'decorative';
  readonly propertyIndex?: number;
  readonly listIndex?: number;
  readonly extensionIndex?: number;
  readonly expectedBoolean?: boolean;
  readonly slideId: string;
  readonly elementId: string;
  readonly part: string;
  readonly sourceObjectId: string;
  readonly namespace: string;
  readonly allowPresenceChange: boolean;
  readonly status: 'bound' | 'unmapped';
  readonly elementIndex?: number;
  readonly nonvisualIndex?: number;
  readonly ownerIndex?: number;
  readonly expected?: string;
  readonly reason?: string;
}

/** Require raw ownership, expanded names and the same source attribute value. */
export async function bindDrawingDescription(
  context: Awaited<ReturnType<typeof structureContext>>,
  part: string, slideId: string, elementId: string, sourceObjectId: string,
  owner: XmlElement, nonvisual: XmlElement, properties: XmlElement, projectedDescription: string | undefined, allowPresenceChange = false, property?: 'name',
): Promise<DrawingDescriptionBinding> {
  const base = { slideId, elementId, part, sourceObjectId, namespace: context.p, allowPresenceChange, ...(property ? { property } : {}) };
  const [rawOwner, rawNonvisual, rawProperties] = await Promise.all([
    context.sourceElement(part, owner), context.sourceElement(part, nonvisual), context.sourceElement(part, properties),
  ]);
  const indices = { ownerIndex: rawOwner.elementIndex, nonvisualIndex: rawNonvisual.elementIndex, elementIndex: rawProperties.elementIndex };
  const containers: Record<string, string> = { sp: 'nvSpPr', grpSp: 'nvGrpSpPr', pic: 'nvPicPr', cxnSp: 'nvCxnSpPr', graphicFrame: 'nvGraphicFramePr' };
  const expectedContainer = containers[rawOwner.node.local];
  if (!expectedContainer || rawOwner.node.namespaceUri !== context.p || rawNonvisual.node.namespaceUri !== context.p || rawNonvisual.node.local !== expectedContainer || rawProperties.node.namespaceUri !== context.p || rawProperties.node.local !== 'cNvPr' || rawNonvisual.parent !== rawOwner.node || rawProperties.parent !== rawNonvisual.node) return { ...base, ...indices, status: 'unmapped', reason: 'Raw object/nonvisual/property ancestry or dialect does not match the effective owner.' };
  const id = rawProperties.node.attributes.find(item => item.namespaceUri === '' && item.local === 'id')?.value;
  const description = rawProperties.node.attributes.find(item => item.namespaceUri === '' && item.local === (property ?? 'descr'))?.value;
  if (id === undefined || !/^\+?\d+$/.test(id.trim()) || !Number.isSafeInteger(Number(id)) || Number(id) < 0 || Number(id) > 4294967295 || String(Number(id)) !== sourceObjectId) return { ...base, ...indices, status: 'unmapped', reason: 'Raw drawing identity does not match the selected source object.' };
  if (description !== projectedDescription || description === undefined && !allowPresenceChange) return { ...base, ...indices, status: 'unmapped', reason: property === 'name' ? 'Raw name presence and value must match the selected projection profile.' : 'Raw description presence and value must match the selected projection profile.' };
  return { ...base, ...indices, status: 'bound', ...(description !== undefined ? { expected: description } : {}) };
}

/** Bind an existing projected decorative value through its unmodified raw ancestry. */
export async function bindDrawingDecorative(
  context: Awaited<ReturnType<typeof structureContext>>, part: string, slideId: string, elementId: string, sourceObjectId: string,
  owner: XmlElement, nonvisual: XmlElement, properties: XmlElement, projectedDescription: string | undefined, decorative: XmlElement | undefined, projectedValue: boolean | undefined,
): Promise<DrawingDescriptionBinding> {
  const base = await bindDrawingDescription(context, part, slideId, elementId, sourceObjectId, owner, nonvisual, properties, projectedDescription, true);
  const unavailable = (reason: string): DrawingDescriptionBinding => ({ ...base, property: 'decorative', status: 'unmapped', reason });
  if (base.status !== 'bound' || !decorative || projectedValue === undefined) return unavailable('Existing decorative value and direct drawing ownership are required.');
  const elements = (node: XmlElement) => node.children.filter((n): n is XmlElement => n.kind === 'element');
  const lists = elements(properties).filter(n => n.namespaceUri === context.a && n.local === 'extLst');
  if (lists.length !== 1) return unavailable('Decorative extension list is not unique.');
  const extensions = elements(lists[0]!).filter(n => n.namespaceUri === context.a && n.local === 'ext' && n.attributes.some(a => a.namespaceUri === '' && a.local === 'uri' && a.value.toUpperCase() === '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'));
  if (extensions.length !== 1 || !elements(extensions[0]!).includes(decorative)) return unavailable('Decorative extension ownership is not unique.');
  const [list, extension, target] = await Promise.all([context.sourceElement(part, lists[0]!), context.sourceElement(part, extensions[0]!), context.sourceElement(part, decorative)]);
  const rawProperties = await context.sourceElement(part, properties);
  const value = target.node.attributes.find(a => a.namespaceUri === '' && a.local === 'val')?.value;
  const normalized = value?.trim();
  const xmlns = 'http://www.w3.org/2000/xmlns/';
  const meaningfulText = (n: XmlElement) => n.children.some(c => (c.kind === 'text' || c.kind === 'cdata') && !/^[\t\r\n ]*$/.test(c.value));
  const rawLists = elements(rawProperties.node).filter(n => n.namespaceUri === context.a && n.local === 'extLst');
  const rawExtensions = elements(list.node).filter(n => n.namespaceUri === context.a && n.local === 'ext' && n.attributes.some(a => a.namespaceUri === '' && a.local === 'uri' && a.value.toUpperCase() === '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'));
  if (rawLists.length !== 1 || rawLists[0] !== list.node || rawExtensions.length !== 1 || rawExtensions[0] !== extension.node || elements(extension.node).length !== 1 || elements(extension.node)[0] !== target.node || elements(target.node).length !== 0 || meaningfulText(extension.node) || meaningfulText(target.node) || extension.node.attributes.some(a => a.namespaceUri !== xmlns && (a.namespaceUri !== '' || a.local !== 'uri')) || target.node.attributes.some(a => a.namespaceUri !== xmlns && (a.namespaceUri !== '' || a.local !== 'val'))) return unavailable('Raw decorative extension is ambiguous or contains unsupported content.');
  if (list.parent !== rawProperties.node || extension.parent !== list.node || target.parent !== extension.node || list.node.namespaceUri !== context.a || list.node.local !== 'extLst' || extension.node.namespaceUri !== context.a || extension.node.local !== 'ext' || target.node.namespaceUri !== 'http://schemas.microsoft.com/office/drawing/2017/decorative' || target.node.local !== 'decorative' || !['0', '1', 'true', 'false'].includes(normalized ?? '') || (normalized === '1' || normalized === 'true') !== projectedValue) return unavailable('Raw decorative ancestry or value disagrees with the effective projection.');
  return { ...base, property: 'decorative', elementIndex: target.elementIndex, propertyIndex: rawProperties.elementIndex, listIndex: list.elementIndex, extensionIndex: extension.elementIndex, expected: value!, expectedBoolean: projectedValue, allowPresenceChange: false, status: 'bound' };
}
