import { snapshotOptions } from '../opc/read-package.js';
import { structureContext } from '../presentation/structure-context.js';
import { readStructureFromContext } from '../presentation/read-structure.js';
import { attribute, children } from '../presentation/structure-xml.js';
import { PackageInspectionError } from '../types/archive.js';
import type { PptxCaptionInspection, PptxCaptionInspectionLimits, PptxCaptionInspectionOptions, PptxCaptionTrackSource } from '../types/captions.js';
import type { XmlElement } from '../xml/read-xml.js';

const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const P173 = 'http://schemas.microsoft.com/office/powerpoint/2017/3/main';
const REL = 'http://schemas.microsoft.com/office/2017/04/relationships/track';
const MEDIA = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
const TRACKS = '{3AFAAA56-56D3-431D-BCD4-E75A35582382}';
interface Meta { node: XmlElement; index: number; ordinal: number; parent?: Meta }

function limits(value?: Partial<PptxCaptionInspectionLimits>): PptxCaptionInspectionLimits {
  const defaults = { maxWorkUnits: 1000000, maxTracks: 10000, maxTextCharacters: 1000000, maxTrackBytes: 1048576, maxTotalTrackBytes: 8388608 };
  if (value !== undefined && (!value || typeof value !== 'object' || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value)))) throw new TypeError('Invalid caption limits.');
  const captured: Record<string, number> = { ...defaults };
  if (value) for (const key of Reflect.ownKeys(value)) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
    if (typeof key !== 'string' || !Object.hasOwn(defaults, key) || !('value' in descriptor)) throw new TypeError('Unknown or accessor caption limit.');
    if (descriptor.value !== undefined) captured[key] = descriptor.value;
  }
  for (const [key, amount] of Object.entries(captured)) if (!Number.isSafeInteger(amount) || amount <= 0) throw new RangeError(`Invalid caption limit ${key}.`);
  return captured as unknown as PptxCaptionInspectionLimits;
}

/** Read the observed raw source form; no active-branch, editing or playback semantics are inferred. */
export async function inspectPptxCaptionTracks(bytes: Uint8Array, options: PptxCaptionInspectionOptions = {}): Promise<PptxCaptionInspection> {
  const budget = limits(options.captionLimits), captured = snapshotOptions(options);
  const c = await structureContext(bytes, { limits: captured.limits, zipLimits: captured.zipLimits, xmlLimits: captured.xmlLimits, ...(captured.signal ? { signal: captured.signal } : {}) }, true);
  const structure = await readStructureFromContext(c), tracks: PptxCaptionTrackSource[] = [], cache = new Map<string, Pick<PptxCaptionTrackSource, 'status' | 'contentType' | 'byteLength'>>();
  let workUnits = 0, textCharacters = 0, inspectedBytes = 0;
  const fail = (message: string): never => { throw new PackageInspectionError('package-limit-exceeded', message); };
  const step = async () => { if (++workUnits > budget.maxWorkUnits) fail('Caption inventory exceeds maxWorkUnits.'); await c.checkpoint(); };
  const text = (value: string): string => { textCharacters += value.length; if (textCharacters > budget.maxTextCharacters) fail('Caption inventory exceeds maxTextCharacters.'); return value; };
  const is = (meta: Meta | undefined, ns: string, local: string) => meta?.node.namespaceUri === ns && meta.node.local === local;
  for (const slide of structure.slides) {
    const effective = await c.root(slide.part, 'sld', 'presentationml.slide'), cSld = children(effective).find(node => node.namespaceUri === c.p && node.local === 'cSld');
    const tree = cSld && children(cSld).find(node => node.namespaceUri === c.p && node.local === 'spTree'), pictures = new Set<number>();
    const drawings = tree ? children(tree) : [];
    while (drawings.length) {
      await step(); const node = drawings.pop()!;
      if (node.namespaceUri !== c.p) continue;
      if (node.local === 'pic') pictures.add((await c.sourceElement(slide.part, node)).elementIndex);
      else if (node.local === 'grpSp') for (const child of children(node)) { await step(); drawings.push(child); }
    }
    const raw = c.rawXml.get(slide.part)!; let index = 0;
    const stack: Meta[] = [{ node: raw, index: 0, ordinal: 1 }];
    while (stack.length) {
      await step(); const item = stack.pop()!; item.index = ++index;
      const list = item.parent, info = list?.parent, ext = info?.parent, extList = ext?.parent, media = extList?.parent, outer = media?.parent, outerList = outer?.parent, nvPr = outerList?.parent, nvPic = nvPr?.parent, picture = nvPic?.parent;
      if (is(item, P173, 'track') && is(list, P173, 'trackLst') && is(info, P173, 'tracksInfo') && is(ext, c.p, 'ext') && attribute(ext!.node, 'uri')?.toUpperCase() === TRACKS && is(extList, P14, 'extLst') && is(media, P14, 'media') && is(outer, c.p, 'ext') && attribute(outer!.node, 'uri')?.toUpperCase() === MEDIA && is(outerList, c.p, 'extLst') && is(nvPr, c.p, 'nvPr') && is(nvPic, c.p, 'nvPicPr') && is(picture, c.p, 'pic')) {
        if (tracks.length >= budget.maxTracks) fail('Caption inventory exceeds maxTracks.');
        const components: string[] = [];
        for (let current: Meta | undefined = item; current; current = current.parent) { await step(); components.push(text(`Q{${current.node.namespaceUri}}${current.node.local}[${current.ordinal}]`)); }
        const relationshipId = attribute(item.node, 'embed', c.r), trackId = attribute(item.node, 'id'), language = attribute(item.node, 'lang'), label = attribute(item.node, 'label');
        const identity = children(nvPic!.node).find(node => node.namespaceUri === c.p && node.local === 'cNvPr'), objectId = identity && attribute(identity, 'id');
        const base = { slideId: slide.id, source: { part: text(slide.part), elementIndex: item.index, path: '/' + components.reverse().join('/') }, pictureView: pictures.has(picture!.index) ? 'effective-drawing' as const : 'raw-only' as const, ...(objectId === undefined ? {} : { objectId: text(objectId) }), ...(trackId === undefined ? {} : { trackId: text(trackId) }), ...(language === undefined ? {} : { language: text(language) }), ...(label === undefined ? {} : { label: text(label) }), ...(relationshipId === undefined ? {} : { relationshipId: text(relationshipId) }) };
        const relation = c.findRelationship(slide.part, relationshipId);
        if (!relation || relation.type !== REL || relation.fragment !== undefined || (relation.targetMode === 'Internal' && !relation.targetPart)) tracks.push({ ...base, status: 'unsupported-relationship' });
        else if (relation.targetMode === 'External') tracks.push({ ...base, status: 'external', externalTarget: text(relation.target) });
        else {
          const targetPart = text(relation.targetPart!), payload = c.payload(targetPart); let result = cache.get(targetPart);
          if (!result) {
            const length = payload.bytes.length;
            if (length > budget.maxTrackBytes || inspectedBytes + length > budget.maxTotalTrackBytes) fail('Caption inventory exceeds track byte limits.');
            inspectedBytes += length; const contentType = text(payload.contentType.trim().toLowerCase());
            if (contentType !== 'text/vtt') result = { status: 'unsupported-content-type', contentType, byteLength: length };
            else {
              try { const decoded = new TextDecoder('utf-8', { fatal: true }).decode(payload.bytes); result = { status: /^WEBVTT(?:\r\n|\n)/u.test(decoded) ? 'vtt-header-only' : 'invalid-vtt-header', contentType, byteLength: length }; }
              catch (error) { if (!(error instanceof TypeError)) throw error; result = { status: 'invalid-utf8', contentType, byteLength: length }; }
            }
            cache.set(targetPart, result);
          }
          if (result.contentType !== undefined) text(result.contentType);
          tracks.push({ ...base, targetPart, ...result });
        }
      }
      const counts = new Map<string, number>(), nested: Meta[] = [];
      for (const child of children(item.node)) { await step(); const key = JSON.stringify([child.namespaceUri, child.local]), ordinal = (counts.get(key) ?? 0) + 1; counts.set(key, ordinal); nested.push({ node: child, parent: item, ordinal, index: 0 }); }
      for (let i = nested.length - 1; i >= 0; i--) stack.push(nested[i]!);
    }
  }
  c.abort();
  for (const track of tracks) { Object.freeze(track.source); Object.freeze(track); }
  return Object.freeze({ schemaVersion: 'pptkit-caption-source-v1', dialect: structure.dialect, view: 'raw-only', tracks: Object.freeze(tracks), workUnits, inspectedBytes });
}
