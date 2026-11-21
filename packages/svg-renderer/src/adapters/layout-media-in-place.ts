import type { LayoutResult } from '@pptkit/layout';
import { findLayoutMedia } from './find-layout-media.js';
import { mountLayoutMedia, type BrowserLayoutMedia, type BrowserLayoutMediaOptions } from './layout-media.js';
import { groupTransform, shapeTransform, validateComposite, validateGeometry } from './layout-media-overlay.js';

const SVG = 'http://www.w3.org/2000/svg', XHTML = 'http://www.w3.org/1999/xhtml';

/** Insert a media layer adjacent to its static SVG frame, preserving its drawing order. */
export async function mountLayoutMediaInPlace(svg: SVGSVGElement, layout: LayoutResult, options: BrowserLayoutMediaOptions): Promise<BrowserLayoutMedia> {
  if (!svg || svg.namespaceURI !== SVG || svg.localName !== 'svg' || !svg.isConnected || !svg.ownerDocument.defaultView) throw new TypeError('A connected SVG slide is required.');
  if (!options || typeof options.slideId !== 'string' || !options.slideId || typeof options.elementId !== 'string' || !options.elementId) throw new TypeError('slideId and elementId are required.');
  const { target, parents } = findLayoutMedia(layout, options.slideId, options.elementId);
  if (parents.length > 256) throw new RangeError('Media overlay exceeds the group depth limit.');
  if (!layout.size || !Number.isFinite(layout.size.width) || !Number.isFinite(layout.size.height) || layout.size.width <= 0 || layout.size.height <= 0 || svg.getAttribute('data-pptkit-slide-id') !== options.slideId || !svg.hasAttribute('viewBox')) throw new TypeError('SVG slide identity or size does not match the Layout result.');
  const viewBox = svg.getAttribute('viewBox')!.trim().split(/[\s,]+/u).map(Number);
  if (viewBox.length !== 4 || viewBox[0] !== 0 || viewBox[1] !== 0 || viewBox[2] !== layout.size.width || viewBox[3] !== layout.size.height) throw new TypeError('SVG viewBox does not match the Layout size.');
  validateGeometry(target.box, target.transform, target.opacity);
  for (const group of parents) groupTransform(group);
  validateComposite(parents, target.box, target.transform);
  if (target.action !== undefined || parents.some(group => group.action !== undefined)) throw new TypeError('Media or containing group actions are not supported by in-place playback.');
  const decorative = target.accessibility.decorative || parents.some(group => group.accessibility.decorative);
  const controls = options.controls ?? false, label = target.accessibility.description ?? target.name;
  if (decorative && controls) throw new TypeError('Decorative media cannot expose native controls.');
  const overlays = Array.from(svg.querySelectorAll('[data-pptkit-media-overlay]'));
  if (overlays.some(node => node.getAttribute('data-pptkit-media-overlay') === options.elementId)) throw new TypeError('A media overlay for this object already exists.');
  const matches = Array.from(svg.querySelectorAll('[data-pptkit-element-id]')).filter(node => node.getAttribute('data-pptkit-element-id') === options.elementId);
  if (matches.length !== 1 || matches[0]!.namespaceURI !== SVG || matches[0]!.localName !== 'g') throw new TypeError('SVG media frame is missing or ambiguous.');
  const poster = matches[0] as SVGGElement;
  const ancestors: Element[] = []; let node: Element | null = poster.parentNode as Element | null;
  while (node && node !== svg) {
    if (node.namespaceURI !== SVG || node.localName === 'a') throw new TypeError('SVG media frame has unsupported ancestors.');
    if (node.hasAttribute('data-pptkit-element-id')) { if (node.hasAttribute('style')) throw new TypeError('SVG media frame has unverified group styling.'); ancestors.push(node); }
    else if (node.hasAttribute('transform') || node.hasAttribute('opacity') || node.hasAttribute('style')) throw new TypeError('SVG media frame has unverified wrapper geometry.');
    node = node.parentNode as Element | null;
  }
  ancestors.reverse();
  if (node !== svg || ancestors.length !== parents.length || ancestors.some((element, index) => element.getAttribute('data-pptkit-element-id') !== parents[index]!.id || element.getAttribute('transform') !== groupTransform(parents[index]!) || !element.hasAttribute('opacity') || Number(element.getAttribute('opacity')) !== parents[index]!.opacity)) throw new TypeError('SVG group chain does not match Layout.');
  const expectedTransform = target.transform.rotation || target.transform.flipH || target.transform.flipV ? shapeTransform(target.box, target.transform) : null;
  if (poster.getAttribute('transform') !== expectedTransform || !poster.hasAttribute('opacity') || Number(poster.getAttribute('opacity')) !== target.opacity || poster.hasAttribute('style')) throw new TypeError('SVG media frame does not match Layout.');
  const geometry = Array.from(poster.children).find(child => child.namespaceURI === SVG && (child.localName === 'rect' || child.localName === 'image'));
  if (!geometry || (['x', 'y', 'width', 'height'] as const).some(key => !geometry.hasAttribute(key) || Number(geometry.getAttribute(key)) !== target.box[key])) throw new TypeError('SVG media frame geometry does not match Layout.');
  const parent = poster.parentNode;
  if (!parent) throw new TypeError('SVG media frame is detached.');
  const doc = svg.ownerDocument, root = doc.createElementNS(SVG, 'g');
  root.setAttribute('data-pptkit-media-overlay', options.elementId); root.setAttribute('transform', shapeTransform(target.box, target.transform)); root.setAttribute('opacity', String(target.opacity));
  const foreign = doc.createElementNS(SVG, 'foreignObject');
  for (const key of ['x', 'y', 'width', 'height'] as const) foreign.setAttribute(key, String(target.box[key]));
  const host = doc.createElementNS(XHTML, 'div') as HTMLDivElement; host.style.width = '100%'; host.style.height = '100%'; foreign.appendChild(host); root.appendChild(foreign);
  const originalStyle = poster.getAttribute('style');
  const restore = () => { if (originalStyle === null) poster.removeAttribute('style'); else poster.setAttribute('style', originalStyle); };
  let mounted: BrowserLayoutMedia | undefined;
  try {
    parent.insertBefore(root, poster.nextSibling);
    const originalMatrix = poster.getCTM(), overlayMatrix = root.getCTM(), view = doc.defaultView!;
    const originalStyle = view.getComputedStyle(poster), overlayStyle = view.getComputedStyle(root);
    const matrixMatches = originalMatrix && overlayMatrix && (['a', 'b', 'c', 'd', 'e', 'f'] as const).every(key => Math.abs(originalMatrix[key] - overlayMatrix[key]) <= 1e-9 * Math.max(1, Math.abs(originalMatrix[key]), Math.abs(overlayMatrix[key])));
    if (!matrixMatches || originalStyle.visibility !== 'visible' || originalStyle.display === 'none' || originalStyle.opacity !== overlayStyle.opacity) throw new TypeError('SVG media frame has unverified computed styling.');
    mounted = await mountLayoutMedia(host, layout, options);
    if (mounted.controller.state.status === 'disposed') { root.remove(); return mounted; }
    const media = mounted.controller.element; media.style.width = '100%'; media.style.height = '100%'; media.style.display = 'block';
    if (mounted.kind === 'video') media.style.objectFit = 'fill';
    if (decorative) { media.setAttribute('aria-hidden', 'true'); media.tabIndex = -1; }
    else media.setAttribute('aria-label', label);
    mounted.controller.subscribe(state => {
      if (state.status === 'disposed') { restore(); root.remove(); return; }
      const visible = state.status !== 'loading' && state.status !== 'error';
      media.style.visibility = visible ? 'visible' : 'hidden';
      if (visible && (mounted!.kind === 'video' || controls)) poster.setAttribute('style', 'visibility: hidden !important'); else restore();
    });
    return mounted;
  } catch (error) { mounted?.controller.dispose(); restore(); root.remove(); throw error; }
}
