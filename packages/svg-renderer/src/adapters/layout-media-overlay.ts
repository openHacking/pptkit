import type { Box, NormalizedTransform } from '@pptkit/core';
import type { LayoutGroupElement, LayoutResult } from '@pptkit/layout';
import { findLayoutMedia } from './find-layout-media.js';
import { mountLayoutMedia, type BrowserLayoutMedia, type BrowserLayoutMediaOptions } from './layout-media.js';

const SVG = 'http://www.w3.org/2000/svg', XHTML = 'http://www.w3.org/1999/xhtml';

function validateGeometry(box: Box, transform: NormalizedTransform, opacity: number): void {
  if (!box || ![box.x, box.y, box.width, box.height].every(Number.isFinite) || box.width <= 0 || box.height <= 0 || !transform || !Number.isFinite(transform.rotation) || typeof transform.flipH !== 'boolean' || typeof transform.flipV !== 'boolean' || !Number.isFinite(opacity) || opacity < 0 || opacity > 1) throw new TypeError('Media overlay geometry, transform or opacity is invalid.');
}
function shapeTransform(box: Box, transform: NormalizedTransform): string {
  const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
  return `translate(${cx} ${cy}) rotate(${transform.rotation}) scale(${transform.flipH ? -1 : 1} ${transform.flipV ? -1 : 1}) translate(${-cx} ${-cy})`;
}
type Matrix = readonly [number, number, number, number, number, number];
function multiply(left: Matrix, right: Matrix): Matrix {
  const [a, b, c, d, e, f] = left, [g, h, i, j, k, l] = right;
  const value: Matrix = [a * g + c * h, b * g + d * h, a * i + c * j, b * i + d * j, a * k + c * l + e, b * k + d * l + f];
  if (!value.every(Number.isFinite)) throw new RangeError('Media overlay transform exceeds finite geometry.');
  return value;
}
function shapeMatrix(box: Box, transform: NormalizedTransform): Matrix {
  const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
  const angle = transform.rotation * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle), fx = transform.flipH ? -1 : 1, fy = transform.flipV ? -1 : 1;
  const a = cos * fx, b = sin * fx, c = -sin * fy, d = cos * fy;
  const value: Matrix = [a, b, c, d, cx - a * cx - c * cy, cy - b * cx - d * cy];
  if (!value.every(Number.isFinite)) throw new RangeError('Media overlay transform exceeds finite geometry.');
  return value;
}
function validateComposite(parents: readonly LayoutGroupElement[], box: Box, transform: NormalizedTransform): void {
  let matrix: Matrix = [1, 0, 0, 1, 0, 0];
  for (const parent of parents) {
    matrix = multiply(matrix, shapeMatrix(parent.box, parent.transform));
    matrix = multiply(matrix, [parent.box.width / parent.coordinateSize.width, 0, 0, parent.box.height / parent.coordinateSize.height, parent.box.x, parent.box.y]);
  }
  matrix = multiply(matrix, shapeMatrix(box, transform));
  for (const x of [box.x, box.x + box.width]) for (const y of [box.y, box.y + box.height]) {
    if (![x, y, matrix[0] * x + matrix[2] * y + matrix[4], matrix[1] * x + matrix[3] * y + matrix[5]].every(Number.isFinite)) throw new RangeError('Media overlay bounds exceed finite geometry.');
  }
}
function groupTransform(group: LayoutGroupElement): string {
  const { box, coordinateSize, transform, opacity } = group;
  validateGeometry(box, transform, opacity);
  if (!coordinateSize || !Number.isFinite(coordinateSize.width) || !Number.isFinite(coordinateSize.height) || coordinateSize.width <= 0 || coordinateSize.height <= 0 || !Number.isFinite(box.width / coordinateSize.width) || !Number.isFinite(box.height / coordinateSize.height)) throw new TypeError('Media overlay group coordinates are invalid.');
  return `${shapeTransform(box, transform)} translate(${box.x} ${box.y}) scale(${box.width / coordinateSize.width} ${box.height / coordinateSize.height})`;
}

/** Append a top-layer foreignObject in a matching rendered SVG; no automatic playback. */
export async function mountLayoutMediaOverlay(svg: SVGSVGElement, layout: LayoutResult, options: BrowserLayoutMediaOptions): Promise<BrowserLayoutMedia> {
  if (!svg || svg.namespaceURI !== SVG || svg.localName !== 'svg' || !svg.isConnected || !svg.ownerDocument.defaultView) throw new TypeError('A connected SVG slide is required.');
  if (!options || typeof options.slideId !== 'string' || !options.slideId || typeof options.elementId !== 'string' || !options.elementId) throw new TypeError('slideId and elementId are required.');
  if (Array.from(svg.children).some(child => child.getAttribute('data-pptkit-media-overlay') === options.elementId)) throw new TypeError('A media overlay for this object already exists.');
  const { target, parents } = findLayoutMedia(layout, options.slideId, options.elementId);
  if (parents.length > 256) throw new RangeError('Media overlay exceeds the group depth limit.');
  if (!layout.size || !Number.isFinite(layout.size.width) || !Number.isFinite(layout.size.height) || layout.size.width <= 0 || layout.size.height <= 0 || svg.getAttribute('data-pptkit-slide-id') !== options.slideId || !svg.hasAttribute('viewBox')) throw new TypeError('SVG slide identity or size does not match the Layout result.');
  const viewBox = svg.getAttribute('viewBox')!.trim().split(/[\s,]+/u).map(Number);
  if (viewBox.length !== 4 || viewBox[0] !== 0 || viewBox[1] !== 0 || viewBox[2] !== layout.size.width || viewBox[3] !== layout.size.height) throw new TypeError('SVG viewBox does not match the Layout size.');
  validateGeometry(target.box, target.transform, target.opacity);
  const groupValues = parents.map(group => ({ transform: groupTransform(group), opacity: group.opacity }));
  validateComposite(parents, target.box, target.transform);
  const targetTransform = shapeTransform(target.box, target.transform), box = { ...target.box }, opacity = target.opacity;
  const doc = svg.ownerDocument, root = doc.createElementNS(SVG, 'g'); root.setAttribute('data-pptkit-media-overlay', options.elementId);
  let parent: SVGElement = root;
  for (const group of groupValues) { const node = doc.createElementNS(SVG, 'g'); node.setAttribute('transform', group.transform); node.setAttribute('opacity', String(group.opacity)); parent.appendChild(node); parent = node; }
  const item = doc.createElementNS(SVG, 'g'); item.setAttribute('transform', targetTransform); item.setAttribute('opacity', String(opacity)); parent.appendChild(item);
  const foreign = doc.createElementNS(SVG, 'foreignObject');
  for (const key of ['x', 'y', 'width', 'height'] as const) foreign.setAttribute(key, String(box[key]));
  const host = doc.createElementNS(XHTML, 'div') as HTMLDivElement; host.style.width = '100%'; host.style.height = '100%'; foreign.appendChild(host); item.appendChild(foreign);
  let mounted: BrowserLayoutMedia | undefined;
  try {
    svg.appendChild(root);
    mounted = await mountLayoutMedia(host, layout, options);
    if (mounted.controller.state.status === 'disposed') { root.remove(); return mounted; }
    const media = mounted.controller.element; media.style.width = '100%'; media.style.height = '100%'; media.style.display = 'block';
    if (mounted.kind === 'video') media.style.objectFit = 'fill';
    mounted.controller.subscribe(state => {
      if (state.status === 'disposed') root.remove();
      else media.style.visibility = state.status === 'loading' || state.status === 'error' ? 'hidden' : 'visible';
    });
    return mounted;
  } catch (error) { mounted?.controller.dispose(); root.remove(); throw error; }
}
