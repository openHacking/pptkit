import { normalizePresentation } from '../normalization/normalize-presentation.js';
import type { NormalizedElement, NormalizedTextParagraph } from '../types/element.js';
import type { PresentationData } from '../types/presentation.js';
import type { ColorValue, NormalizedPaint, NormalizedStrokeStyle, NormalizedTextRunStyle } from '../types/style.js';
import type { NormalizedPresentationTheme } from '../types/theme.js';

export type TextContrastFinding = Readonly<{
  slideId: string; elementId: string; path: string; paragraphIndex: number; runIndex: number;
  /** Present for a visible occurrence of an inherited text definition. */
  source?: Readonly<{ kind: 'master' | 'layout'; id: string }>;
} & ({ status: 'indeterminate'; reason: 'complex-scene' | 'non-solid-background' | 'unsupported-text-owner' | 'translucent-object' | 'dynamic-timeline' }
  | { status: 'indeterminate'; reason: 'output-size-dependent' | 'output-opacity-dependent'; ratio: number; nativeRatio: number; threshold: 3 | 4.5; nativeThreshold: 3 | 4.5; nominalSvgThreshold: 3 | 4.5; effectiveFontSize: number }
  | { status: 'meets-threshold' | 'below-threshold'; ratio: number; nativeRatio: number; threshold: 3 | 4.5; nativeThreshold: 3 | 4.5; nominalSvgThreshold: 3 | 4.5; effectiveFontSize: number })>;
export interface TextContrastInspection { readonly scope: 'slide-text-runs'; readonly findings: readonly TextContrastFinding[] }
export class TextContrastLimitError extends Error {
  constructor(readonly maxFindings: number) {
    super('Text contrast finding limit exceeded.');
    this.name = 'TextContrastLimitError';
  }
}

function rgb(color: ColorValue, theme: NormalizedPresentationTheme): [number, number, number] {
  const hex = (typeof color === 'string' ? color : theme.colors[color.theme]).replace(/^#/, '');
  return [0, 2, 4].map(index => parseInt(hex.slice(index, index + 2), 16) / 255) as [number, number, number];
}
function luminance(channels: readonly number[]): number {
  const linear = channels.map(c => c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  return 0.2126 * linear[0]! + 0.7152 * linear[1]! + 0.0722 * linear[2]!;
}
function coversSlideAxis(offset: number, length: number, slideLength: number): boolean {
  // PPTX rounds the shape's origin, extent and slide extent independently to EMU.
  // A mathematical cover can otherwise become a one-EMU gap after export.
  return offset <= 0 && offset + length >= slideLength
    && Math.round(offset * 12700) <= 0
    && Math.round(offset * 12700) + Math.round(length * 12700) >= Math.round(slideLength * 12700);
}
type Rgb = [number, number, number];
type ColorField = { svg: Rgb; native: Rgb };
function composite(foreground: Rgb, background: Rgb, alpha: number): Rgb {
  return foreground.map((channel, index) => alpha * channel + (1 - alpha) * background[index]!) as Rgb;
}
function invisibleStroke(stroke: NormalizedStrokeStyle): boolean {
  return stroke.width === 0 || stroke.paint.type === 'none';
}

type SolidLayer = { color: Rgb; opacity: number };
function solidSlideCover(element: Extract<NormalizedElement, { type: 'shape' }>, size: PresentationData['size'], theme: NormalizedPresentationTheme): SolidLayer | undefined {
  const { box, style, transform } = element, fill = style.fill;
  if (element.shape !== 'rect' || transform.rotation !== 0 || style.shadow || !invisibleStroke(style.stroke) || fill.type !== 'solid' ||
      !coversSlideAxis(box.x, box.width, size.width) || !coversSlideAxis(box.y, box.height, size.height)) return undefined;
  return { color: rgb(fill.color, theme), opacity: fill.opacity };
}

function composeLayer(field: ColorField | undefined, { color, opacity }: SolidLayer): ColorField | undefined {
  if (opacity === 1) return { svg: color, native: color };
  return field && {
    svg: composite(color, field.svg, opacity),
    native: composite(color, field.native, Math.round(opacity * 100000) / 100000),
  };
}

/** Only text-free rectangles with uniform paint across the entire slide qualify. */
function backdropLayers(elements: readonly NormalizedElement[], size: PresentationData['size'], theme: NormalizedPresentationTheme): SolidLayer[] | undefined {
  const layers: SolidLayer[] = [];
  for (const element of elements) {
    if (element.hidden) continue;
    if (element.type !== 'shape' || element.text !== undefined || element.opacity !== 1) return undefined;
    const layer = solidSlideCover(element, size, theme);
    if (!layer) return undefined;
    layers.push(layer);
  }
  return layers;
}

/** A group can pass through contrast only when it does not scale, tint or overlap its child. */
function unscaledForeground(element: NormalizedElement | undefined): NormalizedElement | undefined {
  while (element?.type === 'group') {
    if (element.opacity !== 1 || element.accessibility.decorative || element.transform.rotation !== 0 ||
        element.transform.flipH || element.transform.flipV || element.box.width !== element.coordinateSize.width ||
        element.box.height !== element.coordinateSize.height ||
        Math.round(element.box.width * 12700) === 0 || Math.round(element.box.height * 12700) === 0) return undefined;
    const visible = element.children.filter(child => !child.hidden);
    if (visible.length !== 1) return undefined;
    element = visible[0];
  }
  return element;
}

function ownerScene(elements: readonly NormalizedElement[], path: string, size: PresentationData['size'], theme: NormalizedPresentationTheme, source?: TextContrastFinding['source']) {
  const visible = elements.filter(element => !element.hidden);
  const target = visible.at(-1);
  return {
    elements, path, theme, source, target, foregroundElement: unscaledForeground(target),
    backdrop: backdropLayers(visible.slice(0, -1), size, theme),
    complete: target ? backdropLayers([target], size, theme) : [],
  };
}

function compare(style: NormalizedTextRunStyle, field: ColorField, theme: NormalizedPresentationTheme, fontScale: number) {
  const foreground = rgb(style.color, theme), alpha = style.colorOpacity ?? 1;
  const ratioAt = (opacity: number, background: Rgb) => {
    const b = luminance(background), a = luminance(composite(foreground, background, opacity));
    return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
  };
  const ratio = ratioAt(alpha, field.svg), nativeRatio = ratioAt(Math.round(alpha * 100000) / 100000, field.native);
  const effectiveFontSize = style.fontSize * fontScale;
  // Export writes hundredths of a point and a 1/100000 shrink scale; SVG
  // serializes its nominal CSS pixel size to three decimal places.
  const nativeSize = Math.round(style.fontSize * 100) / 100 * (Math.round(fontScale * 100000) / 100000);
  const svgSize = Math.round(effectiveFontSize * 1000) / 1000;
  const nativeThreshold = nativeSize >= 18 || style.bold && nativeSize >= 14 ? 3 as const : 4.5 as const;
  const nominalSvgThreshold = svgSize >= 24 || style.bold && svgSize >= 56 / 3 ? 3 as const : 4.5 as const;
  const threshold = nominalSvgThreshold;
  const svgPass = ratio >= nominalSvgThreshold, nativePass = nativeRatio >= nativeThreshold;
  if (svgPass !== nativePass) return { status: 'indeterminate' as const, reason: nativeThreshold !== nominalSvgThreshold ? 'output-size-dependent' as const : 'output-opacity-dependent' as const, ratio, nativeRatio, threshold, nativeThreshold, nominalSvgThreshold, effectiveFontSize };
  return { status: svgPass ? 'meets-threshold' as const : 'below-threshold' as const, ratio, nativeRatio, threshold, nativeThreshold, nominalSvgThreshold, effectiveFontSize };
}

/** A conservative color-pair review, not a rendered-pixel or WCAG conformance assessment. */
export function inspectSlideTextContrast(document: PresentationData, options: { maxFindings?: number } = {}): TextContrastInspection {
  if (!options || typeof options !== 'object' || Array.isArray(options) || options.maxFindings !== undefined && (!Number.isSafeInteger(options.maxFindings) || options.maxFindings < 0)) throw new TypeError('maxFindings must be a nonnegative safe integer.');
  const { maxFindings } = options;
  const normalized = normalizePresentation(document), findings: TextContrastFinding[] = [];
  const layouts = new Map(normalized.layouts.map((layout, index) => [layout.id, {
    masterId: layout.masterId, showMasterShapes: layout.showMasterShapes,
    scene: ownerScene(layout.elements, `layouts.${index}.elements`, normalized.size, layout.theme ?? normalized.theme, Object.freeze({ kind: 'layout', id: layout.id })),
  }]));
  const masters = new Map((normalized.masters ?? []).map((master, index) => [master.id,
    ownerScene(master.elements ?? [], `masters.${index}.elements`, normalized.size, master.theme, Object.freeze({ kind: 'master', id: master.id })),
  ]));
  for (const [slideIndex, slide] of normalized.slides.entries()) {
    if (slide.hidden) continue;
    const layout = layouts.get(slide.layoutId), slideTheme = layout?.scene.theme ?? normalized.theme;
    const owners: ReturnType<typeof ownerScene>[] = [];
    if (slide.showInheritedShapes !== false && layout) {
      const master = layout.masterId === undefined ? undefined : masters.get(layout.masterId);
      if (master && layout.showMasterShapes !== false) owners.push(master);
      owners.push(layout.scene);
    }
    owners.push(ownerScene(slide.elements, `slides.${slideIndex}.elements`, normalized.size, slideTheme));
    let foreground: ReturnType<typeof ownerScene> | undefined;
    for (const owner of owners) if (owner.target !== undefined) foreground = owner;
    const uniformScene = owners.every(owner => owner.backdrop !== undefined && (owner === foreground || owner.complete !== undefined));
    const background = slide.background.type === 'solid' && slide.background.opacity === 1 ? rgb(slide.background.color, slideTheme) : undefined;
    let backgroundField: ColorField | undefined = background ? { svg: background, native: background } : undefined;
    if (uniformScene) for (const owner of owners) {
      for (const layer of owner.backdrop!) backgroundField = composeLayer(backgroundField, layer);
      if (owner !== foreground) for (const layer of owner.complete!) backgroundField = composeLayer(backgroundField, layer);
    }
    for (const owner of owners) {
      const { theme } = owner;
      const sameField = (fill: NormalizedPaint, allowTransparent = true): boolean => {
        if (!backgroundField) return false;
        if (fill.type === 'none' || allowTransparent && fill.opacity === 0) return true;
        return rgb(fill.color, theme).every((channel, index) => channel === backgroundField!.svg[index] && channel === backgroundField!.native[index]);
      };
      // Shape text can overflow its frame, so the field must be uniform across the
      // slide. A full-slide rectangle supplies a uniform field; translucent fill
      // also requires a known background and distinct SVG/PPTX alpha precision.
      const shapeField = (element: Extract<NormalizedElement, { type: 'shape' }>, grouped: boolean): ColorField | undefined => {
        if (element.style.shadow || !invisibleStroke(element.style.stroke)) return undefined;
        const fill = element.style.fill;
        // Child coordinates are local to the group, so they cannot prove slide coverage.
        const layer = grouped ? undefined : solidSlideCover(element, normalized.size, theme);
        if (layer) return composeLayer(backgroundField, layer);
        return (fill.type === 'none' || fill.type === 'solid') && sameField(fill) ? backgroundField : undefined;
      };
      // A uniform field avoids assuming that cell geometry clips overflowing text
      // identically in every renderer, including merged-cell continuations.
      const uniformTable = (element: Extract<NormalizedElement, { type: 'table' }>): boolean => Boolean(backgroundField) && element.rows.every(row => row.cells.every(cell => sameField(cell.style.fill, false) && invisibleStroke(cell.style.stroke) && Object.values(cell.style.borders ?? {}).every(invisibleStroke)));
      const visit = (element: NormalizedElement, path: string, grouped: boolean): void => {
        if (element.hidden || element.accessibility.decorative) return;
        if (element.type === 'group') { element.children.forEach((child, index) => visit(child, `${path}.children.${index}`, true)); return; }
        const emit = (paragraphs: readonly NormalizedTextParagraph[], contentPath: string, reason: 'unsupported-text-owner' | undefined, scale = 1, field = backgroundField) => {
          paragraphs.forEach((paragraph, paragraphIndex) => paragraph.runs.forEach((run, runIndex) => {
            if (!run.text.trim()) return;
            if (maxFindings !== undefined && findings.length >= maxFindings) throw new TextContrastLimitError(maxFindings);
            const location = { slideId: slide.id, elementId: element.id, path: `${contentPath}.${paragraphIndex}.runs.${runIndex}`, paragraphIndex, runIndex, ...(owner.source ? { source: owner.source } : {}) };
            const why = reason ?? (slide.eventTimeline ? 'dynamic-timeline' : !uniformScene || owner !== foreground || element !== owner.foregroundElement ? 'complex-scene' : element.opacity !== 1 ? 'translucent-object' : !field ? 'non-solid-background' : undefined);
            findings.push(Object.freeze(why ? { ...location, status: 'indeterminate', reason: why } : { ...location, ...compare(run.style, field!, theme, scale) }));
          }));
        };
        if (element.type === 'text') emit(element.content, `${path}.content`, undefined, element.frame.autoFit.mode === 'shrink' ? element.frame.autoFit.fontScale : 1);
        else if (element.type === 'shape' && element.text) { const field = shapeField(element, grouped); emit(element.text.content, `${path}.text.content`, field ? undefined : 'unsupported-text-owner', element.text.frame.autoFit.mode === 'shrink' ? element.text.frame.autoFit.fontScale : 1, field); }
        else if (element.type === 'table') { const reason = uniformTable(element) ? undefined : 'unsupported-text-owner'; element.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => emit(cell.content, `${path}.rows.${rowIndex}.cells.${cellIndex}.content`, reason))); }
      };
      owner.elements.forEach((element, index) => visit(element, `${owner.path}.${index}`, false));
    }
  }
  return Object.freeze({ scope: 'slide-text-runs', findings: Object.freeze(findings) });
}
