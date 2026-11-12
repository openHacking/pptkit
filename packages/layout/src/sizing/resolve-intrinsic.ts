import { ownerTheme } from "../analysis/owner-theme.js";
import { connectorGeometry } from "../geometry/connector.js";
import type { NormalizedElement, NormalizedPresentation, NormalizedTextFrameStyle, NormalizedTextParagraph, Size } from "@pptkit/core";
import type { LayoutDiagnostic, TextMeasurementOptions } from "../types/analysis.js";
import type { IntrinsicDimension, IntrinsicSizingRule } from "../types/intrinsic.js";
import { elementTransform, transformedBounds } from "../analysis/geometry.js";
import { measureTextBody, type MeasurementContext } from "../analysis/measure-text.js";
import { resolveShapeTextBox } from "../shape/resolve-shape-text-box.js";

type Axis = "width" | "height";
interface Item { element: NormalizedElement; parent?: Item; size: Size; context: MeasurementContext; hidden: boolean }
class SizingFailure extends Error {
  constructor(message: string, readonly retryWider = false) { super(message); }
}

export function validateIntrinsic(rules: readonly IntrinsicSizingRule[] | undefined): void {
  if (rules === undefined) return;
  if (!Array.isArray(rules)) throw new TypeError("intrinsicSizing must be an array");
  const ids = new Set<string>();
  for (const rule of rules as readonly IntrinsicSizingRule[]) {
    if (!rule || typeof rule.elementId !== "string" || ids.has(rule.elementId)) throw new TypeError("Intrinsic sizing requires unique element IDs");
    ids.add(rule.elementId);
    if (rule.width === undefined && rule.height === undefined) throw new TypeError("Intrinsic sizing requires an automatic axis");
    if (rule.includeHidden !== undefined && typeof rule.includeHidden !== "boolean") throw new TypeError("Intrinsic includeHidden must be boolean");
    for (const dimension of [rule.width, rule.height]) {
      if (dimension === undefined) continue;
      if (!dimension || (dimension.type !== "content" && dimension.type !== "parent")) throw new TypeError("Intrinsic dimension must use content or parent");
      for (const value of [dimension.min, dimension.max]) if (value !== undefined && (!Number.isFinite(value) || value < 0)) throw new RangeError("Intrinsic bounds must be finite and nonnegative");
      if ((dimension.min ?? 0) > (dimension.max ?? Infinity)) throw new RangeError("Intrinsic minimum exceeds maximum");
      if (dimension.type === "parent" && dimension.fraction !== undefined && (!Number.isFinite(dimension.fraction) || dimension.fraction < 0)) throw new RangeError("Parent fraction must be finite and nonnegative");
    }
  }
}

/** The whole intrinsic stage is atomic: a failed dependency retains its input. */
export function resolveIntrinsic(source: NormalizedPresentation, rules: readonly IntrinsicSizingRule[], provider: TextMeasurementOptions | undefined): { document: NormalizedPresentation; diagnostics: LayoutDiagnostic[] } {
  const document = structuredClone(source);
  const themeFor = ownerTheme(document);
  const diagnostics: LayoutDiagnostic[] = [];
  const items = new Map<string, Item>();
  const activeRules = new Map<string, IntrinsicSizingRule>();
  const states = new Map<string, "visiting" | "done">();
  const finalized = new Set<string>();
  const groupOrigins = new Map<string, { x: number; y: number }>();
  const original = new Map<string, { box: NormalizedElement["box"]; coordinateSize?: Size }>();
  const assets = new Map(document.assets.map((asset) => [asset.id, asset]));
  const chain: string[] = [];
  function walk(elements: NormalizedElement[], path: string, owner: { slideId?: string; layoutId?: string; masterId?: string }, hidden: boolean, parent?: Item): void {
    elements.forEach((element, index) => {
      const item: Item = { element, size: parent?.element.type === "group" ? parent.element.coordinateSize : document.size,
        context: { ...owner, elementId: element.id, path: `${path}.${index}` }, hidden: hidden || element.hidden, ...(parent ? { parent } : {}) };
      items.set(element.id, item);
      original.set(element.id, { box: { ...element.box }, ...(element.type === "group" ? { coordinateSize: { ...element.coordinateSize } } : {}) });
      if (element.type === "group") walk(element.children, `${item.context.path}.children`, owner, item.hidden, item);
    });
  }
  (document.masters ?? []).forEach((master, i) => walk(master.elements ?? [], `masters.${i}.elements`, { masterId: master.id }, false));
  document.layouts.forEach((layout, i) => walk(layout.elements, `layouts.${i}.elements`, { layoutId: layout.id }, false));
  document.slides.forEach((slide, i) => walk(slide.elements, `slides.${i}.elements`, { slideId: slide.id }, slide.hidden));
  function fail(item: Item | undefined, code: string, message: string, path?: string, retryWider = false): never {
    diagnostics.push({ ...(item?.context ?? { path: path ?? "options.intrinsicSizing" }), severity: "warning", code, message });
    throw new SizingFailure(message, retryWider);
  }
  function bounded(value: number, dimension: IntrinsicDimension): number {
    return Math.max(dimension.min ?? 0, Math.min(dimension.max ?? Infinity, value));
  }
  function finite(value: number, item: Item): number {
    if (!Number.isFinite(value) || value < 0) fail(item, "intrinsic-size-invalid", "Intrinsic sizing produced a nonfinite or negative size. No intrinsic rules were applied.");
    return value;
  }
  function imageSize(item: Item): Size {
    const element = item.element;
    if (element.type !== "image") throw new Error("Expected image");
    const asset = assets.get(element.assetId);
    if (asset?.width === undefined || asset.height === undefined || asset.width <= 0 || asset.height <= 0 || !Number.isFinite(asset.width) || !Number.isFinite(asset.height)) fail(item, "intrinsic-image-size-unavailable", "Image metadata must provide positive intrinsic dimensions; Layout does not load assets.");
    return { width: asset.width, height: asset.height };
  }
  function textSize(item: Item, axis: Axis, spec: IntrinsicDimension): number {
    const element = item.element;
    const text = element.type === "text" ? element : element.type === "shape" ? element.text : undefined;
    if (!text) return fail(item, "intrinsic-content-unavailable", "This element has no intrinsic text content.");
    if (!provider) return fail(item, "text-measurement-unavailable", "Intrinsic text sizing requires a synchronous text measurement provider.");
    const direction = text.frame.textDirection;
    const vertical = direction !== undefined && direction !== "horizontal" && direction !== "rotate90" && direction !== "rotate270";
    const rightToLeft = direction === "eastAsianVertical" || direction === "stackedVerticalRtl";
    if (vertical || text.frame.textDirection === "rotate90" || text.frame.textDirection === "rotate270") {
      if (element.type === "shape" && element.shape === "parallelogram") return fail(item, "text-direction-sizing-unsupported", "Intrinsic sizing of a rotated parallelogram text rectangle is not supported.");
      const { left, right, top, bottom } = text.frame.margin;
      const inline = axis === "height";
      const availableWidth = inline ? 0 : Math.max(0, dimension(item, "height") - top - bottom);
      const frame: NormalizedTextFrameStyle = { ...text.frame, textDirection: vertical ? direction : "horizontal", verticalAlign: "top", wrap: !inline && text.frame.wrap };
      const content = text.content.map(paragraph => ({ ...paragraph, style: { ...paragraph.style, align: "left" as const } }));
      const physical = measureTextBody(content, frame, { x: 0, y: 0, width: vertical ? 0 : availableWidth, height: vertical ? availableWidth : 0 }, themeFor(item.context), provider,
        { ...item.context, path: `${item.context.path}.${element.type === "shape" ? "text." : ""}content` }, diagnostics);
      if (!physical) throw new SizingFailure("Text measurement failed");
      // Vertical analysis returns physical coordinates. Recover logical extents.
      const measured = vertical ? { x: physical.y, y: rightToLeft ? -physical.x - physical.width : physical.x, width: physical.height, height: physical.width } : physical;
      if (measured.x < -1e-9 || measured.y < -1e-9)
        return fail(item, "intrinsic-size-unsatisfiable", "Measured vertical text extends outside the leading edge of the frame.");
      if (!inline && measured.x + measured.width > availableWidth + 1e-9)
        return fail(item, "intrinsic-size-unsatisfiable", "Measured vertical text exceeds the available height.");
      const needed = inline ? Math.max(0, measured.x + measured.width) + top + bottom : Math.max(0, measured.y + measured.height) + left + right;
      if (!inline && needed > (spec.max ?? Infinity))
        return fail(item, "intrinsic-size-unsatisfiable", "Measured vertical text exceeds maximum width.");
      return needed;
    }
    if (element.type === "shape" && element.shape === "parallelogram" && axis === "height") {
      if (activeRules.get(element.id)?.width?.type === "content" && states.get(`${element.id}.width`) !== "done") return fail(item, "intrinsic-shape-height-unavailable", "Parallelogram content width must be resolved before its height. No intrinsic rules were applied.");
      const width = dimension(item, "width"), frame = { ...text.frame, verticalAlign: "top" as const };
      const context = { ...item.context, path: `${item.context.path}.text.content` };
      const measureAt = (availableWidth: number) => measureTextBody(text.content, frame, { x: 0, y: 0, width: availableWidth, height: 0 }, themeFor(item.context), provider, context, diagnostics);
      const initialWidth = frame.wrap ? Math.max(0, (5 / 6) * width - frame.margin.left - frame.margin.right) : 0;
      const initial = measureAt(initialWidth);
      if (!initial) throw new SizingFailure("Text measurement failed");
      let needed = finite(Math.max(0, initial.y + initial.height) + frame.margin.top + frame.margin.bottom, item);
      const slant = element.slant ?? 0.25, a = 5 / 6;
      const unsatisfied = (message: string, retryWider = false): never => fail(item, frame.wrap ? "intrinsic-shape-height-unavailable" : "intrinsic-size-unsatisfiable", message, undefined, frame.wrap && retryWider);
      for (let attempt = 0; attempt < (frame.wrap ? 32 : 1); attempt++) {
        let height = needed;
        if (width > 0 && needed > 0) {
          // Below width, the preset height is a*h - a*slant*h²/width.
          // Use the stable quadratic root, then the linear region above width.
          const ratio = needed / width, discriminant = a * a - 4 * a * slant * ratio;
          const root = discriminant >= 0 ? width * (2 * ratio / (a + Math.sqrt(discriminant))) : Infinity;
          if (root <= width) height = root;
          else if (slant < 1) height = needed / (a * (1 - slant));
          else return unsatisfied("The parallelogram text rectangle cannot contain the measured height at the fixed width and slant.", true);
        }
        height = Math.max(height, spec.min ?? 0);
        if (!Number.isFinite(height)) return unsatisfied("The required parallelogram height exceeds finite coordinates.");
        // For high slants the text rectangle initially grows, then shrinks as
        // height approaches width. A minimum can land in that infeasible gap;
        // the second, linear region may still satisfy the constraint.
        const candidateBox = resolveShapeTextBox({ ...element, box: { x: 0, y: 0, width, height } });
        if (candidateBox.height + Math.max(1e-9, candidateBox.height * 1e-12) < needed) {
          if (slant === 1 || width === 0) return unsatisfied("The minimum height leaves insufficient parallelogram text space.", true);
          height = Math.max(height, width, needed / (a * (1 - slant)));
        }
        if (!Number.isFinite(height) || height > (spec.max ?? Infinity)) return unsatisfied("The measured parallelogram height exceeds its maximum.", true);
        const box = resolveShapeTextBox({ ...element, box: { x: 0, y: 0, width, height } });
        const actual = measureAt(Math.max(0, box.width - frame.margin.left - frame.margin.right));
        if (!actual) throw new SizingFailure("Text measurement failed");
        const actualNeeded = Math.max(0, actual.y + actual.height) + frame.margin.top + frame.margin.bottom;
        if (!Number.isFinite(actualNeeded) || actual.y < -1e-9) return fail(item, "intrinsic-shape-height-unavailable", "The remeasured vertical bounds are invalid or extend above the content area. No intrinsic rules were applied.");
        if (actualNeeded <= box.height + Math.max(1e-9, box.height * 1e-12)) return height;
        if (!frame.wrap || actualNeeded <= needed) return fail(item, "intrinsic-shape-height-unavailable", "The resulting shape frame does not contain the remeasured vertical bounds. No intrinsic rules were applied.");
        needed = actualNeeded;
      }
      return fail(item, "intrinsic-shape-height-unavailable", "Wrapped parallelogram height did not converge within 32 iterations. No intrinsic rules were applied.");
    }
    const frame: NormalizedTextFrameStyle = { ...text.frame, verticalAlign: "top", wrap: axis === "width" ? false : text.frame.wrap };
    const content: readonly NormalizedTextParagraph[] = axis === "width" ? text.content.map((paragraph) => ({ ...paragraph, style: { ...paragraph.style, align: "left" } })) : text.content;
    const width = axis === "width" ? 0 : Math.max(0, dimension(item, "width") - frame.margin.left - frame.margin.right);
    const measured = measureTextBody(content, frame, { x: 0, y: 0, width, height: 0 }, themeFor(item.context), provider, { ...item.context, path: `${item.context.path}.${element.type === "shape" ? "text." : ""}content` }, diagnostics);
    if (!measured) throw new SizingFailure("Text measurement failed");
    const needed = axis === "width" ? Math.max(0, measured.x + measured.width) + frame.margin.left + frame.margin.right : Math.max(0, measured.y + measured.height) + frame.margin.top + frame.margin.bottom;
    if (axis === "width" && element.type === "shape" && element.shape === "parallelogram") {
      const height = dimension(item, "height"), slant = element.slant ?? 0.25;
      if (needed === 0) return 0;
      // Invert the preset rectangle width in its two regions: width <= height
      // and width >= height. At slant 1 the first region has no usable width.
      const threshold = (5 / 6) * (1 - slant) * height;
      return slant < 1 && needed <= threshold ? needed / ((5 / 6) * (1 - slant)) : needed / (5 / 6) + slant * height;
    }
    return needed;
  }
  function parallelogramContentSize(item: Item, widthRule: IntrinsicDimension, heightRule: IntrinsicDimension, fixed?: { axis: Axis; value: number }): Size {
    const element = item.element;
    if (element.type !== "shape" || element.shape !== "parallelogram" || !element.text) return fail(item, "intrinsic-content-unavailable", "Parallelogram content sizing requires shape text.");
    if (!provider) return fail(item, "text-measurement-unavailable", "Intrinsic text sizing requires a synchronous text measurement provider.");
    const frame = { ...element.text.frame, verticalAlign: "top" as const };
    const content = element.text.content.map((paragraph) => ({ ...paragraph, style: { ...paragraph.style, align: "left" as const } }));
    const context = { ...item.context, path: `${item.context.path}.text.content` };
    const direction = frame.textDirection;
    const quarter = direction === "rotate90" || direction === "rotate270";
    const vertical = direction !== undefined && direction !== "horizontal";
    const rightToLeft = direction === "eastAsianVertical" || direction === "stackedVerticalRtl";
    const measureAt = (width: number, height: number, wrap: boolean, authoredAlignment = false) => {
      const style = { ...frame, wrap };
      const measured = measureTextBody(authoredAlignment ? element.text!.content : content,
        quarter ? { ...style, textDirection: "horizontal" } : style,
        { x: 0, y: 0, width: quarter ? height : width, height: quarter ? width : height },
        themeFor(item.context), provider, context, diagnostics);
      if (!measured || !vertical) return measured;
      if (quarter) return { x: measured.y, y: measured.x, width: measured.height, height: measured.width };
      return rightToLeft ? { ...measured, x: width - measured.x - measured.width } : measured;
    };
    const marginWidth = frame.margin.left + frame.margin.right;
    const marginHeight = frame.margin.top + frame.margin.bottom;
    const measured = measureAt(0, 0, false);
    if (!measured) throw new SizingFailure("Text measurement failed");
    const naturalWidth = finite(Math.max(0, measured.x + measured.width) + marginWidth, item);
    const naturalHeight = finite(Math.max(0, measured.y + measured.height) + marginHeight, item);
    const slant = element.slant ?? .25, a = 5 / 6;
    const minWidth = widthRule.min ?? 0, maxWidth = widthRule.max ?? Infinity;
    const minHeight = heightRule.min ?? 0, maxHeight = heightRule.max ?? Infinity;
    const inset = (width: number, height: number) => resolveShapeTextBox({ ...element, box: { x: 0, y: 0, width, height } });
    const fits = (needed: number, available: number) =>
      Number.isFinite(needed) && needed <= available + Math.max(1e-9, available * 1e-12);
    function heightAtWidth(width: number, requiredHeight: number): number {
      if (width === 0) return Math.max(minHeight, requiredHeight);
      // Below the fixed width the inset height is quadratic; above it, linear.
      const ratio = requiredHeight / width;
      const discriminant = a * a - 4 * a * slant * ratio;
      const root = discriminant >= 0 ? width * (2 * ratio / (a + Math.sqrt(discriminant))) : Infinity;
      const linear = slant < 1 ? Math.max(width, requiredHeight / (a * (1 - slant))) : Infinity;
      const height = Math.max(minHeight, root <= width ? root : linear);
      // A minimum can fall into the gap between the two fitting ranges.
      return Number.isFinite(height) && !fits(requiredHeight, inset(width, height).height) ? Math.max(minHeight, linear) : height;
    }
    function candidate(requiredWidth: number, requiredHeight: number): Size | undefined {
      let width: number, height: number;
      if (fixed?.axis === "width") {
        width = fixed.value;
        height = heightAtWidth(width, requiredHeight);
      } else {
        // Both axes are content sized: select a deterministic aspect ratio
        // that stays away from the degenerate square at slant 1.
        const ratio = Math.min(1, requiredWidth > 0 ? requiredHeight / requiredWidth : 1, slant > .5 ? .5 / slant : 1);
        height = fixed?.value ?? bounded(requiredHeight / (a * (1 - slant * ratio)), heightRule);
        const threshold = a * (1 - slant) * height;
        const horizontal = slant < 1 && requiredWidth <= threshold ? requiredWidth / (a * (1 - slant)) : requiredWidth / a + slant * height;
        width = Math.max(minWidth, horizontal);
        if (requiredHeight > threshold) {
          const denominator = 1 - requiredHeight / (a * height);
          if (height === 0 || denominator < 0 || (denominator === 0 && slant > 0)) width = Infinity;
          else if (slant > 0) width = Math.max(width, slant * height / denominator);
        }
        // The preferred aspect ratio may need too much width. Try the
        // narrowest fitting height at the allowed width before rejecting.
        if (frame.wrap && !fixed && width > maxWidth && Number.isFinite(maxWidth)) {
          width = maxWidth;
          height = heightAtWidth(width, requiredHeight);
        }
      }
      if (!Number.isFinite(width) || !Number.isFinite(height) || width > maxWidth || height > maxHeight) return undefined;
      const box = inset(width, height);
      return fits(requiredWidth, box.width) && fits(requiredHeight, box.height) ? { width, height } : undefined;
    }
    function verified(size: Size): boolean {
      const box = inset(size.width, size.height);
      const actual = measureAt(Math.max(0, box.width - marginWidth), Math.max(0, box.height - marginHeight), frame.wrap, true);
      if (!actual) throw new SizingFailure("Text measurement failed");
      return actual.x >= -1e-9 && actual.y >= -1e-9 &&
        fits(Math.max(0, actual.x + actual.width) + marginWidth, box.width) &&
        fits(Math.max(0, actual.y + actual.height) + marginHeight, box.height);
    }
    const natural = candidate(naturalWidth, naturalHeight);
    if (natural && verified(natural)) return natural;
    if (!frame.wrap) {
      if (!natural) return fail(item, "intrinsic-size-unsatisfiable", "The unwrapped text cannot fit within the parallelogram size bounds.");
      return fail(item, "intrinsic-shape-height-unavailable", "The resulting parallelogram frame does not fit the remeasured content. No intrinsic rules were applied.");
    }

    // Probe actual inline limits, then invert the inset using the measured
    // column width. Wrapping changes at word boundaries, so an unsuccessful
    // probe does not establish that all smaller limits are infeasible.
    let limit = Math.min(naturalHeight, (fixed?.axis === "width" && fixed.value === 0 ? 1 : a) * maxHeight);
    if (Number.isFinite(maxWidth) && (Number.isFinite(maxHeight) || slant === 1 && maxWidth > 0)) {
      const peak = slant > .5 ? Math.max(minHeight, Math.min(maxHeight, maxWidth / (2 * slant))) : maxHeight;
      limit = Math.min(limit, Math.max(inset(maxWidth, peak).height, inset(maxWidth, Number.isFinite(maxHeight) ? maxHeight : Math.max(minHeight, maxWidth)).height));
    }
    const span = limit - marginHeight;
    if (span > 0) {
      const authored = inset(element.box.width, element.box.height).height;
      const probes = new Set([limit, ...(authored > marginHeight && authored < limit ? [authored] : []),
        ...Array.from({ length: 24 }, (_, i) => marginHeight + span * (24 - i) / 25)]);
      for (const height of probes) {
        const wrapped = measureAt(0, height - marginHeight, true);
        if (!wrapped) throw new SizingFailure("Text measurement failed");
        const width = Math.max(0, wrapped.x + wrapped.width) + marginWidth;
        if (wrapped.x < -1e-9 || wrapped.y < -1e-9 || !fits(wrapped.y + wrapped.height + marginHeight, height)) continue;
        const size = candidate(width, height);
        if (size && verified(size)) return size;
      }
    }
    return fail(item, "intrinsic-shape-height-unavailable", "No wrapped parallelogram candidate fit the measured text within the size bounds. No intrinsic rules were applied.");
  }
  function wrappedParallelogramContentSize(item: Item, widthRule: IntrinsicDimension, heightRule: IntrinsicDimension): Size {
    const element = item.element;
    if (element.type !== "shape" || element.shape !== "parallelogram" || !element.text) return fail(item, "intrinsic-content-unavailable", "Parallelogram content sizing requires shape text.");
    if (!provider) return fail(item, "text-measurement-unavailable", "Intrinsic text sizing requires a synchronous text measurement provider.");
    const frame = { ...element.text.frame, verticalAlign: "top" as const, wrap: false };
    const content = element.text.content.map(paragraph => ({ ...paragraph, style: { ...paragraph.style, align: "left" as const } }));
    const context = { ...item.context, path: `${item.context.path}.text.content` };
    const natural = measureTextBody(content, frame, { x: 0, y: 0, width: 0, height: 0 }, themeFor(item.context), provider, context, diagnostics);
    if (!natural) throw new SizingFailure("Text measurement failed");
    const required = finite(Math.max(0, natural.x + natural.width) + frame.margin.left + frame.margin.right, item);
    // Pick an explicit candidate width; its reference height is the authored
    // height. Width bounds may cause wrapping, which the height pass verifies.
    let width = bounded(finite(required / (5 / 6) + (element.slant ?? .25) * element.box.height, item), widthRule);
    for (let attempt = 0; attempt < 16; attempt++) {
      element.box.width = width;
      states.set(`${element.id}.width`, "done");
      const diagnosticsStart = diagnostics.length;
      let height: number;
      try { height = textSize(item, "height", heightRule); }
      catch (error) {
        if (!(error instanceof SizingFailure) || !error.retryWider || attempt === 15) throw error;
        const next = Math.min(widthRule.max ?? Infinity, Math.max(width * 2, width + 1));
        if (!Number.isFinite(next) || next <= width) throw error;
        diagnostics.splice(diagnosticsStart);
        width = next;
        continue;
      }
      const box = resolveShapeTextBox({ ...element, box: { x: 0, y: 0, width, height } });
      const actual = measureTextBody(element.text.content, { ...element.text.frame, verticalAlign: "top" }, { x: 0, y: 0, width: Math.max(0, box.width - frame.margin.left - frame.margin.right), height: 0 }, themeFor(item.context), provider, context, diagnostics);
      if (!actual) throw new SizingFailure("Text measurement failed");
      const actualWidth = Math.max(0, actual.x + actual.width) + frame.margin.left + frame.margin.right;
      const actualHeight = Math.max(0, actual.y + actual.height) + frame.margin.top + frame.margin.bottom;
      if (!Number.isFinite(actualWidth) || !Number.isFinite(actualHeight) || actual.x < -1e-9 || actual.y < -1e-9 || actualHeight > box.height + Math.max(1e-9, box.height * 1e-12)) return fail(item, "intrinsic-shape-height-unavailable", "The selected wrapped parallelogram frame does not contain the remeasured content. No intrinsic rules were applied.");
      if (actualWidth <= box.width + Math.max(1e-9, box.width * 1e-12)) return { width, height };
      // Invert text width at the measured height and retry with a larger frame.
      const slant = element.slant ?? .25, threshold = (5 / 6) * (1 - slant) * height;
      const requiredFrame = slant < 1 && actualWidth <= threshold ? actualWidth / ((5 / 6) * (1 - slant)) : actualWidth / (5 / 6) + slant * height;
      const next = Math.min(widthRule.max ?? Infinity, Math.max(requiredFrame, width + Math.max(1e-7, Math.abs(width) * 1e-12)));
      if (!Number.isFinite(next) || next <= width) return fail(item, "intrinsic-shape-height-unavailable", "The wrapped content does not fit within the selected width bounds. No intrinsic rules were applied.");
      width = next;
    }
    return fail(item, "intrinsic-shape-height-unavailable", "Wrapped parallelogram width did not converge within 16 iterations. No intrinsic rules were applied.");
  }
  function groupExtent(item: Item, axis: Axis): number {
    const element = item.element;
    if (element.type !== "group") throw new Error("Expected group");
    const rule = activeRules.get(element.id)!;
    const children = element.children.filter((child) => rule.includeHidden || !child.hidden);
    if (!children.length) return fail(item, "intrinsic-content-unavailable", "Cannot derive group size from an empty visible child set.");
    let start = 0;
    let end = 0;
    for (const child of children) {
      if (child.type !== "connector") resolveElement(items.get(child.id)!);
      else for (const endpoint of [child.start, child.end]) {
        if (!("elementId" in endpoint)) continue;
        const target = element.children.find((candidate) => candidate.id === endpoint.elementId);
        if (!target || target.type === "connector") return fail(item, "intrinsic-connector-dependency", "Group fitting requires connector targets within the same group and cannot infer connector-to-connector dependencies.");
        resolveElement(items.get(target.id)!);
      }
    }
    const boxes = new Map(element.children.map((child) => [child.id, child.box]));
    const transforms = new Map(element.children.map((child) => [child.id, child.transform]));
    for (const child of children) {
      const frame = child.type === "connector" ? connectorGeometry(child, boxes, transforms).box : child.box;
      const bounds = transformedBounds(frame, elementTransform(frame, child.transform));
      const coordinate = axis === "width" ? bounds.x : bounds.y;
      start = Math.min(start, coordinate);
      end = Math.max(end, coordinate + bounds[axis]);
    }
    const origin = groupOrigins.get(element.id) ?? { x: 0, y: 0 };
    origin[axis === "width" ? "x" : "y"] = start;
    groupOrigins.set(element.id, origin);
    const before = original.get(element.id)!;
    const scale = before.box[axis] === 0 ? 1 : before.box[axis] / before.coordinateSize![axis];
    return finite((end - start) * scale, item);
  }
  function dimension(item: Item, axis: Axis): number {
    const element = item.element;
    const rule = activeRules.get(element.id);
    const spec = rule?.[axis];
    if (!spec) return element.box[axis];
    const key = `${element.id}.${axis}`;
    if (states.get(key) === "done") return element.box[axis];
    if (states.get(key) === "visiting") return fail(item, "intrinsic-sizing-cycle", `Cyclic size dependency: ${[...chain, key].join(" -> ")}. No intrinsic rules were applied.`);
    states.set(key, "visiting"); chain.push(key);
    const frame = element.type === "text" ? element.frame : element.type === "shape" ? element.text?.frame : undefined;
    const direction = frame?.textDirection;
    let value: number;
    if (spec.type === "parent") {
      let parentSize = item.size[axis];
      if (item.parent) {
        dimension(item.parent, axis);
        const parent = item.parent.element;
        if (parent.type === "group") {
          const parentSpec = activeRules.get(parent.id)?.[axis];
          const before = original.get(parent.id)!;
          parentSize = parentSpec?.type === "content" ? parent.box[axis] / (before.box[axis] === 0 ? 1 : before.box[axis] / before.coordinateSize![axis]) : parent.coordinateSize[axis];
        }
      }
      value = bounded(finite(parentSize * (spec.fraction ?? 1), item), spec);
    } else if (element.type === "image") {
      const natural = imageSize(item);
      const other: Axis = axis === "width" ? "height" : "width";
      if (rule?.width?.type === "content" && rule.height?.type === "content") {
        const low = Math.max((rule.width.min ?? 0) / natural.width, (rule.height.min ?? 0) / natural.height);
        const high = Math.min((rule.width.max ?? Infinity) / natural.width, (rule.height.max ?? Infinity) / natural.height);
        if (low > high) return fail(item, "intrinsic-size-unsatisfiable", "Image bounds cannot preserve its natural aspect ratio.");
        const scale = Math.max(low, Math.min(high, 1));
        element.box.width = finite(natural.width * scale, item);
        element.box.height = finite(natural.height * scale, item);
        states.set(`${element.id}.width`, "done"); states.set(`${element.id}.height`, "done");
        value = element.box[axis];
      } else {
        value = finite(dimension(item, other) * natural[axis] / natural[other], item);
        if (value < (spec.min ?? 0) || value > (spec.max ?? Infinity)) return fail(item, "intrinsic-size-unsatisfiable", "Image bounds conflict with the fixed opposite axis and natural aspect ratio.");
      }
    } else if (element.type === "shape" && element.shape === "parallelogram" && rule?.width?.type === "content" && rule.height?.type === "content") {
      const size = element.text?.frame.wrap && (direction === undefined || direction === "horizontal")
        ? wrappedParallelogramContentSize(item, rule.width, rule.height)
        : parallelogramContentSize(item, rule.width, rule.height);
      element.box.width = size.width; element.box.height = size.height;
      states.set(`${element.id}.width`, "done"); states.set(`${element.id}.height`, "done");
      value = size[axis];
    } else if (element.type === "shape" && element.shape === "parallelogram" && direction !== undefined && direction !== "horizontal") {
      const other = axis === "width" ? "height" : "width";
      const fixed = dimension(item, other);
      const fixedRule = { type: "content" as const, min: fixed, max: fixed };
      value = parallelogramContentSize(item, axis === "width" ? spec : fixedRule,
        axis === "height" ? spec : fixedRule, { axis: other, value: fixed })[axis];
    } else if (element.type === "group") {
      const needed = groupExtent(item, axis);
      if (needed > (spec.max ?? Infinity)) return fail(item, "intrinsic-size-unsatisfiable", "Group contents exceed the maximum size without scaling children.");
      value = bounded(needed, spec);
    } else {
      const needed = finite(textSize(item, axis, spec), item);
      const remeasureVerticalWidth = direction !== undefined && direction !== "horizontal" && rule?.width?.type === "content";
      if (axis === "height" && needed > (spec.max ?? Infinity) && !remeasureVerticalWidth) return fail(item, "intrinsic-size-unsatisfiable", "Measured text exceeds maximum height. Increase the bound or change content/font size.");
      value = bounded(needed, spec);
    }
    element.box[axis] = finite(value, item);
    states.set(key, "done"); chain.pop();
    return value;
  }
  function resolveElement(item: Item): void {
    const element = item.element;
    if (finalized.has(element.id)) return;
    dimension(item, "width"); dimension(item, "height");
    if (element.type === "group") {
      const rule = activeRules.get(element.id);
      const before = original.get(element.id)!;
      const origin = groupOrigins.get(element.id) ?? { x: 0, y: 0 };
      const sx = before.box.width === 0 ? 1 : before.box.width / before.coordinateSize!.width;
      const sy = before.box.height === 0 ? 1 : before.box.height / before.coordinateSize!.height;
      if (rule?.width?.type === "content") element.coordinateSize.width = finite(element.box.width / sx, item);
      if (rule?.height?.type === "content") element.coordinateSize.height = finite(element.box.height / sy, item);
      if (element.coordinateSize.width <= 0 || element.coordinateSize.height <= 0) fail(item, "intrinsic-size-unsatisfiable", "Group coordinate sizes must remain positive.");
      if (rule?.width?.type === "content" || rule?.height?.type === "content") {
        const matrix = elementTransform(before.box, element.transform);
        const dw = rule?.width?.type === "content" ? element.box.width - before.box.width : 0;
        const dh = rule?.height?.type === "content" ? element.box.height - before.box.height : 0;
        element.box.x += matrix[0] * sx * origin.x + matrix[2] * sy * origin.y + (matrix[0] - 1) * dw / 2 + matrix[2] * dh / 2;
        element.box.y += matrix[1] * sx * origin.x + matrix[3] * sy * origin.y + matrix[1] * dw / 2 + (matrix[3] - 1) * dh / 2;
        for (const child of element.children) {
          child.box.x -= origin.x; child.box.y -= origin.y;
          if (child.type === "connector") {
            if (!("elementId" in child.start)) { child.start.x -= origin.x; child.start.y -= origin.y; }
            if (!("elementId" in child.end)) { child.end.x -= origin.x; child.end.y -= origin.y; }
            child.route.forEach((point) => { point.x -= origin.x; point.y -= origin.y; });
          }
        }
      }
    }
    if (![element.box.x, element.box.y, element.box.x + element.box.width, element.box.y + element.box.height].every(Number.isFinite)) fail(item, "intrinsic-size-invalid", "Intrinsic sizing exceeded finite coordinates.");
    finalized.add(element.id);
  }
  try {
    for (const [index, rule] of rules.entries()) {
      const item = items.get(rule.elementId);
      if (!item) fail(undefined, "intrinsic-target-missing", `Element "${rule.elementId}" was not found.`, `options.intrinsicSizing.${index}`);
      if (item.hidden && !rule.includeHidden) continue;
      if (["table", "connector", "chart"].includes(item.element.type)) fail(item, "intrinsic-content-unavailable", "Intrinsic rules support text, shapes, images and groups. Use tableSizing for table rows.");
      activeRules.set(rule.elementId, rule);
    }
    for (const id of activeRules.keys()) resolveElement(items.get(id)!);
    return { document, diagnostics };
  } catch (error) {
    if (!(error instanceof SizingFailure)) throw error;
    return { document: source, diagnostics };
  }
}
