import type { Box, Insets, NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import type { LayoutDiagnostic } from "../types/analysis.js";
import type { CompositionAlignment, LayoutComposition } from "../types/composition.js";
import { elementTransform, transformedBounds } from "../analysis/geometry.js";
import { gridTargets, validateGrid } from "./grid.js";
import { constraintTargets, validateConstraints } from "./constraints.js";

const alignments = ["start", "center", "end"];
function finiteBox(box: Box): boolean {
  return [box.x, box.y, box.width, box.height, box.x + box.width, box.y + box.height].every(Number.isFinite) && box.width >= 0 && box.height >= 0;
}
function padding(value: LayoutComposition["padding"]): Insets {
  return typeof value === "number" ? { top: value, right: value, bottom: value, left: value }
    : { top: value?.top ?? 0, right: value?.right ?? 0, bottom: value?.bottom ?? 0, left: value?.left ?? 0 };
}

export function validateComposition(rules: readonly LayoutComposition[] | undefined): void {
  if (rules === undefined) return;
  if (!Array.isArray(rules)) throw new TypeError("composition must be an array of rules");
  for (const rule of rules) {
    if (!rule || !["align", "stack", "distribute", "grid", "constraints"].includes(rule.type)) throw new TypeError("Unknown composition operation");
    if (!rule.container || !["slide", "layout", "master", "group"].includes(rule.container.kind) || typeof rule.container.id !== "string") throw new TypeError("Composition requires a container kind and ID");
    if (rule.type === "grid") validateGrid(rule);
    else if (rule.type === "constraints") validateConstraints(rule);
    else if (!Array.isArray(rule.elementIds) || rule.elementIds.some((id: unknown) => typeof id !== "string") || new Set(rule.elementIds).size !== rule.elementIds.length) throw new TypeError("Composition elementIds must be unique strings");
    if (rule.area !== undefined && (!rule.area || !finiteBox(rule.area))) throw new RangeError("Composition area must have finite coordinates and nonnegative dimensions");
    if (rule.padding !== undefined && typeof rule.padding !== "number" && (rule.padding === null || typeof rule.padding !== "object" || Array.isArray(rule.padding))) throw new TypeError("Composition padding must be a number or insets");
    if (typeof rule.padding === "object" && rule.padding !== null && Object.values(rule.padding).some((value) => value !== undefined && (typeof value !== "number" || !Number.isFinite(value) || value < 0))) throw new RangeError("Composition padding must be finite and nonnegative");
    if (Object.values(padding(rule.padding)).some((value) => !Number.isFinite(value) || value < 0)) throw new RangeError("Composition padding must be finite and nonnegative");
    if (rule.bounds !== undefined && !["frame", "visual"].includes(rule.bounds)) throw new TypeError("Composition bounds must be frame or visual");
    if (rule.includeHidden !== undefined && typeof rule.includeHidden !== "boolean") throw new TypeError("Composition includeHidden must be boolean");
    if (rule.type === "align") {
      if (rule.horizontal === undefined && rule.vertical === undefined) throw new TypeError("Alignment requires at least one axis");
      for (const value of [rule.horizontal, rule.vertical]) if (value !== undefined && !alignments.includes(value)) throw new TypeError("Invalid composition alignment");
    } else if (rule.type !== "grid" && rule.type !== "constraints") {
      if (!["horizontal", "vertical"].includes(rule.direction)) throw new TypeError("Composition direction must be horizontal or vertical");
      if (rule.type === "stack") {
        if (rule.gap !== undefined && (!Number.isFinite(rule.gap) || rule.gap < 0)) throw new RangeError("Stack gap must be finite and nonnegative");
        if (rule.justify !== undefined && ![...alignments, "space-between"].includes(rule.justify)) throw new TypeError("Invalid stack justification");
        if (rule.align !== undefined && !alignments.includes(rule.align)) throw new TypeError("Invalid stack alignment");
      }
    }
  }
}

interface Container {
  elements: NormalizedElement[];
  area: Box;
  hidden: boolean;
  slideId?: string;
  layoutId?: string;
  masterId?: string;
  elementId?: string;
}

function containerIndex(document: NormalizedPresentation): Map<string, Container> {
  const index = new Map<string, Container>();
  function groups(elements: NormalizedElement[], owner: Container): void {
    for (const element of elements) {
      if (element.type !== "group") continue;
      const group: Container = { ...owner, elementId: element.id, elements: element.children,
        area: { x: 0, y: 0, ...element.coordinateSize }, hidden: owner.hidden || element.hidden };
      index.set(`group:${element.id}`, group);
      groups(element.children, group);
    }
  }
  for (const master of document.masters ?? []) {
    const container: Container = { elements: master.elements ?? [], area: { x: 0, y: 0, ...document.size }, hidden: false, masterId: master.id };
    index.set(`master:${master.id}`, container); groups(container.elements, container);
  }
  for (const layout of document.layouts) {
    const container: Container = { elements: layout.elements, area: { x: 0, y: 0, ...document.size }, hidden: false, layoutId: layout.id };
    index.set(`layout:${layout.id}`, container);
    groups(layout.elements, container);
  }
  for (const slide of document.slides) {
    const container: Container = { elements: slide.elements, area: { x: 0, y: 0, ...document.size }, hidden: slide.hidden, slideId: slide.id };
    index.set(`slide:${slide.id}`, container);
    groups(slide.elements, container);
  }
  return index;
}

function fits(size: number, available: number): boolean {
  return size <= available + Number.EPSILON * 8 * Math.max(1, size, available);
}

function offset(space: number, alignment: CompositionAlignment | undefined): number {
  return alignment === "center" ? space / 2 : alignment === "end" ? space : 0;
}

function frameBox(element: NormalizedElement): Box {
  if (element.type !== "connector" || "elementId" in element.start || "elementId" in element.end) return element.box;
  let x = Math.min(element.start.x, element.end.x), right = Math.max(element.start.x, element.end.x);
  let y = Math.min(element.start.y, element.end.y), bottom = Math.max(element.start.y, element.end.y);
  for (const point of element.route) {
    x = Math.min(x, point.x); right = Math.max(right, point.x);
    y = Math.min(y, point.y); bottom = Math.max(bottom, point.y);
  }
  return { x, y, width: right - x, height: bottom - y };
}

function validMove(element: NormalizedElement, x: number, y: number, width = element.box.width, height = element.box.height): boolean {
  if (!finiteBox({ x, y, width, height })) return false;
  if (element.type !== "connector") return true;
  const dx = x - element.box.x;
  const dy = y - element.box.y;
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) return false;
  return [element.start, ...element.route, element.end].every((point) => !("elementId" in point) && Number.isFinite(point.x + dx) && Number.isFinite(point.y + dy));
}

function move(element: NormalizedElement, x: number, y: number): void {
  const dx = x - element.box.x;
  const dy = y - element.box.y;
  element.box = { ...element.box, x, y };
  if (element.type === "connector") {
    if (!("elementId" in element.start)) element.start = { x: element.start.x + dx, y: element.start.y + dy };
    if (!("elementId" in element.end)) element.end = { x: element.end.x + dx, y: element.end.y + dy };
    element.route = element.route.map((point) => ({ x: point.x + dx, y: point.y + dy }));
  }
}

/** Mutates only the private clone passed by the Layout orchestrator. */
export function compose(document: NormalizedPresentation, rules: readonly LayoutComposition[]): LayoutDiagnostic[] {
  const diagnostics: LayoutDiagnostic[] = [];
  const containers = containerIndex(document);
  rules.forEach((rule, ruleIndex) => {
    const container = containers.get(`${rule.container.kind}:${rule.container.id}`);
    const context = { path: `options.composition.${ruleIndex}`, ...(container?.slideId ? { slideId: container.slideId } : {}), ...(container?.layoutId ? { layoutId: container.layoutId } : {}), ...(container?.masterId ? { masterId: container.masterId } : {}), ...(container?.elementId ? { elementId: container.elementId } : {}) };
    const warn = (code: string, message: string): void => { diagnostics.push({ ...context, code, message, severity: "warning" }); };
    if (!container) { warn("composition-container-missing", `Container "${rule.container.id}" was not found. No geometry changed for this rule.`); return; }
    if (container.hidden && !rule.includeHidden) return;
    const indexed = new Map(container.elements.map((element) => [element.id, element]));
    const selected: NormalizedElement[] = [];
    const ids = rule.type === "grid" ? rule.placements.map((placement) => placement.elementId) : rule.type === "constraints" ? rule.elements.map((element) => element.elementId) : rule.elementIds;
    for (const id of ids) {
      const element = indexed.get(id);
      if (!element) { warn("composition-element-missing", `Element "${id}" is not a direct child of "${rule.container.id}". No geometry changed for this rule.`); return; }
      if (!element.hidden || rule.includeHidden) selected.push(element);
    }
    if (selected.length === 0) return;
    if (selected.some((element) => element.type === "connector" && ("elementId" in element.start || "elementId" in element.end))) {
      warn("composition-anchored-connector", "Arrange the referenced shapes instead of independently positioning an anchored connector. No geometry changed for this rule."); return;
    }
    const area = rule.area ?? container.area;
    const inset = padding(rule.padding);
    const inner = { x: area.x + inset.left, y: area.y + inset.top, width: area.width - inset.left - inset.right, height: area.height - inset.top - inset.bottom };
    const boxes = selected.map((element) => {
      const frame = frameBox(element);
      return rule.bounds === "visual" ? transformedBounds(frame, elementTransform(frame, element.transform)) : { ...frame };
    });
    if (!finiteBox(inner) || boxes.some((box) => !finiteBox(box))) { warn("composition-space-unavailable", "Padding or numeric bounds leave no valid composition area. No geometry changed for this rule."); return; }
    let targets = boxes.map((box) => ({ ...box }));
    if (rule.type === "constraints") {
      // Anchored geometry is resolved only after all composition. Do not expose stale frames.
      const selectedIds = new Set(selected.map((element) => element.id));
      for (const item of rule.elements) {
        if (!selectedIds.has(item.elementId)) continue;
        for (const constraint of [item.x, item.y, item.width, item.height]) for (const expression of [constraint?.value, constraint?.min, constraint?.max]) {
          if (expression === undefined || typeof expression === "number") continue;
          for (const term of expression.terms) {
            const reference = term.source === "container" ? undefined : indexed.get(term.source.elementId);
            if (reference?.type === "connector" && ("elementId" in reference.start || "elementId" in reference.end)) { warn("composition-constraint-unresolved-reference", "Constraints cannot reference an anchored connector before final connector resolution. Reference its target elements instead. No geometry changed for this rule."); return; }
          }
        }
      }
      const constrained = constraintTargets(rule, selected.map((element) => element.id), new Map(container.elements.map((element) => [element.id, frameBox(element)])), inner);
      if (!("targets" in constrained)) { warn(constrained.code, constrained.message); return; }
      targets = constrained.targets;
      if (selected.some((element, i) => (element.type === "table" || element.type === "connector") && (targets[i]!.width !== boxes[i]!.width || targets[i]!.height !== boxes[i]!.height))) { warn("composition-constraint-resize-unsupported", "Resize table rows/columns or connector points explicitly; constraints can only translate these elements. No geometry changed for this rule."); return; }
    } else if (rule.type === "grid") {
      const placed = gridTargets(rule, selected.map((element) => element.id), boxes, inner);
      if (!placed) { warn("composition-grid-unsatisfiable", "Grid track bounds, element spans and gaps cannot fit the available area. No geometry changed for this rule."); return; }
      targets = placed;
    } else if (rule.type === "align") {
      if (boxes.some((box) => (rule.horizontal !== undefined && !fits(box.width, inner.width)) || (rule.vertical !== undefined && !fits(box.height, inner.height)))) { warn("composition-overflow", "A selected element is larger than the aligned area. Increase the area or resize the element. No geometry changed for this rule."); return; }
      targets.forEach((box) => {
        if (rule.horizontal !== undefined) box.x = inner.x + offset(inner.width - box.width, rule.horizontal);
        if (rule.vertical !== undefined) box.y = inner.y + offset(inner.height - box.height, rule.vertical);
      });
    } else {
      const horizontal = rule.direction === "horizontal";
      const length = horizontal ? "width" : "height";
      const cross = horizontal ? "height" : "width";
      const coordinate = horizontal ? "x" : "y";
      const crossCoordinate = horizontal ? "y" : "x";
      const total = boxes.reduce((sum, box) => sum + box[length], 0);
      let gap = rule.type === "stack" ? rule.gap ?? 0 : 0;
      const remainder = inner[length] - total - gap * (boxes.length - 1);
      const epsilon = Number.EPSILON * 8 * Math.max(1, inner[length], total);
      const free = Math.max(0, remainder);
      if (!Number.isFinite(remainder) || remainder < -epsilon || (rule.type === "stack" && boxes.some((box) => !fits(box[cross], inner[cross])))) { warn("composition-overflow", "Fixed sizes and gaps exceed available space. Increase the area, reduce gaps, or resize elements. No geometry changed for this rule."); return; }
      const spread = rule.type === "distribute" || rule.justify === "space-between";
      if (spread && boxes.length > 1) gap += free / (boxes.length - 1);
      const justification = rule.type === "stack" && rule.justify !== "space-between" ? rule.justify : undefined;
      let cursor = inner[coordinate] + (spread ? 0 : offset(free, justification));
      targets.forEach((box) => {
        box[coordinate] = cursor;
        if (rule.type === "stack") box[crossCoordinate] = inner[crossCoordinate] + offset(inner[cross] - box[cross], rule.align);
        cursor += box[length] + gap;
      });
    }
    const moves = selected.map((element, i) => ({ element, x: targets[i]!.x + (element.box.x - boxes[i]!.x), y: targets[i]!.y + (element.box.y - boxes[i]!.y), width: rule.type === "constraints" && element.type !== "connector" ? targets[i]!.width : element.box.width, height: rule.type === "constraints" && element.type !== "connector" ? targets[i]!.height : element.box.height }));
    if (moves.some(({ element, x, y, width, height }) => !validMove(element, x, y, width, height))) { warn("composition-space-unavailable", "Placement exceeded finite coordinates. No geometry changed for this rule."); return; }
    for (const placement of moves) {
      move(placement.element, placement.x, placement.y);
      placement.element.box.width = placement.width; placement.element.box.height = placement.height;
    }
  });
  return diagnostics;
}
