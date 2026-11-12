import type { Box } from "@pptkit/core";
import type { ConstraintExpression, ConstraintMetric, ConstraintProperty, ConstraintsComposition, FrameConstraint } from "../types/composition.js";

const properties: ConstraintProperty[] = ["x", "y", "width", "height"];
const metrics: ConstraintMetric[] = [...properties, "right", "bottom", "centerX", "centerY"];
const components = (metric: ConstraintMetric): ConstraintProperty[] => metric === "right" || metric === "centerX" ? ["x", "width"] : metric === "bottom" || metric === "centerY" ? ["y", "height"] : [metric];
function metricValue(box: Box, metric: ConstraintMetric): number {
  if (metric === "right") return box.x + box.width;
  if (metric === "bottom") return box.y + box.height;
  if (metric === "centerX") return box.x + box.width / 2;
  if (metric === "centerY") return box.y + box.height / 2;
  return box[metric];
}

export function validateConstraints(rule: ConstraintsComposition): void {
  if (rule.bounds !== undefined && rule.bounds !== "frame") throw new TypeError("Constraints require frame bounds");
  if (!Array.isArray(rule.elements)) throw new TypeError("Constraints require an elements array");
  const ids = new Set<string>();
  for (const element of rule.elements) {
    if (!element || typeof element.elementId !== "string" || ids.has(element.elementId)) throw new TypeError("Constrained elements require unique IDs");
    ids.add(element.elementId);
    if (!properties.some((property) => element[property] !== undefined)) throw new TypeError("Constrained elements require at least one property");
    for (const property of properties) {
      const constraint = element[property];
      if (constraint === undefined) continue;
      if (!constraint || typeof constraint !== "object" || Array.isArray(constraint) || [constraint.value, constraint.min, constraint.max].every((value) => value === undefined)) throw new TypeError("Frame constraints require value, min or max");
      for (const expression of [constraint.value, constraint.min, constraint.max]) {
        if (expression === undefined) continue;
        if (typeof expression === "number") {
          if (!Number.isFinite(expression)) throw new RangeError("Constraint constants must be finite");
          continue;
        }
        if (!expression || !Array.isArray(expression.terms) || (expression.constant !== undefined && !Number.isFinite(expression.constant))) throw new TypeError("Invalid affine constraint expression");
        for (const term of expression.terms) {
          if (!term || !metrics.includes(term.metric) || (term.source !== "container" && (!term.source || typeof term.source.elementId !== "string"))) throw new TypeError("Invalid constraint reference");
          if (term.factor !== undefined && !Number.isFinite(term.factor)) throw new RangeError("Constraint factors must be finite");
        }
      }
    }
  }
}

interface Failure { code: string; message: string }
type Result = { targets: Box[] } | Failure;
interface Node { id: string; property: ConstraintProperty; constraint: FrameConstraint; dependencies: Set<string> }
const key = (id: string, property: ConstraintProperty): string => JSON.stringify([id, property]);

/** Resolve a directed affine dependency graph without recursion or source mutation. */
export function constraintTargets(rule: ConstraintsComposition, ids: readonly string[], incoming: ReadonlyMap<string, Box>, area: Box): Result {
  const boxes = new Map([...incoming].map(([id, box]) => [id, { ...box }]));
  const nodes = new Map<string, Node>();
  const selected = new Set(ids);
  for (const element of rule.elements) {
    if (!selected.has(element.elementId)) continue;
    for (const property of properties) {
      const constraint = element[property];
      if (constraint) nodes.set(key(element.elementId, property), { id: element.elementId, property, constraint, dependencies: new Set() });
    }
  }
  const ordered = [...nodes.keys()].sort();
  for (const name of ordered) {
    const node = nodes.get(name)!;
    for (const expression of [node.constraint.value, node.constraint.min, node.constraint.max]) {
      if (expression === undefined || typeof expression === "number") continue;
      for (const term of expression.terms) {
        if (term.source === "container") continue;
        const id = term.source.elementId;
        if (!boxes.has(id)) return { code: "composition-constraint-reference-missing", message: `Constraint for ${node.id}.${node.property} references "${id}", which is not a direct child of this container. No geometry changed for this rule.` };
        if (term.factor === 0) continue;
        for (const property of components(term.metric)) {
          const dependency = key(id, property);
          if (nodes.has(dependency)) node.dependencies.add(dependency);
        }
      }
    }
  }
  const pending = new Map(ordered.map((name) => [name, nodes.get(name)!.dependencies.size]));
  const dependents = new Map<string, string[]>();
  for (const name of ordered) for (const dependency of nodes.get(name)!.dependencies) {
    const list = dependents.get(dependency) ?? []; list.push(name); dependents.set(dependency, list);
  }
  const queue = ordered.filter((name) => pending.get(name) === 0);
  let cursor = 0;
  const evaluate = (expression: ConstraintExpression): number => {
    if (typeof expression === "number") return expression;
    let value = expression.constant ?? 0;
    for (const term of expression.terms) {
      if (term.factor === 0) continue;
      value += metricValue(term.source === "container" ? area : boxes.get(term.source.elementId)!, term.metric) * (term.factor ?? 1);
    }
    return value;
  };
  while (cursor < queue.length) {
    const name = queue[cursor++]!;
    const node = nodes.get(name)!;
    const box = boxes.get(node.id)!;
    const lower = node.constraint.min === undefined ? (node.property === "width" || node.property === "height" ? 0 : -Infinity) : evaluate(node.constraint.min);
    const upper = node.constraint.max === undefined ? Infinity : evaluate(node.constraint.max);
    const value = node.constraint.value === undefined ? Math.max(lower, Math.min(upper, box[node.property])) : evaluate(node.constraint.value);
    if ((node.constraint.min !== undefined && !Number.isFinite(lower)) || (node.constraint.max !== undefined && !Number.isFinite(upper)) || !Number.isFinite(value)) return { code: "composition-constraint-invalid", message: `Constraint ${node.id}.${node.property} produced nonfinite arithmetic. No geometry changed for this rule.` };
    const tolerance = Number.EPSILON * 16 * Math.max(1, Math.abs(value), Number.isFinite(lower) ? Math.abs(lower) : 0, Number.isFinite(upper) ? Math.abs(upper) : 0);
    if (lower > upper || value < lower - tolerance || value > upper + tolerance || ((node.property === "width" || node.property === "height") && value < 0)) return { code: "composition-constraint-unsatisfiable", message: `Constraint ${node.id}.${node.property} cannot satisfy its equality and bounds. No geometry changed for this rule.` };
    box[node.property] = value;
    pending.delete(name);
    for (const dependent of dependents.get(name) ?? []) {
      const count = pending.get(dependent)! - 1;
      pending.set(dependent, count);
      if (count === 0) queue.push(dependent);
    }
  }
  if (pending.size) {
    const path: string[] = [];
    const visited = new Map<string, number>();
    let current = [...pending.keys()][0]!;
    while (!visited.has(current)) {
      visited.set(current, path.length); path.push(current);
      current = [...nodes.get(current)!.dependencies].filter((name) => pending.has(name)).sort()[0]!;
    }
    const cycle = [...path.slice(visited.get(current)!), current].map((name) => { const node = nodes.get(name)!; return `${node.id}.${node.property}`; });
    return { code: "composition-constraint-cycle", message: `Cyclic constraint dependency: ${cycle.join(" -> ")}. No geometry changed for this rule.` };
  }
  const targets = ids.map((id) => boxes.get(id)!);
  if (targets.some((box) => ![box.x, box.y, box.width, box.height, box.x + box.width, box.y + box.height].every(Number.isFinite))) return { code: "composition-constraint-invalid", message: "Constraint frames exceed finite coordinates. No geometry changed for this rule." };
  return { targets };
}
