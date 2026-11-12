import { PresentationEditError, type PresentationEditAccessibility, type PresentationEditAction, type PresentationEditBox, type PresentationEditTransform, type PresentationEditReadingOrder, type PresentationEditChange, type PresentationEditOperation } from "../types/edit-operation.js";

function fail(message: string): never { throw new PresentationEditError("invalid-operation", message); }
function fields(value: unknown, keys: readonly string[]): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) return fail("Operation fields must be plain records.");
  const names = Reflect.ownKeys(value);
  if (names.length !== keys.length || names.some(name => typeof name !== "string" || !keys.includes(name))) return fail("Operation contains missing or unknown fields.");
  const result: Record<string, unknown> = {};
  for (const key of keys) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !("value" in descriptor) || !descriptor.enumerable) return fail("Operation fields must be enumerable data properties.");
    result[key] = descriptor.value;
  }
  return result;
}
function number(value: unknown): value is number | null { return value === null || (typeof value === "number" && Number.isSafeInteger(value) && value >= 0); }
function box(value: unknown): PresentationEditBox {
  if (value === null) return null;
  const item = fields(value, ["x", "y", "width", "height"]);
  if (typeof item.x !== "number" || !Number.isFinite(item.x) || typeof item.y !== "number" || !Number.isFinite(item.y) || typeof item.width !== "number" || !Number.isFinite(item.width) || item.width < 0 || (item.height !== null && (typeof item.height !== "number" || !Number.isFinite(item.height) || item.height < 0))) return fail("Box requires finite coordinates and nonnegative dimensions.");
  return Object.freeze({ x: item.x, y: item.y, width: item.width, height: item.height as number | null });
}
function transform(value: unknown): PresentationEditTransform {
  if (value === null) return null;
  const item = fields(value, ["rotation", "flipH", "flipV"]);
  if ((item.rotation !== null && (typeof item.rotation !== "number" || !Number.isFinite(item.rotation))) ||
      (item.flipH !== null && typeof item.flipH !== "boolean") || (item.flipV !== null && typeof item.flipV !== "boolean")) return fail("Transform requires finite rotation and boolean flips, or null for absent properties.");
  return Object.freeze({ rotation: item.rotation as number | null, flipH: item.flipH as boolean | null, flipV: item.flipV as boolean | null });
}
function accessibility(value: unknown): PresentationEditAccessibility {
  if (value === null) return null;
  const item = fields(value, ["description", "decorative"]);
  if ((item.description !== null && typeof item.description !== "string") || (item.decorative !== null && typeof item.decorative !== "boolean")) return fail("Accessibility requires a string description and boolean decorative flag, or null for absent properties.");
  return Object.freeze({ description: item.description as string | null, decorative: item.decorative as boolean | null });
}
function action(value: unknown): PresentationEditAction {
  if (value === null) return null;
  if (!value || typeof value !== "object") return fail("Action must be a record or null.");
  const type = Object.getOwnPropertyDescriptor(value, "type");
  if (!type || !("value" in type) || (type.value !== "url" && type.value !== "slide")) return fail("Unknown action type.");
  const item = fields(value, ["type", type.value === "url" ? "url" : "slideId", "tooltip"]);
  if (item.tooltip !== null && typeof item.tooltip !== "string") return fail("Action tooltip must be a string or null.");
  const target = type.value === "url" ? item.url : item.slideId;
  if (typeof target !== "string" || !target.trim()) return fail("Action target must be nonempty.");
  return type.value === "url" ? Object.freeze({ type: "url", url: target, tooltip: item.tooltip as string | null }) : Object.freeze({ type: "slide", slideId: target, tooltip: item.tooltip as string | null });
}
function slideOrder(value: unknown): readonly string[] {
  if (!Array.isArray(value) || value.length > 100000) return fail("Slide order must be an array with at most 100000 IDs.");
  const keys = Reflect.ownKeys(value);
  if (keys.length !== value.length + 1 || keys.some(key => key !== "length" && (typeof key !== "string" || !/^(0|[1-9][0-9]*)$/.test(key) || Number(key) >= value.length))) return fail("Slide order must be a dense array without extra properties.");
  const result: string[] = [], seen = new Set<string>();
  for (let i = 0; i < value.length; i++) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(i));
    if (!descriptor || !("value" in descriptor) || !descriptor.enumerable || typeof descriptor.value !== "string" || !descriptor.value || seen.has(descriptor.value)) return fail("Slide order must contain distinct nonempty IDs as data properties.");
    seen.add(descriptor.value); result.push(descriptor.value);
  }
  return Object.freeze(result);
}
function readingOrder(value: unknown): PresentationEditReadingOrder {
  if (value === null) return null;
  if (!Array.isArray(value) || !value.length || value.length > 10000) return fail("Reading order must be null or a nonempty array of at most 10000 scopes.");
  const keys = Reflect.ownKeys(value);
  if (keys.length !== value.length + 1 || keys.some(key => key !== "length" && (typeof key !== "string" || !/^(0|[1-9][0-9]*)$/.test(key) || Number(key) >= value.length))) return fail("Reading order must be dense without extra properties.");
  const scopes: { groupId?: string; elementIds: readonly string[] }[] = [], owners = new Set<string | undefined>();
  let count = 0;
  for (let i = 0; i < value.length; i++) {
    const entry = Object.getOwnPropertyDescriptor(value, String(i));
    if (!entry || !("value" in entry) || !entry.enumerable || !entry.value || typeof entry.value !== "object") return fail("Reading order scopes must be data records.");
    const hasGroup = Object.prototype.hasOwnProperty.call(entry.value, "groupId"), scope = fields(entry.value, hasGroup ? ["groupId", "elementIds"] : ["elementIds"]);
    if (hasGroup && (typeof scope.groupId !== "string" || !scope.groupId)) return fail("Reading order group ID must be nonempty.");
    const owner = scope.groupId as string | undefined;
    if (owners.has(owner)) return fail("Reading order owners must be unique."); owners.add(owner);
    const elementIds = slideOrder(scope.elementIds); count += elementIds.length;
    if (!elementIds.length || count > 100000) return fail("Reading order requires nonempty scopes and at most 100000 IDs.");
    scopes.push(Object.freeze({ ...(owner === undefined ? {} : { groupId: owner }), elementIds }));
  }
  return Object.freeze(scopes);
}

/** Capture before touching a document; accessors, sparse arrays and unknown properties reject. */
export function captureEditOperation(input: PresentationEditOperation): PresentationEditOperation {
  const root = fields(input, ["schemaVersion", "documentId", "changes"]);
  if (root.schemaVersion !== "pptkit-edit-v1" || typeof root.documentId !== "string" || !root.documentId) return fail("Operation version and document ID are required.");
  if (!Array.isArray(root.changes) || root.changes.length > 10000) return fail("Operation changes must be an array with at most 10000 entries.");
  const array = root.changes, keys = Reflect.ownKeys(array);
  if (keys.length !== array.length + 1 || keys.some(key => key !== "length" && (typeof key !== "string" || !/^(0|[1-9][0-9]*)$/.test(key) || Number(key) >= array.length))) return fail("Operation changes must be a dense array without extra properties.");
  const result: PresentationEditChange[] = [], targets = new Set<string>();
  for (let i = 0; i < array.length; i++) {
    const entry = Object.getOwnPropertyDescriptor(array, String(i));
    if (!entry || !("value" in entry) || !entry.enumerable) return fail("Operation changes must be enumerable data properties.");
    const value = entry.value as unknown;
    if (!value || typeof value !== "object") return fail("Invalid operation change.");
    const kind = Object.getOwnPropertyDescriptor(value, "type");
    if (!kind || !("value" in kind)) return fail("Change type must be a data property.");
    if (kind.value === "element-hidden" || kind.value === "element-opacity" || kind.value === "element-name") {
      const change = fields(value, ["type", "slideId", "elementId", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId) return fail("Invalid element target.");
      const key = JSON.stringify([kind.value, change.slideId, change.elementId]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      if (kind.value === "element-name") {
        if ((change.expected !== null && typeof change.expected !== "string") || (change.value !== null && typeof change.value !== "string")) return fail("Element name must be a string or null.");
        result.push(Object.freeze({ type: "element-name", slideId: change.slideId, elementId: change.elementId, expected: change.expected, value: change.value }));
      } else if (kind.value === "element-hidden") {
        if ((change.expected !== null && typeof change.expected !== "boolean") || (change.value !== null && typeof change.value !== "boolean")) return fail("Element visibility must be boolean or null.");
        result.push(Object.freeze({ type: "element-hidden", slideId: change.slideId, elementId: change.elementId, expected: change.expected, value: change.value }));
      } else {
        const valid = (v: unknown): v is number | null => v === null || typeof v === "number" && Number.isFinite(v) && v >= 0 && v <= 1;
        if (!valid(change.expected) || !valid(change.value)) return fail("Element opacity must be null or a finite number in 0..1.");
        result.push(Object.freeze({ type: "element-opacity", slideId: change.slideId, elementId: change.elementId, expected: change.expected, value: change.value }));
      }
    } else if (kind.value === "element-action") {
      const change = fields(value, ["type", "slideId", "elementId", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId) return fail("Invalid element action target.");
      const key = JSON.stringify(["element-action", change.slideId, change.elementId]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      result.push(Object.freeze({ type: "element-action", slideId: change.slideId, elementId: change.elementId, expected: action(change.expected), value: action(change.value) }));
    } else if (kind.value === "chart-category" || kind.value === "chart-series-name") {
      const indexKey = kind.value === "chart-category" ? "categoryIndex" : "seriesIndex";
      const change = fields(value, ["type", "slideId", "elementId", indexKey, "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId ||
        !Number.isSafeInteger(change[indexKey]) || (change[indexKey] as number) < 0 || typeof change.expected !== "string" || typeof change.value !== "string") return fail("Invalid chart text change.");
      const key = JSON.stringify([kind.value, change.slideId, change.elementId, change[indexKey]]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      const common = { slideId: change.slideId, elementId: change.elementId, expected: change.expected, value: change.value };
      result.push(kind.value === "chart-category" ? Object.freeze({ type: "chart-category", ...common, categoryIndex: change.categoryIndex as number }) : Object.freeze({ type: "chart-series-name", ...common, seriesIndex: change.seriesIndex as number }));
    } else if (kind.value === "chart-point") {
      const change = fields(value, ["type", "slideId", "elementId", "seriesIndex", "pointIndex", "coordinate", "expected", "value"]);
      const numberOrNull = (v: unknown): v is number | null | "-0" => v === null || v === "-0" || typeof v === "number" && Number.isFinite(v) && !Object.is(v, -0);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId ||
        !Number.isSafeInteger(change.seriesIndex) || (change.seriesIndex as number) < 0 || !Number.isSafeInteger(change.pointIndex) || (change.pointIndex as number) < 0 ||
        !["values", "xValues", "yValues", "sizes"].includes(change.coordinate as string) || !numberOrNull(change.expected) || !numberOrNull(change.value)) return fail("Invalid chart point change.");
      const key = JSON.stringify(["chart-point", change.slideId, change.elementId, change.seriesIndex, change.pointIndex, change.coordinate]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      result.push(Object.freeze({ type: "chart-point", slideId: change.slideId, elementId: change.elementId, seriesIndex: change.seriesIndex as number, pointIndex: change.pointIndex as number, coordinate: change.coordinate as "values" | "xValues" | "yValues" | "sizes", expected: change.expected, value: change.value }));
    } else if (kind.value === "element-accessibility") {
      const change = fields(value, ["type", "slideId", "elementId", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId) return fail("Invalid element accessibility target.");
      const key = JSON.stringify(["element-accessibility", change.slideId, change.elementId]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      result.push(Object.freeze({ type: "element-accessibility", slideId: change.slideId, elementId: change.elementId, expected: accessibility(change.expected), value: accessibility(change.value) }));
    } else if (kind.value === "element-transform") {
      const change = fields(value, ["type", "slideId", "elementId", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId) return fail("Invalid element transform target.");
      const key = JSON.stringify(["element-transform", change.slideId, change.elementId]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      result.push(Object.freeze({ type: "element-transform", slideId: change.slideId, elementId: change.elementId, expected: transform(change.expected), value: transform(change.value) }));
    } else if (kind.value === "element-box") {
      const change = fields(value, ["type", "slideId", "elementId", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId) return fail("Invalid element box target.");
      const key = JSON.stringify(["element-box", change.slideId, change.elementId]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      result.push(Object.freeze({ type: "element-box", slideId: change.slideId, elementId: change.elementId, expected: box(change.expected), value: box(change.value) }));
    } else if (kind.value === "text-run" || kind.value === "table-cell-text-run" || kind.value === "text-run-action" || kind.value === "table-cell-text-run-action") {
      const table = kind.value === "table-cell-text-run" || kind.value === "table-cell-text-run-action";
      const isAction = kind.value === "text-run-action" || kind.value === "table-cell-text-run-action";
      const change = fields(value, ["type", "slideId", "elementId", ...(table ? ["rowIndex", "cellIndex"] : []), "paragraphIndex", "runIndex", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId || !Number.isSafeInteger(change.paragraphIndex) || (change.paragraphIndex as number) < 0 || !Number.isSafeInteger(change.runIndex) || (change.runIndex as number) < 0 || (!isAction && (typeof change.expected !== "string" || typeof change.value !== "string"))) return fail("Invalid text run change.");
      if (table && (!Number.isSafeInteger(change.rowIndex) || (change.rowIndex as number) < 0 || !Number.isSafeInteger(change.cellIndex) || (change.cellIndex as number) < 0)) return fail("Invalid table cell path.");
      const key = JSON.stringify([kind.value, change.slideId, change.elementId, ...(table ? [change.rowIndex, change.cellIndex] : []), change.paragraphIndex, change.runIndex]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      const common = { slideId: change.slideId, elementId: change.elementId, paragraphIndex: change.paragraphIndex as number, runIndex: change.runIndex as number };
      if (isAction) {
        const values = { expected: action(change.expected), value: action(change.value) };
        result.push(table ? Object.freeze({ type: "table-cell-text-run-action", ...common, ...values, rowIndex: change.rowIndex as number, cellIndex: change.cellIndex as number }) : Object.freeze({ type: "text-run-action", ...common, ...values }));
      } else {
        const values = { expected: change.expected as string, value: change.value as string };
        result.push(table ? Object.freeze({ type: "table-cell-text-run", ...common, ...values, rowIndex: change.rowIndex as number, cellIndex: change.cellIndex as number }) : Object.freeze({ type: "text-run", ...common, ...values }));
      }
    } else if (kind.value === "text-string" || kind.value === "table-cell-text-string") {
      const table = kind.value === "table-cell-text-string";
      const change = fields(value, ["type", "slideId", "elementId", ...(table ? ["rowIndex", "cellIndex"] : []), "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.elementId !== "string" || !change.elementId || typeof change.expected !== "string" || typeof change.value !== "string") return fail("Invalid text string change.");
      if (table && (!Number.isSafeInteger(change.rowIndex) || (change.rowIndex as number) < 0 || !Number.isSafeInteger(change.cellIndex) || (change.cellIndex as number) < 0)) return fail("Invalid table cell path.");
      const key = JSON.stringify([kind.value, change.slideId, change.elementId, ...(table ? [change.rowIndex, change.cellIndex] : [])]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      const common = { slideId: change.slideId, elementId: change.elementId, expected: change.expected, value: change.value };
      result.push(table ? Object.freeze({ type: "table-cell-text-string", ...common, rowIndex: change.rowIndex as number, cellIndex: change.cellIndex as number }) : Object.freeze({ type: "text-string", ...common }));
    } else if (kind.value === "notes-string" || kind.value === "notes-run" || kind.value === "notes-run-action") {
      const isRun = kind.value !== "notes-string", isAction = kind.value === "notes-run-action";
      const change = fields(value, ["type", "slideId", ...(isRun ? ["paragraphIndex", "runIndex"] : []), "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || (isRun && (!Number.isSafeInteger(change.paragraphIndex) || (change.paragraphIndex as number) < 0 || !Number.isSafeInteger(change.runIndex) || (change.runIndex as number) < 0)) || (!isAction && (typeof change.expected !== "string" || typeof change.value !== "string"))) return fail("Invalid notes change.");
      const key = JSON.stringify([kind.value, change.slideId, ...(isRun ? [change.paragraphIndex, change.runIndex] : [])]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      if (kind.value === "notes-string") result.push(Object.freeze({ type: "notes-string", slideId: change.slideId, expected: change.expected as string, value: change.value as string }));
      else {
        const common = { slideId: change.slideId, paragraphIndex: change.paragraphIndex as number, runIndex: change.runIndex as number };
        if (isAction) result.push(Object.freeze({ type: "notes-run-action", ...common, expected: action(change.expected), value: action(change.value) }));
        else result.push(Object.freeze({ type: "notes-run", ...common, expected: change.expected as string, value: change.value as string }));
      }
    } else if (kind.value === "slide-reading-order") {
      const change = fields(value, ["type", "slideId", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId) return fail("Invalid reading order slide ID.");
      const key = JSON.stringify(["slide-reading-order", change.slideId]);
      if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      result.push(Object.freeze({ type: "slide-reading-order", slideId: change.slideId, expected: readingOrder(change.expected), value: readingOrder(change.value) }));
    } else if (kind.value === "slide-hidden") {
      const change = fields(value, ["type", "slideId", "expected", "value"]);
      if (typeof change.slideId !== "string" || !change.slideId || typeof change.expected !== "boolean" || typeof change.value !== "boolean") return fail("Invalid slide visibility change.");
      const key = `slide:${change.slideId}`; if (targets.has(key)) return fail("Operation targets must be unique."); targets.add(key);
      result.push(Object.freeze({ type: "slide-hidden", slideId: change.slideId, expected: change.expected, value: change.value }));
    } else if (kind.value === "first-slide-number") {
      const change = fields(value, ["type", "expected", "value"]);
      if (!number(change.expected) || !number(change.value)) return fail("Slide number must be null or a nonnegative safe integer.");
      if (targets.has("number")) return fail("Operation targets must be unique."); targets.add("number");
      result.push(Object.freeze({ type: "first-slide-number", expected: change.expected, value: change.value }));
    } else if (kind.value === "slide-order") {
      const change = fields(value, ["type", "expected", "value"]);
      if (targets.has("order")) return fail("Operation targets must be unique."); targets.add("order");
      const expected = slideOrder(change.expected), next = slideOrder(change.value), expectedIds = new Set(expected);
      if (expected.length !== next.length || next.some(id => !expectedIds.has(id))) return fail("Slide order must be a permutation of the expected order.");
      result.push(Object.freeze({ type: "slide-order", expected, value: next }));
    } else return fail("Unknown operation change type.");
  }
  return Object.freeze({ schemaVersion: "pptkit-edit-v1", documentId: root.documentId, changes: Object.freeze(result) });
}
