import { FiniteTimelineError, type FiniteTimelineErrorCode, type FiniteTimelineLimits } from "../types/timeline.js";
import { EventTimelineError, type EventTimelineErrorCode, type EventTimelineLimits, type EventTimelineTarget, type EventTimelineTrigger } from "../types/event-timeline.js";
import { array, fail, id, integer, limits, record, type Limits } from "./input.js";

export function eventError(error: unknown, prefix = ""): never {
  if (error instanceof FiniteTimelineError) {
    const codes: Record<FiniteTimelineErrorCode, EventTimelineErrorCode> = { "invalid-timeline-input": "invalid-event-timeline-input", "unsupported-timeline-feature": "unsupported-event-timeline-feature", "timeline-limit-exceeded": "event-timeline-limit-exceeded", "timeline-time-overflow": "event-timeline-time-overflow", "invalid-timeline-plan": "invalid-event-timeline-plan" };
    throw new EventTimelineError(codes[error.code], error.message, prefix + error.path, { cause: error });
  }
  throw error;
}
export type EventLimits = Limits & Required<Pick<EventTimelineLimits, "maxElements" | "maxStages" | "maxEvents">>;
export class EventCapture {
  readonly budget: EventLimits;
  characters = 0;
  nodes = 0;
  constructor(options: EventTimelineLimits | undefined) {
    const defaults = { maxElements: 10000, maxStages: 1000, maxEvents: 10000 };
    const caps = { maxElements: 100000, maxStages: 10000, maxEvents: 100000 };
    const finiteKeys = ["maxNodes", "maxDepth", "maxTimeMs", "maxRepeat", "maxIdLength", "maxTotalIdLength"];
    const data = record(options === undefined ? {} : options, [...finiteKeys, ...Object.keys(defaults)], "/limits");
    const finite = Object.fromEntries(Object.entries(data).filter(([key]) => finiteKeys.includes(key)));
    for (const key of Object.keys(defaults) as (keyof typeof defaults)[]) {
      const value = integer(data[key] === undefined ? defaults[key] : data[key], `/limits/${key}`, 1);
      if (value > caps[key]) fail("invalid-timeline-input", `/limits/${key}`, "Event limit exceeds its supported cap.");
      defaults[key] = value;
    }
    this.budget = { ...limits(finite), ...defaults };
  }
  identity(value: unknown, path: string, code: FiniteTimelineErrorCode = "invalid-timeline-input"): string {
    const captured = id(value, path, this.budget, code);
    this.charge(captured.length, path);
    return captured;
  }
  charge(characters: number, path: string): void {
    this.characters += characters;
    if (this.characters > this.budget.maxTotalIdLength) fail("timeline-limit-exceeded", path, "Event input exceeds maxTotalIdLength.");
  }
  inventory(value: unknown, path: string, code: FiniteTimelineErrorCode = "invalid-timeline-input"): readonly string[] {
    const values = array(value, this.budget.maxElements, path, code), known = new Set<string>(), result: string[] = [];
    for (let i = 0; i < values.length; i++) {
      const identity = this.identity(values[i], `${path}/${i}`, code);
      if (known.has(identity)) fail(code, `${path}/${i}`, "Duplicate whole-object inventory identity.");
      known.add(identity); result.push(identity);
    }
    return Object.freeze(result);
  }
  finiteOptions(path: string): FiniteTimelineLimits {
    const nodes = this.budget.maxNodes - this.nodes, characters = this.budget.maxTotalIdLength - this.characters;
    if (nodes <= 0 || characters <= 0) fail("timeline-limit-exceeded", path, "Event timing exhausted its aggregate node/identifier budget.");
    return { maxNodes: nodes, maxTotalIdLength: characters, maxDepth: this.budget.maxDepth, maxTimeMs: this.budget.maxTimeMs, maxRepeat: this.budget.maxRepeat, maxIdLength: this.budget.maxIdLength };
  }
}
interface CapturedStage { readonly id: string; readonly target: EventTimelineTarget; readonly trigger: EventTimelineTrigger; readonly timing: unknown }
export function captureEventProgram(value: unknown, context: EventCapture, compiled: boolean) {
  const code = compiled ? "invalid-timeline-plan" : "invalid-timeline-input", path = compiled ? "/plan" : "";
  const data = record(value, ["pageId", "elementIds", "stages", ...(compiled ? ["schemaVersion", "timingNodeCount"] : [])], path, code);
  if (compiled && data.schemaVersion !== "pptkit-event-timeline-v1") fail(code, `${path}/schemaVersion`, "Unknown event timeline plan schema.");
  const pageId = context.identity(data.pageId, `${path}/pageId`, code), elementIds = context.inventory(data.elementIds, `${path}/elementIds`, code), elements = new Set(elementIds);
  const rows = array(data.stages, context.budget.maxStages, `${path}/stages`, code), ids = new Set<string>(), stages: CapturedStage[] = [];
  for (let i = 0; i < rows.length; i++) {
    const at = `${path}/stages/${i}`, row = record(rows[i], ["id", "target", "trigger", "timing"], at, code), identity = context.identity(row.id, `${at}/id`, code);
    if (ids.has(identity)) fail(code, `${at}/id`, "Duplicate event stage identity."); ids.add(identity);
    const rawTarget = record(row.target, ["pageId", "elementId"], `${at}/target`, code);
    const targetPage = context.identity(rawTarget.pageId, `${at}/target/pageId`, code), element = context.identity(rawTarget.elementId, `${at}/target/elementId`, code);
    if (targetPage !== pageId || !elements.has(element)) fail(code, `${at}/target`, "Cross-page or dangling whole-object target.");
    const trigger = record(row.trigger, ["type", "elementId"], `${at}/trigger`, code), type = trigger.type;
    if (!["enter", "with-previous", "after-previous", "page-click", "object-click"].includes(type as string)) fail(compiled ? code : "unsupported-timeline-feature", `${at}/trigger/type`, "Unknown event trigger.");
    if (type !== "object-click" && Object.hasOwn(trigger, "elementId")) fail(compiled ? code : "unsupported-timeline-feature", `${at}/trigger/elementId`, "Only object-click triggers have a source element.");
    if ((i === 0 && (type === "with-previous" || type === "after-previous")) || (i !== 0 && type === "enter")) fail(code, `${at}/trigger`, "Invalid trigger predecessor.");
    let source: string | undefined;
    if (type === "object-click") { source = context.identity(trigger.elementId, `${at}/trigger/elementId`, code); if (!elements.has(source)) fail(code, `${at}/trigger/elementId`, "Dangling object-click event source."); }
    const capturedTrigger = Object.freeze(source === undefined ? { type } : { type, elementId: source }) as EventTimelineTrigger;
    stages.push({ id: identity, target: Object.freeze({ pageId, elementId: element }), trigger: capturedTrigger, timing: row.timing });
  }
  return { pageId, elementIds, stages, ...(compiled ? { timingNodeCount: integer(data.timingNodeCount, `${path}/timingNodeCount`, 0, code) } : {}) };
}
