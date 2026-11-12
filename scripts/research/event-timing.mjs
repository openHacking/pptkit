// Candidate event/target policy, not a published or native slideshow contract.
import { compileTiming, sampleTiming } from "./timing-model.mjs";
const time = n => { if (!Number.isSafeInteger(n) || n < 0) throw new RangeError("Invalid event time."); return n; };
const id = n => { if (typeof n !== "string" || !n.length) throw new TypeError("Invalid identity."); return n; };
const freeze = x => { if (x && typeof x === "object") { for (const v of Object.values(x)) freeze(v); Object.freeze(x); } return x; };

export function bindEventTiming(pageId, objects, stages) {
  id(pageId);
  if (!Array.isArray(objects) || !Array.isArray(stages) || objects.length > 10000 || stages.length > 1000) throw new RangeError("Event research budget exceeded.");
  const known = new Set();
  for (const objectId of objects) { id(objectId); if (known.has(objectId)) throw new TypeError("Duplicate drawing identity."); known.add(objectId); }
  const seen = new Set(); let total = 0;
  const bound = stages.map((s,i) => {
    id(s.id); if (seen.has(s.id)) throw new TypeError("Duplicate stage identity."); seen.add(s.id);
    const target = s.target;
    if (!target || Object.keys(target).some(k => !["pageId", "elementId"].includes(k)) || target.pageId !== pageId || !known.has(target.elementId)) throw new TypeError("Unsupported, dangling or cross-page target.");
    const trigger = s.trigger;
    if (!trigger || !["enter", "with-previous", "after-previous", "page-click", "object-click"].includes(trigger.type)) throw new TypeError("Unknown trigger.");
    if (Object.keys(trigger).some(k => !["type", ...(trigger.type === "object-click" ? ["elementId"] : [])].includes(k))) throw new TypeError("Unknown trigger field.");
    if (trigger.type === "object-click" && !known.has(trigger.elementId)) throw new TypeError("Dangling event source.");
    if ((i === 0 && ["with-previous", "after-previous"].includes(trigger.type)) || (i !== 0 && trigger.type === "enter")) throw new TypeError("Invalid trigger predecessor.");
    const timing = compileTiming(s.timing, { maxNodes: 10000 - total }); total += timing.nodeCount;
    if (Object.keys(s).some(k => !["id", "target", "trigger", "timing"].includes(k))) throw new TypeError("Unknown stage property.");
    return { id: s.id, target: { pageId, elementId: target.elementId }, trigger: { ...trigger }, timing };
  });
  return freeze({ pageId, stages: bound });
}

/** Candidate sequential gate policy. Events are immutable input, never dispatched. */
export function replayEventTiming(plan, events, atMs) {
  time(atMs);
  if (!Array.isArray(events) || events.length > 10000) throw new RangeError("Event budget exceeded.");
  let previous = -1;
  for (const e of events) {
    time(e.timeMs); id(e.pageId);
    if (e.timeMs < previous || !["page-click", "object-click"].includes(e.type)) throw new TypeError("Invalid ordered event log.");
    if (e.type === "object-click") id(e.elementId);
    if (Object.keys(e).some(k => !["timeMs", "pageId", "type", ...(e.type === "object-click" ? ["elementId"] : [])].includes(k))) throw new TypeError("Unknown event field.");
    previous = e.timeMs;
  }
  const started = []; let cursor = 0;
  for (const s of plan.stages) {
    const prev = started.at(-1); let startMs;
    if (s.trigger.type === "enter") startMs = 0;
    else if (s.trigger.type === "with-previous") { if (!prev) break; startMs = prev.startMs; }
    else if (s.trigger.type === "after-previous") { if (!prev) break; startMs = prev.endMs; }
    else {
      const ready = prev?.endMs ?? 0;
      for (; cursor < events.length; cursor++) {
        const e = events[cursor]; if (e.timeMs > atMs) break;
        if (e.timeMs >= ready && e.pageId === plan.pageId && e.type === s.trigger.type && (e.type !== "object-click" || e.elementId === s.trigger.elementId)) { startMs = e.timeMs; cursor++; break; }
      }
    }
    if (startMs === undefined || startMs > atMs) break;
    const endMs = time(startMs + s.timing.durationMs);
    started.push({ id: s.id, target: s.target, startMs, endMs, samples: sampleTiming(s.timing, atMs - startMs) });
  }
  return started;
}
