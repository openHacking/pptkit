import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";

type Settings = { playbackVolume: "low" | "medium" | "high" | "mute"; loop: boolean; rewindAfterPlayback: boolean };
type Result = { status: "mapped"; settings: Settings } | { status: "unsupported" };

/** Interpret only direct root media nodes with the observed native audio timing shape. */
export async function audioPlayback(root: XmlElement, p: string, checkpoint: () => Promise<void>): Promise<ReadonlyMap<string, Result>> {
  const selected = (node: XmlElement, local: string) => children(node).filter(x => x.namespaceUri === p && x.local === local);
  const one = (node: XmlElement, local: string) => { const list = selected(node, local); return list.length === 1 ? list[0] : undefined; };
  const map = new Map<string, Result>();
  const canonicalId = (value: string | undefined) => {
    if (value === undefined || !/^\d+$/u.test(value)) return undefined;
    const n = Number(value);
    return Number.isSafeInteger(n) && n <= 4294967295 ? String(n) : undefined;
  };
  const timings = selected(root, "timing");
  if (timings.length !== 1) return map;
  const timeIds = new Map<string, number>(), pending = [timings[0]!];
  while (pending.length) {
    await checkpoint();
    const node = pending.pop()!;
    if (node.namespaceUri === p && node.local === "cTn") {
      const id = canonicalId(attribute(node, "id"));
      if (id !== undefined) timeIds.set(id, (timeIds.get(id) ?? 0) + 1);
    }
    for (const child of children(node)) pending.push(child);
  }
  const list = one(timings[0]!, "tnLst"), par = list && one(list, "par"), time = par && one(par, "cTn");
  if (!time || attribute(time, "nodeType") !== "tmRoot") return map;
  const timeline = one(time, "childTnLst");
  if (!timeline) return map;
  for (const audio of selected(timeline, "audio")) {
    await checkpoint();
    const media = one(audio, "cMediaNode"), target = media && one(media, "tgtEl"), shape = target && one(target, "spTgt");
    const id = shape && canonicalId(attribute(shape, "spid"));
    if (id === undefined) continue;
    if (map.has(id)) { map.set(id, { status: "unsupported" }); continue; }
    const t = media && one(media, "cTn"), start = t && one(t, "stCondLst"), end = t && one(t, "endCondLst");
    const startConds = start && selected(start, "cond"), endConds = end && selected(end, "cond");
    const allowed = (node: XmlElement, names: readonly string[]) => node.attributes.every(x => x.namespaceUri === "" && names.includes(x.local));
    const vol = media && attribute(media, "vol"), mute = media && attribute(media, "mute"), fill = t && attribute(t, "fill"), repeat = t && attribute(t, "repeatCount");
    const volume = mute === "1" || mute === "true" ? vol === undefined ? "mute" : undefined : mute === undefined || mute === "0" || mute === "false" ? vol === undefined || vol === "50000" ? "medium" : vol === "20000" ? "low" : vol === "80000" ? "high" : undefined : undefined;
    const simpleTarget = shape && children(target!).length === 1 && children(shape).length === 0 && allowed(target!, []) && allowed(shape, ["spid"]);
    const endTarget = endConds?.length === 1 ? one(endConds[0]!, "tgtEl") : undefined;
    const slideTarget = endTarget && one(endTarget, "sldTgt");
    const timeId = t && canonicalId(attribute(t, "id"));
    const valid = media && t && timeId !== undefined && timeIds.get(timeId) === 1 && simpleTarget && children(audio).length === 1 && allowed(audio, []) && children(media).length === 2 && allowed(media, ["vol", "mute"]) && allowed(t, ["id", "fill", "display", "repeatCount"]) && (fill === "hold" || fill === "remove") && (repeat === undefined || repeat === "indefinite") && attribute(t, "display") === "0" && startConds?.length === 1 && children(start!).length === 1 && attribute(startConds[0]!, "delay") === "indefinite" && allowed(startConds[0]!, ["delay"]) && children(startConds[0]!).length === 0 && endConds?.length === 1 && children(end!).length === 1 && attribute(endConds[0]!, "evt") === "onStopAudio" && attribute(endConds[0]!, "delay") === "0" && allowed(endConds[0]!, ["evt", "delay"]) && children(endConds[0]!).length === 1 && endTarget && allowed(endTarget, []) && children(endTarget).length === 1 && slideTarget && allowed(slideTarget, []) && children(slideTarget).length === 0 && children(t).length === 2 && volume !== undefined;
    map.set(id, valid ? { status: "mapped", settings: { playbackVolume: volume!, loop: repeat === "indefinite", rewindAfterPlayback: fill === "remove" } } : { status: "unsupported" });
  }
  return map;
}
