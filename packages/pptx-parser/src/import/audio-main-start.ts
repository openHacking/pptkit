import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";

export interface MainAudioEffect { readonly target: string; readonly durationMs: number; readonly start: "click-sequence" | "automatic" }

/** Interpret only complete ordered native media calls, with verified optional video pause controls. */
export async function audioMainStart(root: XmlElement, p: string, checkpoint: () => Promise<void>, videoPauses?: ReadonlyMap<string, { status: "mapped"; durationMs: number } | { status: "unsupported" }>, videoObjects?: ReadonlySet<string>, videoSettings?: ReadonlyMap<string, { status: "mapped" } | { status: "unsupported" }>): Promise<readonly MainAudioEffect[] | undefined> {
  const select = (node: XmlElement | undefined, name: string) => node ? children(node).filter(x => x.namespaceUri === p && x.local === name) : [];
  const one = (node: XmlElement | undefined, name: string) => { const nodes = select(node, name); return nodes.length === 1 ? nodes[0] : undefined; };
  const canonical = (text: string | undefined) => { if (text === undefined || !/^\d+$/u.test(text)) return undefined; const n = Number(text); return Number.isSafeInteger(n) && n <= 4294967295 ? String(n) : undefined; };
  const exact = (node: XmlElement | undefined, name: string, attrs: Record<string, string>, kids: readonly string[]): node is XmlElement => {
    if (!node || node.namespaceUri !== p || node.local !== name || node.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim() !== "")) return false;
    if (node.attributes.length !== Object.keys(attrs).length || node.attributes.some(a => a.namespaceUri !== "" || !Object.hasOwn(attrs, a.local) || (attrs[a.local] !== "#" && a.value !== attrs[a.local]))) return false;
    const cs = children(node); return cs.length === kids.length && cs.every((x, i) => x.namespaceUri === p && x.local === kids[i]);
  };
  const timings = select(root, "timing"); if (timings.length !== 1) return undefined;
  const ids = new Map<string, number>(), commands = new Map<string, number>(), pending = [timings[0]!];
  while (pending.length) {
    await checkpoint(); const n = pending.pop()!;
    if (n.namespaceUri === p && n.local === "cTn") { const id = canonical(attribute(n, "id")); if (id !== undefined) ids.set(id, (ids.get(id) ?? 0) + 1); }
    if (n.namespaceUri === p && n.local === "cmd") { const shape = one(one(one(n, "cBhvr"), "tgtEl"), "spTgt"), target = shape && canonical(attribute(shape, "spid")); if (target !== undefined) commands.set(target, (commands.get(target) ?? 0) + 1); }
    for (const child of children(n)) pending.push(child);
  }
  const timeNode = (n: XmlElement | undefined, attrs: Record<string, string>, kids: readonly string[]) => {
    if (!exact(n, "cTn", { id: "#", ...attrs }, kids)) return false;
    const id = canonical(attribute(n, "id")); return id !== undefined && ids.get(id) === 1;
  };
  const list = one(timings[0], "tnLst"), rootPar = one(list, "par"), rootTime = one(rootPar, "cTn");
  if (!exact(timings[0], "timing", {}, ["tnLst"]) || !exact(list, "tnLst", {}, ["par"]) || !exact(rootPar, "par", {}, ["cTn"]) || !timeNode(rootTime, { dur: "indefinite", restart: "never", nodeType: "tmRoot" }, ["childTnLst"])) return undefined;
  const timeline = one(rootTime, "childTnLst");
  if (!timeline || timeline.attributes.length || children(timeline).some(n => n.namespaceUri !== p || !["seq", "audio", "video"].includes(n.local))) return undefined;
  const videos = new Set<string>();
  for (const video of select(timeline, "video")) {
    await checkpoint();
    const media = one(video, "cMediaNode"), time = one(media, "cTn"), start = one(time, "stCondLst"), target = one(media, "tgtEl"), shape = one(target, "spTgt");
    const id = shape && canonical(attribute(shape, "spid"));
    if (!exact(video, "video", {}, ["cMediaNode"]) || id === undefined || videoSettings?.get(id)?.status !== "mapped" || !exact(media, "cMediaNode", Object.fromEntries(media!.attributes.map(a => [a.local, a.value])), ["cTn", "tgtEl"]) || !timeNode(time, Object.fromEntries(time!.attributes.filter(a => a.local !== "id").map(a => [a.local, a.value])), ["stCondLst"]) || !exact(start, "stCondLst", {}, ["cond"]) || !exact(one(start, "cond"), "cond", { delay: "indefinite" }, []) || !exact(target, "tgtEl", {}, ["spTgt"]) || !exact(shape, "spTgt", { spid: "#" }, []) || id === undefined || videos.has(id) || videoPauses?.get(id)?.status !== "mapped" && !videoObjects?.has(id)) return undefined;
    if (!videoObjects?.has(id)) videos.add(id);
  }
  const sequences = select(timeline, "seq").filter(seq => { const node = one(seq, "cTn"); return node && attribute(node, "nodeType") === "mainSeq"; });
  if (sequences.length !== 1) return undefined;
  const seq = sequences[0]!, main = one(seq, "cTn"), mainId = main && canonical(attribute(main, "id"));
  if (!exact(seq, "seq", { concurrent: "1", nextAc: "seek" }, ["cTn", "prevCondLst", "nextCondLst"]) || !timeNode(main, { dur: "indefinite", nodeType: "mainSeq" }, ["childTnLst"]) || mainId === undefined) return undefined;
  const slideCond = (container: XmlElement | undefined, name: string, event: string) => {
    if (!exact(container, name, {}, ["cond"])) return false;
    const cond = one(container, "cond"), target = one(cond, "tgtEl"), slide = one(target, "sldTgt");
    return exact(cond, "cond", { evt: event, delay: "0" }, ["tgtEl"]) && exact(target, "tgtEl", {}, ["sldTgt"]) && exact(slide, "sldTgt", {}, []);
  };
  if (!slideCond(one(seq, "prevCondLst"), "prevCondLst", "onPrev") || !slideCond(one(seq, "nextCondLst"), "nextCondLst", "onNext")) return undefined;
  const groups = one(main, "childTnLst"), pars = children(groups!); if (!pars.length || !exact(groups, "childTnLst", {}, pars.map(() => "par"))) return undefined;
  const result: MainAudioEffect[] = [];
  const simpleDelay = (list: XmlElement | undefined, expected: number) => exact(list, "stCondLst", {}, ["cond"]) && exact(one(list, "cond"), "cond", { delay: String(expected) }, []);
  for (let groupIndex = 0; groupIndex < pars.length; groupIndex++) {
    await checkpoint(); const par = pars[groupIndex]!, group = one(par, "cTn");
    if (!exact(par, "par", {}, ["cTn"]) || !timeNode(group, { fill: "hold" }, ["stCondLst", "childTnLst"])) return undefined;
    const conditions = one(group, "stCondLst"), conds = children(conditions!);
    if (!exact(conditions, "stCondLst", {}, conds.map(() => "cond")) || (conds.length !== 1 && conds.length !== 2) || !exact(conds[0], "cond", { delay: "indefinite" }, [])) return undefined;
    const begins = conds.length === 2;
    if (begins) { const cond = conds[1]!, ref = one(cond, "tn"); if (groupIndex !== 0 || !exact(cond, "cond", { evt: "onBegin", delay: "0" }, ["tn"]) || !exact(ref, "tn", { val: "#" }, []) || canonical(attribute(ref, "val")) !== mainId) return undefined; }
    const groupChildren = one(group, "childTnLst"), wrappers = children(groupChildren!); if (!wrappers.length || !exact(groupChildren, "childTnLst", {}, wrappers.map(() => "par"))) return undefined;
    let elapsed = 0;
    for (let index = 0; index < wrappers.length; index++) {
      await checkpoint(); const wrapper = wrappers[index]!, outer = one(wrapper, "cTn");
      if (!exact(wrapper, "par", {}, ["cTn"]) || !timeNode(outer, { fill: "hold" }, ["stCondLst", "childTnLst"]) || !simpleDelay(one(outer, "stCondLst"), index === 0 ? 0 : elapsed)) return undefined;
      const inner = one(one(outer, "childTnLst"), "par"), effect = one(inner, "cTn"), kind = index === 0 && !begins ? "clickEffect" : "afterEffect";
      if (!exact(one(outer, "childTnLst"), "childTnLst", {}, ["par"]) || !exact(inner, "par", {}, ["cTn"]) || !timeNode(effect, { presetID: "1", presetClass: "mediacall", presetSubtype: "0", fill: "hold", nodeType: kind }, ["stCondLst", "childTnLst"]) || !simpleDelay(one(effect, "stCondLst"), 0)) return undefined;
      const effects = one(effect, "childTnLst"), cmd = one(effects, "cmd"), bhvr = one(cmd, "cBhvr"), cmdTime = one(bhvr, "cTn"), targetEl = one(bhvr, "tgtEl"), shape = one(targetEl, "spTgt");
      if (!exact(effects, "childTnLst", {}, ["cmd"]) || !exact(cmd, "cmd", { type: "call", cmd: "playFrom(0.0)" }, ["cBhvr"]) || !exact(bhvr, "cBhvr", {}, ["cTn", "tgtEl"]) || !timeNode(cmdTime, { dur: "#", fill: "hold" }, []) || !exact(targetEl, "tgtEl", {}, ["spTgt"]) || !exact(shape, "spTgt", { spid: "#" }, [])) return undefined;
      const target = canonical(attribute(shape, "spid")), rawDuration = attribute(cmdTime!, "dur"), duration = Number(rawDuration);
      if (target === undefined || commands.get(target) !== (videos.has(target) ? 2 : 1) || rawDuration === undefined || !/^\d+$/u.test(rawDuration) || !Number.isSafeInteger(duration) || duration < 1 || duration > 2147483647 || elapsed > 2147483647) return undefined;
      result.push({ target, durationMs: duration, start: kind === "afterEffect" ? "automatic" : "click-sequence" }); elapsed += duration;
    }
  }
  if (videos.size && (result.filter(effect => videos.has(effect.target)).length !== videos.size)) return undefined;
  return result;
}
