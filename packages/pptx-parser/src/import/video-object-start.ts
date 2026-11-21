import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";

/** Validate direct video media nodes before interpreting an independent toggle as a start. */
export async function videoObjectStart(root: XmlElement, p: string, checkpoint: () => Promise<void>, interactions: ReadonlyMap<string, { status: "mapped"; durationMs: number } | { status: "unsupported" }>): Promise<ReadonlySet<string>> {
  const select = (node: XmlElement | undefined, name: string) => node ? children(node).filter(x => x.namespaceUri === p && x.local === name) : [];
  const one = (node: XmlElement | undefined, name: string) => { const items = select(node, name); return items.length === 1 ? items[0] : undefined; };
  const number = (value: string | undefined) => value !== undefined && /^\d+$/u.test(value) && Number.isSafeInteger(Number(value)) && Number(value) <= 4294967295 ? String(Number(value)) : undefined;
  const exact = (node: XmlElement | undefined, name: string, attrs: Record<string, string>, kids: readonly string[]) => Boolean(node && node.namespaceUri === p && node.local === name && !node.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim()) && node.attributes.length === Object.keys(attrs).length && node.attributes.every(a => a.namespaceUri === "" && Object.hasOwn(attrs, a.local) && (attrs[a.local] === "#" || attrs[a.local] === a.value)) && children(node).length === kids.length && children(node).every((x, i) => x.namespaceUri === p && x.local === kids[i]));
  const result = new Set<string>(), timings = select(root, "timing"); if (timings.length !== 1) return result;
  const ids = new Map<string, number>(), pending = [timings[0]!];
  while (pending.length) {
    await checkpoint(); const node = pending.pop()!;
    if (node.namespaceUri === p && node.local === "cTn") { const id = number(attribute(node, "id")); if (id !== undefined) ids.set(id, (ids.get(id) ?? 0) + 1); }
    for (const child of children(node)) pending.push(child);
  }
  const list = one(timings[0], "tnLst"), par = one(list, "par"), time = one(par, "cTn"), timeline = one(time, "childTnLst");
  const rootId = time && number(attribute(time, "id"));
  if (!exact(timings[0], "timing", {}, ["tnLst"]) || !exact(list, "tnLst", {}, ["par"]) || !exact(par, "par", {}, ["cTn"]) || !exact(time, "cTn", { id: "#", dur: "indefinite", restart: "never", nodeType: "tmRoot" }, ["childTnLst"]) || rootId === undefined || ids.get(rootId) !== 1 || !timeline || timeline.attributes.length || children(timeline).some(x => x.namespaceUri !== p || !["seq", "audio", "video"].includes(x.local))) return result;
  const seen = new Set<string>();
  for (const video of select(timeline, "video")) {
    await checkpoint();
    const media = one(video, "cMediaNode"), t = one(media, "cTn"), start = one(t, "stCondLst"), target = one(media, "tgtEl"), shape = one(target, "spTgt");
    const id = shape && number(attribute(shape, "spid")), timeId = t && number(attribute(t, "id"));
    if (id === undefined) continue;
    if (seen.has(id)) { result.delete(id); continue; } seen.add(id);
    if (exact(video, "video", {}, ["cMediaNode"]) && exact(media, "cMediaNode", { vol: "80000" }, ["cTn", "tgtEl"]) && exact(t, "cTn", { id: "#", fill: "hold", display: "0" }, ["stCondLst"]) && timeId !== undefined && ids.get(timeId) === 1 && exact(start, "stCondLst", {}, ["cond"]) && exact(one(start, "cond"), "cond", { delay: "indefinite" }, []) && exact(target, "tgtEl", {}, ["spTgt"]) && exact(shape, "spTgt", { spid: "#" }, []) && interactions.get(id)?.status === "mapped") result.add(id);
  }
  return result;
}
