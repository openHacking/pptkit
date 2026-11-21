import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";

type Result = { status: "mapped"; durationMs: number } | { status: "unsupported" };
type Spec = { name: string; attrs?: Record<string, string>; kids?: Spec[] };
const el = (name: string, attrs: Record<string, string> = {}, kids: Spec[] = []): Spec => ({ name, attrs, kids });
const tn = (attrs: Record<string, string>, kids: Spec[]) => el("cTn", { id: "#id", ...attrs }, kids);
const shapeTarget = () => el("tgtEl", {}, [el("spTgt", { spid: "#target" })]);
const cond = () => el("cond", { evt: "onClick", delay: "0" }, [shapeTarget()]);
const delay = () => el("stCondLst", {}, [el("cond", { delay: "0" })]);

/** Recognize an observed independent audio play or video pause click sequence. */
export async function audioObjectStart(root: XmlElement, p: string, checkpoint: () => Promise<void>, kind: "audio" | "video-pause" = "audio"): Promise<ReadonlyMap<string, Result>> {
  const selected = (node: XmlElement, local: string) => children(node).filter(x => x.namespaceUri === p && x.local === local);
  const one = (node: XmlElement | undefined, local: string) => { const list = node ? selected(node, local) : []; return list.length === 1 ? list[0] : undefined; };
  const canonical = (value: string | undefined) => {
    if (value === undefined || !/^\d+$/u.test(value)) return undefined;
    const n = Number(value); return Number.isSafeInteger(n) && n <= 4294967295 ? String(n) : undefined;
  };
  const result = new Map<string, Result>(), timings = selected(root, "timing");
  if (timings.length !== 1) return result;
  const ids = new Map<string, number>(), commands = new Map<string, number>(), pending = [timings[0]!];
  while (pending.length) {
    await checkpoint(); const node = pending.pop()!;
    if (node.namespaceUri === p && node.local === "cTn") { const id = canonical(attribute(node, "id")); if (id !== undefined) ids.set(id, (ids.get(id) ?? 0) + 1); }
    if (node.namespaceUri === p && node.local === "cmd") {
      const bhvr = one(node, "cBhvr"), target = one(one(bhvr, "tgtEl"), "spTgt"), id = target && canonical(attribute(target, "spid"));
      if (id !== undefined) commands.set(id, (commands.get(id) ?? 0) + 1);
    }
    for (const child of children(node)) pending.push(child);
  }
  const list = one(timings[0], "tnLst"), par = one(list, "par"), top = one(par, "cTn");
  if (!top || attribute(top, "nodeType") !== "tmRoot") return result;
  const timeline = one(top, "childTnLst"); if (!timeline) return result;
  for (const seq of selected(timeline, "seq")) {
    await checkpoint();
    const seqTn = one(seq, "cTn"); if (!seqTn || attribute(seqTn, "nodeType") !== "interactiveSeq") continue;
    const start = one(one(seqTn, "stCondLst"), "cond"), shape = one(one(start, "tgtEl"), "spTgt");
    const target = shape && canonical(attribute(shape, "spid")); if (target === undefined) continue;
    if (result.has(target)) { result.set(target, { status: "unsupported" }); continue; }
    let duration = 0;
    const matches = (node: XmlElement, spec: Spec): boolean => {
      if (node.namespaceUri !== p || node.local !== spec.name || node.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim() !== "")) return false;
      const attrs = spec.attrs ?? {}, actual = node.attributes;
      if (actual.length !== Object.keys(attrs).length || actual.some(a => a.namespaceUri !== "" || !(a.local in attrs))) return false;
      for (const [key, expected] of Object.entries(attrs)) {
        const value = attribute(node, key);
        if (expected === "#id") { const id = canonical(value); if (id === undefined || ids.get(id) !== 1) return false; }
        else if (expected === "#target") { if (canonical(value) !== target) return false; }
        else if (expected === "#duration") { if (value === undefined || !/^\d+$/u.test(value)) return false; const n = Number(value); if (!Number.isSafeInteger(n) || n < 1 || n > 2147483647) return false; duration = n; }
        else if (value !== expected) return false;
      }
      const actualKids = children(node), expectedKids = spec.kids ?? [];
      return actualKids.length === expectedKids.length && actualKids.every((child, i) => matches(child, expectedKids[i]!));
    };
    const command = el("cmd", { type: "call", cmd: kind === "audio" ? "playFrom(0.0)" : "togglePause" }, [el("cBhvr", {}, [tn({ dur: "#duration", fill: "hold" }, []), shapeTarget()])]);
    const effect = tn({ presetID: kind === "audio" ? "1" : "2", presetClass: "mediacall", presetSubtype: "0", fill: "hold", nodeType: "clickEffect" }, [delay(), el("childTnLst", {}, [command])]);
    const layer = (child: Spec) => el("par", {}, [tn({ fill: "hold" }, [delay(), el("childTnLst", {}, [child])])]);
    const body = layer(layer(el("par", {}, [effect])));
    const expected = el("seq", { concurrent: "1", nextAc: "seek" }, [tn({ restart: "whenNotActive", fill: "hold", evtFilter: "cancelBubble", nodeType: "interactiveSeq" }, [el("stCondLst", {}, [cond()]), el("endSync", { evt: "end", delay: "0" }, [el("rtn", { val: "all" })]), el("childTnLst", {}, [body])]), el("nextCondLst", {}, [cond()])]);
    result.set(target, matches(seq, expected) && commands.get(target) === (kind === "audio" ? 1 : 2) && (kind === "audio" || duration === 1) ? { status: "mapped", durationMs: duration } : { status: "unsupported" });
  }
  return result;
}
