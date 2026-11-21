import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";

const MEDIA_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const MEDIA_EXT = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

type Trim = { status: "absent" | "unsupported" } | { status: "mapped"; trimStartMs: number; trimEndMs: number };

/** Only project integral millisecond trim metadata on one effective Office media node. */
export function audioTrim(nvPr: XmlElement | undefined, p: string, durationSeconds: number | undefined): Trim {
  if (!nvPr) return { status: "absent" };
  const media = children(nvPr).filter(x => x.namespaceUri === p && x.local === "extLst").flatMap(list => children(list).filter(x => x.namespaceUri === p && x.local === "ext" && attribute(x, "uri")?.toUpperCase() === MEDIA_EXT).flatMap(ext => children(ext).filter(x => x.namespaceUri === MEDIA_NS && x.local === "media")));
  if (!media.length) return { status: "absent" };
  if (media.length !== 1) return { status: "unsupported" };
  const trims = children(media[0]!).filter(x => x.namespaceUri === MEDIA_NS && x.local === "trim");
  if (!trims.length) return { status: "absent" };
  if (trims.length !== 1 || children(trims[0]!).length || trims[0]!.children.some(x => (x.kind === "text" || x.kind === "cdata") && x.value.trim() !== "") || trims[0]!.attributes.some(a => a.namespaceUri !== "" || !["st", "end"].includes(a.local))) return { status: "unsupported" };
  const parse = (value: string | undefined) => {
    const text = value?.trim() ?? "0";
    if (!/^\d+$/u.test(text)) return undefined;
    const n = Number(text); return Number.isSafeInteger(n) && n <= 2147483647 ? n : undefined;
  };
  const start = parse(attribute(trims[0]!, "st")), end = parse(attribute(trims[0]!, "end"));
  const duration = durationSeconds === undefined ? NaN : Math.round(durationSeconds * 1000);
  if (start === undefined || end === undefined || !Number.isSafeInteger(duration) || start + end >= duration) return { status: "unsupported" };
  return { status: "mapped", trimStartMs: start, trimEndMs: end };
}
