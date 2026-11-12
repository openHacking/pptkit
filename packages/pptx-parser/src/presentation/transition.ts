import type { PptxTransitionInfo } from "../types/structure.js";
import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children, integer, invalidStructure, singleChild } from "./structure-xml.js";

const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const XMLNS = "http://www.w3.org/2000/xmlns/";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
/** Interpret only the selected transition and its verified raw origin. */
export function readTransition(node: XmlElement, raw: XmlElement, p: string, part: string): PptxTransitionInfo {
  const bool = (name: string, fallback: boolean) => {
    const value = attribute(node, name)?.trim();
    if (value === undefined) return fallback;
    if (value === "true" || value === "1") return true;
    if (value === "false" || value === "0") return false;
    return invalidStructure(`Invalid transition ${name}.`, part);
  };
  const speed = attribute(node, "spd")?.trim() ?? "fast";
  if (!["fast", "med", "slow"].includes(speed)) invalidStructure("Invalid transition speed.", part);
  const advance = attribute(node, "advTm"), duration = attribute(raw, "dur", P14);
  const afterMs = advance === undefined ? undefined : integer(advance, 0, 4294967295, "transition/@advTm", part);
  const durationMs = duration === undefined ? undefined : integer(duration, 0, 4294967295, "transition/@p14:dur", part);
  const effects = children(raw), fade = singleChild(raw, p, "fade", part);
  const cleanAttributes = raw.attributes.every(a => a.namespaceUri === XMLNS || a.namespaceUri === MC || (!a.namespaceUri && ["spd", "advClick", "advTm"].includes(a.local)) || (a.namespaceUri === P14 && a.local === "dur"));
  const fadeValid = fade && !children(fade).length && fade.attributes.every(a => a.namespaceUri === XMLNS || (!a.namespaceUri && a.local === "thruBlk"));
  const black = fade ? attribute(fade, "thruBlk")?.trim() : undefined;
  if (black !== undefined && !["0", "1", "false", "true"].includes(black)) invalidStructure("Invalid fade thruBlk.", part);
  const plainText = [raw, ...(fade ? [fade] : [])].every(n => !n.children.some(c => (c.kind === "text" || c.kind === "cdata") && c.value.trim()));
  const side = effects.length === 1 && effects[0]!.namespaceUri === p && ["push", "wipe"].includes(effects[0]!.local) ? effects[0] : undefined;
  const directions = { l: "left", r: "right", u: "up", d: "down" } as const;
  const dir = side ? attribute(side, "dir")?.trim() : undefined;
  if (side && dir !== undefined && !Object.hasOwn(directions, dir)) invalidStructure("Invalid transition side direction.", part);
  const sideValid = side && dir !== undefined && !children(side).length && side.attributes.every(a => a.namespaceUri === XMLNS || (!a.namespaceUri && a.local === "dir")) && !side.children.some(c => (c.kind === "text" || c.kind === "cdata") && c.value.trim());
  const mappedFade = Boolean(plainText && cleanAttributes && effects.length === 1 && fadeValid && black !== "1" && black !== "true" && durationMs !== undefined);
  const mapped = mappedFade || Boolean(cleanAttributes && plainText && sideValid && durationMs !== undefined);
  return {
    status: mapped ? "mapped" : "source-only", speed: speed as "fast" | "med" | "slow", onClick: bool("advClick", true),
    ...(afterMs !== undefined ? { afterMs } : {}), ...(durationMs !== undefined ? { durationMs } : {}),
    ...(mapped ? mappedFade ? { type: "fade" as const } : { type: side!.local as "push" | "wipe", direction: directions[(dir ?? "l") as keyof typeof directions] } : { reason: "Mapping requires a plain fade/push/wipe with explicit p14:dur; speed-only, through-black, sounds, unknown effects/attributes and extensions remain source-only." }),
  };
}
