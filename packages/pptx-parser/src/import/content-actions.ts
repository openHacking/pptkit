import type { ElementAction } from "@pptkit/core";
import type { structureContext } from "../presentation/structure-context.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";

export interface ContentAction {
  node: XmlElement;
  relationshipId?: string;
  action?: ElementAction;
  reason: string;
}

/** Click semantics only; never resolves/opens an external resource or executes an action. */
export function contentAction(owner: XmlElement, r: ContentReader, context: Awaited<ReturnType<typeof structureContext>>, slideIds: ReadonlyMap<string, string>): ContentAction | undefined {
  const node = r.one(owner, r.a, "hlinkClick");
  if (!node) return undefined;
  r.known(node);
  const relationshipId = r.attr(node, "id", context.r), command = r.attr(node, "action"), tooltip = r.attr(node, "tooltip");
  const base = { node, ...(relationshipId !== undefined ? { relationshipId } : {}) };
  const missing = (reason: string): ContentAction => ({ ...base, reason });
  const relationship = context.findRelationship(r.part, relationshipId);
  if (!relationship) return missing(`Click action has no resolvable relationship (${relationshipId ?? "absent"}); relative slideshow, macro/program and relationship-free actions are not synthesized.`);
  const tip = tooltip !== undefined ? { tooltip } : {};
  if (relationship.type === `${context.r}/slide`) {
    if (command !== undefined && command !== "" && command !== "ppaction://hlinksldjump") return missing("This slide click command is outside the explicit slide-jump subset.");
    const slideId = relationship.targetMode === "Internal" && relationship.fragment === undefined && relationship.targetPart ? slideIds.get(relationship.targetPart) : undefined;
    if (!slideId) return missing("Slide click target is external, fragmented or outside the effective presentation slide list.");
    return { ...base, action: { type: "slide", slideId, ...tip }, reason: "Explicit slide relationship maps to the target Core slide identity; source ordering does not determine the link." };
  }
  if (relationship.type !== `${context.r}/hyperlink` || relationship.targetMode !== "External") return missing("Click relationship is not an external hyperlink or a supported internal slide relationship.");
  if (command !== undefined && command !== "") return missing("Hyperlink commands, including file/program/media actions, are retained in source only.");
  // Avoid silently making relative/file/program URIs actionable in a consumer.
  // Keep the exact accepted source spelling, including percent escapes and query strings.
  try {
    const url = new URL(relationship.target);
    if (!["http:", "https:", "mailto:"].includes(url.protocol)) return missing("Hyperlink scheme is outside the HTTP/HTTPS/mailto subset and is not projected as a live action.");
  } catch { return missing("Relative or malformed hyperlink target has no current Core action mapping."); }
  if (/[\u0000-\u0020\u007f]/.test(relationship.target)) return missing("Hyperlink with whitespace/control characters is not projected as a live action.");
  return { ...base, action: { type: "url", url: relationship.target, ...tip }, reason: "Explicit external HTTP/HTTPS/mailto hyperlink and tooltip map without fetching or normalizing the target." };
}
