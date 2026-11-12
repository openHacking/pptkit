import type { ElementAction, PresentationElementInput, TextContentInput } from "../types/element.js";

/** Apply a collector-produced path only to the detached copy. */
export function setCopiedReference(root: object, path: string, value: string): void {
  const parts = path.split("."), last = parts.pop();
  let owner: unknown = root;
  for (const part of parts) {
    if (!owner || typeof owner !== "object" || !Object.hasOwn(owner, part)) throw new TypeError("Invalid copied reference path.");
    owner = (owner as Record<string, unknown>)[part];
  }
  if (!owner || typeof owner !== "object" || last === undefined || !Object.hasOwn(owner, last)) throw new TypeError("Invalid copied reference path.");
  (owner as Record<string, unknown>)[last] = value;
}

function remapAction(action: ElementAction | undefined, slides: ReadonlyMap<string, string>): void {
  if (action?.type === "slide") action.slideId = slides.get(action.slideId) ?? action.slideId;
}

/** Mutates only detached copies and explicitly modeled action/reference fields. */
export function remapTextActions(content: TextContentInput | undefined, slides: ReadonlyMap<string, string>): void {
  if (typeof content === "string" || content === undefined) return;
  for (const paragraph of content) for (const run of paragraph.runs) remapAction(run.action, slides);
}

export function remapCopiedReferences(elements: readonly PresentationElementInput[], elementIds: ReadonlyMap<string, string>, slideIds: ReadonlyMap<string, string>): void {
  const stack = [...elements].reverse();
  while (stack.length) {
    const element = stack.pop()!;
    remapAction(element.action, slideIds);
    if (element.type === "connector") {
      for (const endpoint of [element.start, element.end]) {
        if ("elementId" in endpoint) endpoint.elementId = elementIds.get(endpoint.elementId) ?? endpoint.elementId;
      }
    } else if (element.type === "text") remapTextActions(element.content, slideIds);
    else if (element.type === "shape") remapTextActions(element.text?.content, slideIds);
    else if (element.type === "table") {
      for (const row of element.rows) for (const cell of row.cells) remapTextActions(cell.content, slideIds);
    } else if (element.type === "group") {
      for (let index = element.children.length - 1; index >= 0; index--) stack.push(element.children[index]!);
    }
  }
}
