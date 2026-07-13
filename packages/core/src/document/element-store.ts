import type { GroupElement, PresentationElement, PresentationElementInput } from "../types/element.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import { IdAllocator } from "../utils/id.js";

export class ElementStore {
  private readonly ids = new Set<string>();
  private readonly allocator = new IdAllocator("element");

  materialize(input: PresentationElementInput): PresentationElement {
    const pending = new Set<string>();
    const element = this.materializeNested(input, pending);
    for (const id of pending) this.ids.add(id);
    return element;
  }

  release(element: PresentationElement): void {
    this.ids.delete(element.id);
    if (element.type === "group") {
      for (const child of element.children) this.release(child);
    }
  }

  withoutIds(element: PresentationElement): PresentationElementInput {
    const clone = deepClone(element) as PresentationElementInput;
    delete clone.id;
    if (clone.type === "group") {
      const source = element as GroupElement;
      clone.children = source.children.map((child) => this.withoutIds(child));
    }
    return clone;
  }

  private materializeNested(input: PresentationElementInput, pending: Set<string>): PresentationElement {
    const id = input.id ?? this.allocator.next((candidate) => this.ids.has(candidate) || pending.has(candidate));
    if (this.ids.has(id) || pending.has(id)) {
      throw new Error(`Duplicate element id "${id}" detected.`);
    }
    pending.add(id);

    if (input.type === "group") {
      return deepFreeze({
        ...deepClone(input),
        id,
        children: input.children.map((child) => this.materializeNested(child, pending)),
      }) as GroupElement;
    }
    return deepFreeze({ ...deepClone(input), id }) as PresentationElement;
  }
}
