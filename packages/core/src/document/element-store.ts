import { remapCopiedReferences } from "./copy-references.js";
import type { GroupElement, PresentationElement, PresentationElementInput } from "../types/element.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import { IdAllocator } from "../utils/id.js";

export class ElementStore {
  private ids = new Set<string>();
  private readonly registeredRoots = new WeakSet<PresentationElement>();
  private allocator = new IdAllocator("element");
  private version = 0;
  private appendOrigin: { store: ElementStore; version: number } | undefined;
  private appendBase: ReadonlySet<string> | undefined;

  /** Create a private append-only workspace sharing no mutable allocator or identities. */
  forkForAppend(): ElementStore {
    const staged = new ElementStore();
    staged.appendBase = this.ids;
    staged.allocator = this.allocator.fork();
    staged.appendOrigin = { store: this, version: this.version };
    return staged;
  }

  /** Adopt only the specified new roots after verifying the original has not changed. */
  commitAppend(staged: ElementStore, roots: readonly PresentationElement[]): void {
    if (staged.appendOrigin?.store !== this || staged.appendOrigin.version !== this.version) throw new Error("Element store changed during append transaction.");
    const expected = new Set<string>(), stack = [...roots];
    for (const root of roots) if (!staged.registeredRoots.has(root)) throw new Error("Append transaction contains an unregistered root.");
    while (stack.length) {
      const element = stack.pop()!;
      if (this.ids.has(element.id) || expected.has(element.id)) throw new Error("Append transaction contains a duplicate identity.");
      expected.add(element.id);
      if (element.type === "group") for (const child of element.children) stack.push(child);
    }
    if (expected.size !== staged.ids.size || [...expected].some(id => !staged.ids.has(id))) throw new Error("Append transaction contains unaccounted identities.");
    for (const id of expected) this.ids.add(id);
    this.allocator = staged.allocator;
    for (const root of roots) this.registeredRoots.add(root);
    this.version++;
    staged.appendOrigin = undefined;
  }

  materialize(input: PresentationElementInput, validate?: (element: PresentationElement) => void): PresentationElement {
    const pending = new Set<string>();
    const element = this.materializeNested(input, pending);
    validate?.(element);
    for (const id of pending) this.ids.add(id);
    this.registeredRoots.add(element);
    this.version++;
    return element;
  }

  release(element: PresentationElement): void {
    // Released snapshots must never unregister IDs subsequently reused elsewhere.
    if (!this.registeredRoots.delete(element)) return;
    this.version++;
    const pending = [element];
    while (pending.length) {
      const current = pending.pop()!; this.ids.delete(current.id);
      if (current.type === "group") for (const child of current.children) pending.push(child);
    }
  }

  copyInputs(elements: readonly PresentationElement[], slideIds: ReadonlyMap<string, string> = new Map()): PresentationElementInput[] {
    return this.copyWithIdentityMap(elements, slideIds).elements;
  }

  copyWithIdentityMap(elements: readonly PresentationElement[], slideIds: ReadonlyMap<string, string> = new Map()): { elements: PresentationElementInput[]; elementIds: ReadonlyMap<string, string> } {
    const copies = deepClone(elements) as PresentationElementInput[];
    const allocator = this.allocator.fork();
    const elementIds = new Map<string, string>(), pending = new Set<string>();
    const stack = [...copies].reverse();
    while (stack.length) {
      const element = stack.pop()!, sourceId = element.id!;
      const id = allocator.next(candidate => this.ids.has(candidate) || this.appendBase?.has(candidate) === true || pending.has(candidate));
      pending.add(id); elementIds.set(sourceId, id); element.id = id;
      if (element.type === "group") {
        for (let index = element.children.length - 1; index >= 0; index--) stack.push(element.children[index]!);
      }
    }
    // Allocate every identity before resolving forward or sibling references.
    remapCopiedReferences(copies, elementIds, slideIds);
    this.allocator = allocator;
    this.version++;
    return { elements: copies, elementIds };
  }

  replace(original: PresentationElement, input: PresentationElementInput, validate: (element: PresentationElement) => void): PresentationElement {
    if (input.id !== undefined && input.id !== original.id) throw new Error("replaceElement must retain the root element ID.");
    return this.replaceMany([original], [{ ...input, id: original.id }], elements => validate(elements[0]!))[0]!;
  }

  /** Stage an entire owned collection before changing any registered identities. */
  replaceMany(originals: readonly PresentationElement[], inputs: readonly PresentationElementInput[], validate: (elements: PresentationElement[]) => void): PresentationElement[] {
    for (const original of originals) if (!this.registeredRoots.has(original)) throw new Error("Cannot replace a released element snapshot.");
    const excluded = new Set<string>(), stack = [...originals];
    while (stack.length) { const element = stack.pop()!; excluded.add(element.id); if (element.type === "group") for (const child of element.children) stack.push(child); }
    const allocator = this.allocator.fork();
    const pending = new Set<string>(), elements = inputs.map(input => this.materializeNested(input, pending, excluded, allocator));
    validate(elements);
    for (const id of excluded) this.ids.delete(id);
    for (const id of pending) this.ids.add(id);
    for (const original of originals) this.registeredRoots.delete(original);
    for (const element of elements) this.registeredRoots.add(element);
    this.allocator = allocator;
    this.version++;
    return elements;
  }

  private materializeNested(input: PresentationElementInput, pending: Set<string>, excluded: ReadonlySet<string> = new Set(), allocator = this.allocator): PresentationElement {
    const id = input.id ?? allocator.next((candidate) => this.ids.has(candidate) || this.appendBase?.has(candidate) === true || pending.has(candidate));
    if (((this.ids.has(id) || this.appendBase?.has(id)) && !excluded.has(id)) || pending.has(id)) {
      throw new Error(`Duplicate element id "${id}" detected.`);
    }
    pending.add(id);

    if (input.type === "group") {
      return deepFreeze({
        ...deepClone(input),
        id,
        children: input.children.map((child) => this.materializeNested(child, pending, excluded, allocator)),
      }) as GroupElement;
    }
    return deepFreeze({ ...deepClone(input), id }) as PresentationElement;
  }
}
