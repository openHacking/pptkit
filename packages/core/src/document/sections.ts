import type { PresentationSection } from "../types/presentation.js";

function denseArray(value: unknown): value is unknown[] {
  if (!Array.isArray(value) || value.length > 100000 || Reflect.ownKeys(value).length !== value.length + 1) return false;
  for (let index = 0; index < value.length; index++) if (!Object.hasOwn(value, index)) return false;
  return true;
}

/** Capture explicit identities and membership without relying on names or slide indices. */
export function captureSections(input: readonly PresentationSection[] | undefined, slideIds: ReadonlySet<string>): readonly PresentationSection[] | undefined {
  if (input === undefined) return undefined;
  if (!denseArray(input)) throw new TypeError("Sections must be a dense bounded array.");
  const identities = new Set<string>(), members = new Set<string>();
  const sections = Array.from(input, item => {
    if (!item || typeof item !== "object" || Array.isArray(item) || ![Object.prototype, null].includes(Object.getPrototypeOf(item)) || Reflect.ownKeys(item).some(key => typeof key !== "string" || !["id", "name", "slideIds"].includes(key) || !Object.prototype.propertyIsEnumerable.call(item, key))) throw new TypeError("Invalid section record.");
    const { id, name, slideIds: ids } = item as PresentationSection;
    if (typeof id !== "string" || !id || id.length > 1024 || identities.has(id) || (name !== undefined && (typeof name !== "string" || name.length > 10000)) || !denseArray(ids)) throw new TypeError("Invalid section identity, name or members.");
    identities.add(id);
    const captured = Array.from(ids, slideId => {
      if (typeof slideId !== "string" || !slideIds.has(slideId) || members.has(slideId)) throw new TypeError("Section member must identify an existing, uniquely assigned slide.");
      members.add(slideId); return slideId;
    });
    return Object.freeze({ id, ...(name === undefined ? {} : { name }), slideIds: Object.freeze(captured) });
  });
  return Object.freeze(sections);
}
