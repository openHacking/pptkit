import type { NormalizedPlaceholderDefinition } from "@pptkit/core";

/** Native indices are shared by layout prompts and their slide bindings. */
export function placeholderIndices(placeholders: readonly NormalizedPlaceholderDefinition[]): ReadonlyMap<string, number> {
  const used = new Set(placeholders.flatMap(placeholder => placeholder.index === undefined ? [] : [placeholder.index]));
  const result = new Map<string, number>();
  let nextFree = 1;
  for (const [position, placeholder] of placeholders.entries()) {
    let index = placeholder.index;
    if (index === undefined) {
      index = position + 1;
      if (used.has(index)) {
        while (used.has(nextFree)) nextFree++;
        if (nextFree > 4294967295) throw new RangeError("No native placeholder index remains available.");
        index = nextFree;
      }
      used.add(index);
    }
    result.set(placeholder.key, index);
  }
  return result;
}
