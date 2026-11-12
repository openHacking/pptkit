/** New shared-definition edits accept finite plain data, never executable accessors. */
export function captureDefinitionInput<T>(input: T, allowed: readonly string[]): T {
  let count = 0;
  const allowedKeys = new Set(allowed);
  const ancestors = new Set<object>();
  function capture(value: unknown, depth: number): unknown {
    if (++count > 1000000 || depth > 128) throw new RangeError("Definition data exceeds 1000000 values or depth 128.");
    if (value === undefined || value === null || typeof value === "string" || typeof value === "boolean") return value;
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (typeof value !== "object" || (!Array.isArray(value) && ![Object.prototype, null].includes(Object.getPrototypeOf(value)))) throw new TypeError("Definitions require finite plain data.");
    if (ancestors.has(value)) throw new TypeError("Cyclic definition data is unsupported.");
    const array = Array.isArray(value), keys = Reflect.ownKeys(value);
    if (keys.length > 1000000 - count || (array && keys.length !== value.length + 1)) throw new TypeError("Definition arrays must be dense and data must fit the value limit.");
    const result: Record<string, unknown> = array ? [] as unknown as Record<string, unknown> : {};
    ancestors.add(value);
    for (const key of keys) {
      if (array && key === "length") continue;
      if (typeof key !== "string" || (array && (!Number.isSafeInteger(Number(key)) || String(Number(key)) !== key || Number(key) < 0 || Number(key) >= value.length))) throw new TypeError("Definition fields must be string keys or dense array indices.");
      const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
      if (!("value" in descriptor) || !descriptor.enumerable) throw new TypeError("Definition fields must be enumerable data properties; accessors are unsupported.");
      Object.defineProperty(result, key, { value: capture(descriptor.value, depth + 1), enumerable: true, writable: true, configurable: true });
    }
    ancestors.delete(value);
    return result;
  }
  const result = capture(input, 0) as T;
  if (!result || typeof result !== "object" || Array.isArray(result) || Object.keys(result).some(key => !allowedKeys.has(key))) throw new TypeError("Unsupported definition fields.");
  return result;
}

export function assertDefinitionIdentity(id: unknown, name: unknown): void {
  if (typeof id !== "string" || !id.trim() || typeof name !== "string" || !name.trim()) throw new TypeError("Definition id and name must be nonempty strings.");
}

export function assertDefinitionElements(elements: readonly unknown[]): void {
  const stack = [...elements];
  while (stack.length) {
    const element = stack.pop() as Record<string, unknown> | null;
    if (!element || typeof element !== "object" || !["text", "image", "shape", "connector", "group", "table", "chart"].includes(element.type as string)) throw new TypeError("Layout elements must use a supported element type.");
    if (element.type === "group") {
      if (!Array.isArray(element.children)) throw new TypeError("Group children must be an array.");
      for (const child of element.children) stack.push(child);
    }
  }
}
