/** Validate the authored sequence and its exact ordering after PPTX EMU rounding. */
export function tabStopErrors(value: unknown): { index?: number; message: string }[] {
  if (!Array.isArray(value)) return [{ message: "Tab stops must be an array." }];
  const errors: { index?: number; message: string }[] = [];
  let previous = -1;
  for (const [index, stop] of value.entries()) {
    if (!stop || typeof stop !== "object" || Array.isArray(stop) || ![Object.prototype, null].includes(Object.getPrototypeOf(stop)) || Object.keys(stop).some(key => key !== "position" && key !== "align")) {
      errors.push({ index, message: "Each tab stop must contain only position and alignment." }); continue;
    }
    const { position, align } = stop as Record<string, unknown>;
    const emu = typeof position === "number" ? Math.round(position * 12700) : NaN;
    if (typeof position !== "number" || !Number.isFinite(position) || position < 0 || !Number.isSafeInteger(emu) || emu > 2147483647 || emu <= previous) {
      errors.push({ index, message: "Tab stop positions must be nonnegative, strictly increasing after EMU rounding and within the DrawingML coordinate range." });
    } else previous = emu;
    if (align !== "left" && align !== "center" && align !== "right" && align !== "decimal") errors.push({ index, message: "Tab alignment must be left, center, right or decimal." });
  }
  return errors;
}
