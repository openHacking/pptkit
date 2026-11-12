import type { PresentationInheritanceSource } from "../types/inheritance.js";
import type { StyleOriginMap, StyleValueOrigin } from "./style-origins.js";

/** Declaration sidecar for raw fieldwise merges, before defaults are normalized. */
export class StyleInputTrace {
  constructor(
    readonly source: (property: string) => PresentationInheritanceSource,
    readonly origins: Record<string, StyleValueOrigin> = {},
    readonly prefix = "",
  ) {}

  scope(property: string): StyleInputTrace { return new StyleInputTrace(this.source, this.origins, this.key(property)); }
  private key(property: string): string { return this.prefix ? `${this.prefix}.${property}` : property; }

  record(property: string, sourceProperty = property, cleared = false): void {
    this.origins[this.key(property)] = { source: this.source(this.key(sourceProperty)), ...(cleared ? { state: "cleared" } : {}) };
  }

  /** A discriminated union replaces its old declaration members, including omissions. */
  replaceObject(property: string, input: object): void {
    const prefix = `${this.key(property)}.`;
    for (const key of Object.keys(this.origins)) if (key.startsWith(prefix)) delete this.origins[key];
    for (const [key, value] of Object.entries(input)) if (value !== undefined) this.record(`${property}.${key}`);
  }

  margin(input: number | Partial<Record<"left" | "right" | "top" | "bottom", number>>): void {
    for (const side of ["left", "right", "top", "bottom"] as const) {
      if (typeof input === "number") this.record(`margin.${side}`, "margin");
      else if (input[side] !== undefined) this.record(`margin.${side}`);
    }
  }
}

export function styleOriginsAt(origins: StyleOriginMap | undefined, prefix: string): StyleOriginMap {
  return Object.fromEntries(Object.entries(origins ?? {}).filter(([key]) => key.startsWith(prefix + ".")).map(([key, value]) => [key.slice(prefix.length + 1), value]));
}
