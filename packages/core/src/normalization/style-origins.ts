import type { PresentationInheritanceProperty, PresentationInheritanceSource, PresentationInheritanceTextTarget } from "../types/inheritance.js";
import type { OriginSink } from "./inheritance.js";

export interface StyleValueOrigin { source: PresentationInheritanceSource; state?: "cleared" }
export type StyleOriginMap = Readonly<Record<string, StyleValueOrigin>>;
export interface TextStyleOrigins { frame: StyleOriginMap; paragraph: StyleOriginMap; run: StyleOriginMap }

/** Explicit per-resolution state: the same selection returns the value and records its source. */
export class StyleTrace {
  readonly origins: Record<string, StyleValueOrigin> = {};
  constructor(readonly owner: PresentationInheritanceSource, readonly defaults: string, readonly fallback: StyleOriginMap = {}, readonly declarations: StyleOriginMap = {}) {}

  choose<T>(property: string, local: T | null | undefined, fallback: T, sourceProperty = property, fallbackOrigin?: StyleValueOrigin): T {
    const value = local ?? fallback;
    if (local != null) this.origins[property] = this.declarations[property] ?? { source: { ...this.owner, path: `${this.owner.path}.${sourceProperty}` } };
    else if (value !== undefined || this.fallback[property] !== undefined || fallbackOrigin !== undefined) this.origins[property] = fallbackOrigin ?? this.fallback[property] ?? this.builtin(property);
    return value;
  }

  clear(property: string, sourceProperty = property): void {
    this.origins[property] = this.declarations[property] ?? { source: { ...this.owner, path: `${this.owner.path}.${sourceProperty}` }, state: "cleared" };
  }

  builtin(property: string): StyleValueOrigin { return { source: { kind: "default", path: `${this.defaults}.${property}` } }; }

  emit(target: PresentationInheritanceTextTarget, prefix: string, sink: OriginSink): void {
    for (const [property, origin] of Object.entries(this.origins)) sink({ target, property: `${prefix}.${property}` as PresentationInheritanceProperty, ...origin });
  }
}

export function selectStyle<T>(trace: StyleTrace | undefined, property: string, local: T | null | undefined, fallback: T, sourceProperty = property, fallbackOrigin?: StyleValueOrigin): T {
  return trace ? trace.choose(property, local, fallback, sourceProperty, fallbackOrigin) : local ?? fallback;
}
