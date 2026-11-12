import type { PresentationInheritanceSource, PresentationInheritanceTextTarget } from "../types/inheritance.js";
import type { OriginSink } from "./inheritance.js";
import { StyleTrace, type TextStyleOrigins, type StyleOriginMap } from "./style-origins.js";

export interface TextOriginContext {
  owner: PresentationInheritanceSource;
  target: PresentationInheritanceTextTarget;
  sink: OriginSink;
  prefix: string;
  declarations?: readonly { style: StyleOriginMap; runs: readonly StyleOriginMap[] }[] | undefined;
  fallback: TextStyleOrigins | undefined;
}

export interface ElementOriginContext {
  path: string;
  sink: OriginSink;
  placeholders: ReadonlyMap<string, TextStyleOrigins>;
}

export function childSource(owner: PresentationInheritanceSource, property: string): PresentationInheritanceSource {
  return { ...owner, path: `${owner.path}.${property}` };
}

export function presetSource(name: string | undefined, fallback: PresentationInheritanceSource): PresentationInheritanceSource {
  return name === undefined ? fallback : { kind: "preset", id: name, path: `textStylePresets[${JSON.stringify(name)}]` };
}

export function traceStyles(owner: PresentationInheritanceSource | undefined, fallback?: TextStyleOrigins): { frame?: StyleTrace; paragraph?: StyleTrace; run?: StyleTrace } {
  if (!owner) return {};
  return {
    frame: new StyleTrace(childSource(owner, "frame"), "defaults.text.frame", fallback?.frame),
    paragraph: new StyleTrace(childSource(owner, "paragraph"), "defaults.text.paragraph", fallback?.paragraph),
    run: new StyleTrace(childSource(owner, "run"), "defaults.text.run", fallback?.run),
  };
}

export function styleOrigins(traces: ReturnType<typeof traceStyles>): TextStyleOrigins | undefined {
  return traces.frame && traces.paragraph && traces.run ? { frame: traces.frame.origins, paragraph: traces.paragraph.origins, run: traces.run.origins } : undefined;
}
