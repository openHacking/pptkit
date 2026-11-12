import type { BackgroundPaintInput, PaintInput, ShapeFillPaintInput, ShapeStyleInput } from "@pptkit/core";
import { children, invalidStructure } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { PptxSourceLocation } from "../types/import.js";
import type { ContentReader } from "./content-reader.js";
import type { ThemeTextResolution } from "./theme-text-resolver.js";
import { transformTextColor } from "./color-transforms.js";

/** Bounded explicit fill reader shared by shared and slide-local artwork. */
export function artworkPaintReader(r: ContentReader, checkpoint: () => Promise<void>, resolveColor: (node: XmlElement, source: PptxSourceLocation) => Promise<ThemeTextResolution>, consume: (count: number, source: PptxSourceLocation) => void) {
  const colorValue = async (color: XmlElement, source: PptxSourceLocation, dependencies?: PptxSourceLocation[], locate?: (node: XmlElement) => PptxSourceLocation | undefined): Promise<{ color: string; opacity: number } | undefined> => {
      const modifiers = children(color); let result;
      if (color.local === "schemeClr") {
        r.attr(color, "val"); result = await resolveColor(color, source);
      } else if (color.local === "srgbClr" || color.local === "sysClr") {
        const key = color.local === "srgbClr" ? "val" : "lastClr", value = r.attr(color, key);
        if (color.local === "sysClr") r.attr(color, "val");
        if (!value || !/^[\da-f]{6}$/i.test(value)) { r.note(color, "Color requires a valid explicit RGB value or system fallback."); return undefined; }
        result = transformTextColor(value, modifiers, r.a);
      } else { r.note(color, "Unsupported solid color model."); return undefined; }
      const count = "resolvedFrom" in result ? result.resolvedFrom?.length ?? modifiers.length : modifiers.length;
      consume(count, source);
      if (result.status !== "mapped") { r.note(color, result.reason); return undefined; }
      if (dependencies) {
        if ("resolvedFrom" in result && result.resolvedFrom) dependencies.push(...result.resolvedFrom);
        else for (const modifier of modifiers) { const location = locate?.(modifier); if (location) dependencies.push(location); }
      }
      r.known(color); for (const modifier of modifiers) { await checkpoint(); r.known(modifier); r.attr(modifier, "val"); }
      return { color: result.value, opacity: result.opacity ?? 1 };
    };
  const paint = async (parent: XmlElement, source: PptxSourceLocation): Promise<PaintInput | undefined> => {
      await checkpoint();
      const fills = children(parent).filter(n => n.namespaceUri === r.a && ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(n.local));
      if (fills.length > 1) invalidStructure("Conflicting drawing fill choices.", r.part);
      const fill = fills[0]; if (!fill) return undefined; r.known(fill);
      if (fill.local === "noFill") return { type: "none" };
      if (fill.local !== "solidFill") { r.note(fill, "Only solid and no-fill paints are projected."); return undefined; }
      const choices = children(fill); if (choices.length !== 1 || choices[0]!.namespaceUri !== r.a) { r.note(fill, "Solid fill needs one supported color."); return undefined; }
      const color = await colorValue(choices[0]!, source);
      return color ? { type: "solid", ...color } : undefined;
    };
  const gradient = async (parent: XmlElement, source: PptxSourceLocation, locate?: (node: XmlElement) => PptxSourceLocation | undefined, onResolved?: (locations: readonly PptxSourceLocation[]) => void): Promise<BackgroundPaintInput | undefined> => {
      const fills = children(parent).filter(n => n.namespaceUri === r.a && ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(n.local));
      if (fills.length > 1) invalidStructure("Conflicting drawing fill choices.", r.part);
      const node = r.one(parent, r.a, "gradFill"); if (!node) return undefined; r.known(node);
      const fail = (reason: string) => { r.note(node, reason); return undefined; };
      const flip = r.attr(node, "flip"), rotation = r.attr(node, "rotWithShape"); if (flip !== undefined && flip !== "none") return fail("Gradient flip modes are unsupported.");
      if (flip === undefined || rotation === undefined) r.note(node, "Missing gradient attributes use hardcoded defaults; inherited gradient attributes are not resolved.");
      if (!r.bool(node, "rotWithShape", true)) return fail("Gradient rotation independent of its owner is unsupported.");
      const list = r.one(node, r.a, "gsLst"), linear = r.one(node, r.a, "lin");
      if (!list || !linear || r.one(node, r.a, "path") || r.one(node, r.a, "tileRect")) return fail("Only explicit linear gradients without tile rectangles are projected.");
      const nodes = children(list); if (nodes.length < 2 || nodes.length > 256 || nodes.some(n => n.namespaceUri !== r.a || n.local !== "gs")) return fail("Gradient requires 2–256 ordered stops.");
      consume(nodes.length, source); r.known(list); r.known(linear);
      if (r.attr(linear, "ang") === undefined || r.attr(linear, "scaled") === undefined) r.note(linear, "Missing linear attributes use hardcoded defaults; inherited gradient attributes are not resolved.");
      const angle = r.number(linear, "ang", 0, 21599999, 0) / 60000, scaled = r.bool(linear, "scaled", false), stops: { position: number; color: string; opacity: number }[] = [];
      const dependencies: PptxSourceLocation[] = [];
      for (const stop of nodes) {
        await checkpoint(); r.known(stop); const position = r.number(stop, "pos", 0, 100000) / 100000, choices = children(stop);
        if ((stops.length && position < stops[stops.length - 1]!.position) || choices.length !== 1 || choices[0]!.namespaceUri !== r.a) return fail("Gradient stops need nondecreasing positions and one supported color each.");
        const location = locate?.(stop); if (location) dependencies.push(location);
        const color = await colorValue(choices[0]!, source, dependencies, locate); if (!color) return fail("Gradient stop color cannot be resolved.");
        stops.push({ position, ...color });
      }
      onResolved?.(dependencies);
      return { type: "linearGradient", angle, scaled, stops };
    };
  const pattern = async (parent: XmlElement, source: PptxSourceLocation, locate?: (node: XmlElement) => PptxSourceLocation | undefined, onResolved?: (locations: readonly PptxSourceLocation[]) => void): Promise<ShapeFillPaintInput | undefined> => {
    const fills = children(parent).filter(n => n.namespaceUri === r.a && ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(n.local));
    if (fills.length > 1) invalidStructure("Conflicting drawing fill choices.", r.part);
    const node = r.one(parent, r.a, "pattFill"); if (!node) return undefined; r.known(node);
    const fail = (reason: string) => { r.note(node, reason); return undefined; };
    const value = r.attr(node, "prst"), patterns = { horz: "horizontal", vert: "vertical", cross: "cross", dnDiag: "downwardDiagonal", upDiag: "upwardDiagonal", diagCross: "diagonalCross", pct5: "percent5", pct10: "percent10", pct20: "percent20", pct25: "percent25", pct30: "percent30", pct40: "percent40", pct50: "percent50", pct60: "percent60", pct70: "percent70", pct75: "percent75", pct80: "percent80", pct90: "percent90" } as const;
    if (value === undefined || !Object.hasOwn(patterns, value)) return fail("Pattern preset is outside the supported profile.");
    const fg = r.one(node, r.a, "fgClr"), bg = r.one(node, r.a, "bgClr");
    if (!fg || !bg) return fail("Pattern requires explicit foreground and background colors.");
    const dependencies: PptxSourceLocation[] = [];
    const read = async (container: XmlElement) => {
      await checkpoint(); r.known(container); const choices = children(container);
      if (choices.length !== 1 || choices[0]!.namespaceUri !== r.a) return undefined;
      const location = locate?.(container); if (location) dependencies.push(location);
      return colorValue(choices[0]!, source, dependencies, locate);
    };
    consume(2, source);
    const foreground = await read(fg), background = await read(bg);
    if (!foreground || !background) return fail("Pattern colors cannot be resolved.");
    onResolved?.(dependencies);
    return { type: "pattern", pattern: patterns[value as keyof typeof patterns], foreground: foreground.color, background: background.color, foregroundOpacity: foreground.opacity, backgroundOpacity: background.opacity };
  };
  const shadow = async (parent: XmlElement, source: PptxSourceLocation, locate?: (node: XmlElement) => PptxSourceLocation | undefined, onResolved?: (locations: readonly PptxSourceLocation[]) => void): Promise<ShapeStyleInput["shadow"] | undefined> => {
    const list = r.one(parent, r.a, "effectLst"); if (!list) return undefined;
    const fail = (reason: string) => { r.note(list, reason); return undefined; };
    if (r.one(parent, r.a, "effectDag")) return fail("Effect DAG is outside the shadow profile.");
    const effects = children(list); if (effects.length !== 1 || effects[0]!.namespaceUri !== r.a || effects[0]!.local !== "outerShdw") return fail("Only a single explicit outer shadow is projected.");
    const node = effects[0]!, allowed = new Set(["blurRad", "dist", "dir", "algn", "rotWithShape"]);
    if (node.attributes.some(a => a.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (a.namespaceUri !== "" || !allowed.has(a.local)))) return fail("Shadow scaling, skew or extension attributes are unsupported.");
    const blur = r.attr(node, "blurRad"), distance = r.attr(node, "dist"), direction = r.attr(node, "dir"), alignment = r.attr(node, "algn"), rotation = r.attr(node, "rotWithShape");
    if (![blur ?? "0", distance, direction].every(v => v !== undefined && /^\d+$/.test(v)) || alignment !== "ctr" || !["1", "true"].includes(rotation ?? "")) return fail("Shadow requires valid radius, explicit distance and direction, centered alignment and rotation with its shape.");
    if (blur === undefined) r.note(node, "Missing shadow blur radius is projected as zero; source inheritance and producer-specific defaults are not resolved.");
    const b = Number(blur ?? "0"), d = Number(distance), angle = Number(direction);
    if (!Number.isSafeInteger(b) || b > 12700000 || !Number.isSafeInteger(d) || d > 12700000 || !Number.isSafeInteger(angle) || angle >= 21600000) return fail("Shadow dimensions or direction exceed the authoring profile.");
    const colors = children(node); if (colors.length !== 1 || colors[0]!.namespaceUri !== r.a) return fail("Shadow requires exactly one supported explicit color.");
    const dependencies: PptxSourceLocation[] = []; const at = locate?.(node); if (at) dependencies.push(at);
    consume(1, source); const color = await colorValue(colors[0]!, source, dependencies, locate); if (!color) return fail("Shadow color cannot be resolved.");
    r.known(list); r.known(node); onResolved?.(dependencies);
    return { color: color.color, opacity: color.opacity, blur: b / 12700, distance: d / 12700, direction: angle / 60000 };
  };
  return { paint, gradient, pattern, shadow };
}
