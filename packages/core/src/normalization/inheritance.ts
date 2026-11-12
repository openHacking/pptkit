import { DEFAULT_BACKGROUND, DEFAULT_THEME } from "../constants/presentation.js";
import type { PresentationInheritanceOrigin, PresentationInheritanceProperty, PresentationInheritanceSource, PresentationInheritanceTarget } from "../types/inheritance.js";
import type { NormalizedBackgroundPaint, BackgroundPaintInput, ThemeColorRole } from "../types/style.js";
import type { NormalizedPresentationTheme, PresentationThemeInput } from "../types/theme.js";
import { deepClone } from "../utils/clone.js";
import { normalizePaint } from "./normalize-styles.js";

type ThemeProperty = Extract<PresentationInheritanceProperty, `theme.${string}`>;
export interface ResolvedTheme {
  value: NormalizedPresentationTheme;
  origins: Record<ThemeProperty, PresentationInheritanceSource>;
}
export interface ResolvedBackground {
  value: NormalizedBackgroundPaint;
  origins: Partial<Record<"background" | "background.type" | "background.color" | "background.opacity" | "background.angle" | "background.scaled" | "background.stops" | `background.stops.${number}.${"position" | "color" | "opacity"}`, PresentationInheritanceSource>>;
}
export type OriginSink = (origin: PresentationInheritanceOrigin) => void;

function source(owner: PresentationInheritanceSource, suffix: string): PresentationInheritanceSource {
  return { ...owner, path: `${owner.path}${suffix}` };
}

export function defaultTheme(): ResolvedTheme {
  const origins = { "theme.name": { kind: "default", path: "defaults.theme.name" } } as ResolvedTheme["origins"];
  for (const role of Object.keys(DEFAULT_THEME.colors) as ThemeColorRole[]) origins[`theme.colors.${role}`] = { kind: "default", path: `defaults.theme.colors.${role}` };
  for (const role of ["heading", "body"] as const) origins[`theme.fonts.${role}`] = { kind: "default", path: `defaults.theme.fonts.${role}` };
  return { value: deepClone(DEFAULT_THEME), origins };
}

/** Value selection and its source are recorded together, preserving the legacy normalized contract. */
export function resolveTheme(input: Readonly<PresentationThemeInput> | undefined, fallback: ResolvedTheme, owner: PresentationInheritanceSource): ResolvedTheme {
  const name = input?.name, colors = { ...input?.colors }, fonts = { ...input?.fonts };
  const origins = { ...fallback.origins };
  if (name != null) origins["theme.name"] = source(owner, ".name");
  for (const role of Object.keys(DEFAULT_THEME.colors) as ThemeColorRole[]) if (Object.hasOwn(colors, role)) origins[`theme.colors.${role}`] = source(owner, `.colors.${role}`);
  for (const role of ["heading", "body"] as const) if (Object.hasOwn(fonts, role)) origins[`theme.fonts.${role}`] = source(owner, `.fonts.${role}`);
  return { value: { name: name ?? fallback.value.name, colors: { ...fallback.value.colors, ...colors }, fonts: { ...fallback.value.fonts, ...fonts } }, origins };
}

export function defaultBackground(): ResolvedBackground {
  return { value: deepClone(DEFAULT_BACKGROUND), origins: {
    background: { kind: "default", path: "defaults.background" },
    "background.type": { kind: "default", path: "defaults.background.type" },
    "background.color": { kind: "default", path: "defaults.background.color" },
    "background.opacity": { kind: "default", path: "defaults.background.opacity" },
  } };
}

export function resolveBackground(input: BackgroundPaintInput | undefined, fallback: ResolvedBackground, owner: PresentationInheritanceSource, detailed = false): ResolvedBackground {
  if (input === undefined) return { value: deepClone(fallback.value), origins: { ...fallback.origins } };
  if (input.type === "linearGradient") {
    const origins: ResolvedBackground["origins"] = {
      background: owner, "background.type": source(owner, ".type"), "background.angle": source(owner, ".angle"),
      "background.scaled": input.scaled === undefined ? { kind: "default", path: "defaults.background.linearGradient.scaled" } : source(owner, ".scaled"),
      "background.stops": source(owner, ".stops"),
    };
    if (detailed) input.stops.forEach((stop, index) => {
      origins[`background.stops.${index}.position`] = source(owner, `.stops.${index}.position`);
      origins[`background.stops.${index}.color`] = source(owner, `.stops.${index}.color`);
      origins[`background.stops.${index}.opacity`] = stop.opacity === undefined ? { kind: "default", path: "defaults.background.linearGradient.stop.opacity" } : source(owner, `.stops.${index}.opacity`);
    });
    return { value: { type: "linearGradient", angle: input.angle, scaled: input.scaled ?? false, stops: input.stops.map(stop => ({ position: stop.position, color: deepClone(stop.color), opacity: stop.opacity ?? 1 })) }, origins };
  }
  let paint: BackgroundPaintInput;
  if (input.type === "none") paint = { type: "none" };
  else {
    const opacity = input.opacity;
    paint = { type: "solid", color: input.color, ...(opacity !== undefined ? { opacity } : {}) };
  }
  return { value: normalizePaint(paint, fallback.value.type === "linearGradient" ? { type: "none" } : fallback.value), origins: {
    background: owner, "background.type": source(owner, ".type"),
    ...(paint.type === "solid" ? {
      "background.color": source(owner, ".color"),
      "background.opacity": paint.opacity === undefined ? { kind: "default" as const, path: "defaults.paint.opacity" } : source(owner, ".opacity"),
    } : {}),
  } };
}

export function recordOrigins(target: PresentationInheritanceTarget, theme: ResolvedTheme | undefined, background: ResolvedBackground | undefined, sink: OriginSink | undefined): void {
  if (!sink) return;
  for (const [property, source] of Object.entries({ ...theme?.origins, ...background?.origins })) sink({ target, property: property as PresentationInheritanceProperty, source });
}

export function backgroundSource(background: ResolvedBackground): "slide" | "layout" | "master" | "theme" {
  const kind = background.origins.background!.kind;
  return kind === "slide" || kind === "layout" || kind === "master" ? kind : "theme";
}
