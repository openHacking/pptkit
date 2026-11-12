import { attribute, children } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";

type Channels = [number, number, number];
export type ColorTransformResult =
  | { status: "mapped"; value: string; opacity: number }
  | { status: "unsupported"; reason: string };

const clamp = (value: number) => Math.max(0, Math.min(1, value));
const linear = (value: number) => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
const encoded = (value: number) => value <= 0.0031308 ? value * 12.92 : 1.055 * value ** (1 / 2.4) - 0.055;
const names = new Set(["tint", "shade", "lum", "lumMod", "lumOff", "sat", "satMod", "satOff", "alpha", "alphaMod", "alphaOff"]);

function toHsl([r, g, b]: Channels): Channels {
  const hi = Math.max(r, g, b), lo = Math.min(r, g, b), delta = hi - lo, l = (hi + lo) / 2;
  if (!delta) return [0, 0, l];
  const h = hi === r ? ((g - b) / delta + 6) % 6 : hi === g ? (b - r) / delta + 2 : (r - g) / delta + 4;
  return [h / 6, delta / (1 - Math.abs(2 * l - 1)), l];
}

function fromHsl([h, s, l]: Channels): Channels {
  h = ((h % 1) + 1) % 1;
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs((h * 6) % 2 - 1)), m = l - c / 2;
  const sectors: Channels[] = [[c, x, 0], [x, c, 0], [0, c, x], [0, x, c], [x, 0, c], [c, 0, x]];
  return sectors[Math.floor(h * 6)]!.map(value => value + m) as Channels;
}

function percentage(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined;
  const value = raw.replace(/^[\x20\t\r\n]+|[\x20\t\r\n]+$/g, "");
  if (/^[+-]?\d+$/.test(value)) {
    const integer = Number(value);
    return Number.isInteger(integer) && integer >= -2147483648 && integer <= 2147483647 ? integer / 100000 : undefined;
  }
  if (/^[+-]?\d+(?:\.\d+)?%$/.test(value)) {
    const result = Number(value.slice(0, -1)) / 100;
    return Number.isFinite(result) ? result : undefined;
  }
  return undefined;
}

/**
 * Bounded DrawingML transform profile. RGB remains floating point until final
 * quantization; tint/shade mix in linear scRGB using IEC sRGB transfer.
 * Consecutive HSL operations retain hue/saturation at achromatic colors.
 * Unknown, malformed or out-of-profile transforms never yield partial colors.
 */
export function transformTextColor(rgb: string, modifiers: readonly XmlElement[], namespace: string): ColorTransformResult {
  if (!/^[\da-f]{6}$/i.test(rgb)) return { status: "unsupported", reason: "Invalid direct RGB color." };
  let channels = [0, 2, 4].map(index => parseInt(rgb.slice(index, index + 2), 16) / 255) as Channels;
  let space: "rgb" | "linear" | "hsl" = "rgb", alpha = 1;
  const asRgb = (): Channels => space === "hsl" ? fromHsl(channels) : space === "linear" ? channels.map(encoded) as Channels : channels;
  for (const modifier of modifiers) {
    const name = modifier.local;
    if (modifier.namespaceUri !== namespace || !names.has(name) || children(modifier).length
      || modifier.attributes.some(attr => attr.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (attr.namespaceUri !== "" || attr.local !== "val"))
      || modifier.children.some(child => (child.kind === "text" || child.kind === "cdata") && child.value.trim() !== "")) {
      return { status: "unsupported", reason: `Unsupported color transform ${name} blocks the entire color declaration.` };
    }
    const amount = percentage(attribute(modifier, "val"));
    if (amount === undefined) return { status: "unsupported", reason: `Color transform ${name} has an invalid percentage.` };
    const modulation = name.endsWith("Mod"), offset = name.endsWith("Off");
    if ((["tint", "shade", "alpha"].includes(name) && (amount < 0 || amount > 1))
      || (name === "alphaOff" && (amount < -1 || amount > 1)) || (name === "alphaMod" && amount < 0)) {
      return { status: "unsupported", reason: `Color transform ${name} is outside its schema percentage range.` };
    }
    if (name.startsWith("alpha")) {
      alpha = clamp(name === "alpha" ? amount : modulation ? alpha * amount : alpha + amount);
    } else if (name === "tint" || name === "shade") {
      if (space !== "linear") channels = asRgb().map(linear) as Channels;
      space = "linear";
      channels = channels.map(value => clamp(name === "shade" ? value * amount : 1 - (1 - value) * amount)) as Channels;
    } else {
      if (space !== "hsl") channels = toHsl(asRgb());
      space = "hsl";
      const index = name.startsWith("lum") ? 2 : 1;
      channels[index] = clamp(modulation ? channels[index] * amount : offset ? channels[index] + amount : amount);
    }
  }
  return {
    status: "mapped",
    value: asRgb().map(value => Math.round(clamp(value) * 255 + 1e-10).toString(16).padStart(2, "0")).join("").toUpperCase(),
    opacity: alpha,
  };
}
