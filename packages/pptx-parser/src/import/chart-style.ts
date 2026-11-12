import type { CategoryAxisInput, ChartAxesInput, ComboChartAxesInput, ChartLegendInput, ChartMarkerInput, ChartTextStyleInput, ScatterChartAxesInput, StrokeStyleInput, ValueAxisInput } from "@pptkit/core";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";
import { chartNumber, chartNumberUnderflows } from "./chart-data.js";

export function chartScalar(node: XmlElement, local: string, c: string, r: ContentReader): string | undefined {
  const value = r.one(node, c, local);
  return value ? r.attr(value, "val") : undefined;
}
export function chartBoolean(node: XmlElement, local: string, c: string, r: ContentReader, fallback: boolean): boolean {
  const value = r.one(node, c, local);
  return value ? r.bool(value, "val", true) : fallback;
}
export function chartTextStyle(node: XmlElement | undefined, r: ContentReader): ChartTextStyleInput | undefined {
  if (!node) return undefined;
  r.known(node);
  const style: ChartTextStyleInput = {};
  if (attribute(node, "sz") !== undefined) style.fontSize = r.number(node, "sz", 100, 400000) / 100;
  if (attribute(node, "b") !== undefined) style.bold = r.bool(node, "b");
  if (attribute(node, "i") !== undefined) style.italic = r.bool(node, "i");
  const latin = r.one(node, r.a, "latin"), font = latin ? r.attr(latin, "typeface") : undefined;
  if (font && !font.startsWith("+")) style.fontFamily = font;
  else if (font) r.note(latin!, "Chart theme-font resolution remains source-only.");
  const paint = r.paint(node);
  if (paint?.type === "solid" && paint.opacity === 1) style.color = paint.color;
  else if (paint) r.note(node, "Chart text color cannot represent this paint/opacity.");
  return style;
}
export function chartTextProperties(node: XmlElement, c: string, r: ContentReader): ChartTextStyleInput | undefined {
  const tx = r.one(node, c, "txPr");
  if (!tx) return undefined;
  r.known(tx);
  const paragraphs = children(tx).filter(n => n.namespaceUri === r.a && n.local === "p");
  const p = paragraphs[0];
  if (!p) return undefined;
  r.known(p);
  if (paragraphs.length > 1) r.note(tx, "Multiple chart text-property paragraphs cannot map to a single Core chart text style.");
  const pPr = r.one(p, r.a, "pPr");
  if (pPr) r.known(pPr);
  const defaults = pPr ? r.one(pPr, r.a, "defRPr") : undefined;
  return chartTextStyle(defaults, r);
}
export function chartLegend(chart: XmlElement, c: string, r: ContentReader): ChartLegendInput {
  const legend = r.one(chart, c, "legend");
  if (!legend) return { visible: false };
  r.known(legend);
  const position = chartScalar(legend, "legendPos", c, r);
  if (position !== undefined && !["b", "r"].includes(position)) r.note(legend, "Legend position is outside Core's bottom/right subset; projection uses right.");
  if (chartBoolean(legend, "overlay", c, r, false)) r.note(legend, "Native overlaid legend is projected using Core's non-overlay layout.");
  const textStyle = chartTextProperties(legend, c, r);
  return { visible: true, position: position === "b" ? "bottom" : "right", ...(textStyle ? { textStyle } : {}) };
}
export function chartMarker(series: XmlElement, c: string, r: ContentReader): ChartMarkerInput | undefined {
  const marker = r.one(series, c, "marker");
  if (!marker) return undefined;
  r.known(marker);
  const shape = chartScalar(marker, "symbol", c, r);
  if (shape === "none") return false;
  const map = { diamond: "diamond", square: "square", circle: "circle", triangle: "triangle", star: "star", x: "x", plus: "plus", dash: "dash" } as const;
  if (!shape || !Object.hasOwn(map, shape)) { r.note(marker, "Marker shape is not mapped; Core's default marker is a projection value."); return undefined; }
  const size = r.one(marker, c, "size"), spPr = r.one(marker, c, "spPr");
  if (spPr) r.known(spPr);
  const fill = spPr ? r.paint(spPr) : undefined, stroke = spPr ? r.stroke(spPr) : undefined;
  return { shape: map[shape as keyof typeof map], ...(size ? { size: r.number(size, "val", 2, 72) } : {}), ...(fill ? { fill } : {}), ...(stroke ? { stroke } : {}) };
}
function axisStyle(axis: XmlElement, c: string, r: ContentReader): CategoryAxisInput {
  r.known(axis);
  const properties = r.one(axis, c, "spPr");
  if (properties) r.known(properties);
  const line = properties ? r.stroke(properties) : undefined;
  const tick = chartScalar(axis, "majorTickMark", c, r);
  const tickMap = { none: "none", in: "inside", out: "outside", cross: "cross" } as const;
  if (tick !== undefined && !Object.hasOwn(tickMap, tick)) invalidStructure("Invalid major tick mark.", r.part);
  const grid = r.one(axis, c, "majorGridlines"), gridProperties = grid ? r.one(grid, c, "spPr") : undefined;
  if (grid) r.known(grid); if (gridProperties) r.known(gridProperties);
  const gridStroke: false | StrokeStyleInput | undefined = grid ? gridProperties ? r.stroke(gridProperties) : undefined : false;
  const labelStyle = chartTextProperties(axis, c, r), labels = chartScalar(axis, "tickLblPos", c, r);
  if (labels !== undefined && !["none", "nextTo"].includes(labels)) r.note(axis, "High/low axis labels use the Core next-to-axis approximation.");
  return {
    visible: !chartBoolean(axis, "delete", c, r, false),
    ...(labels === undefined ? {} : { labels: labels !== "none" }),
    ...(line ? { line } : {}), ...(tick ? { majorTick: tickMap[tick as keyof typeof tickMap] } : {}),
    ...(gridStroke === undefined ? {} : { majorGridlines: gridStroke }),
    ...(labelStyle ? { labelStyle } : {}),
  };
}
export function chartAxes(plot: XmlElement, family: XmlElement, c: string, r: ContentReader): { axes: ChartAxesInput } | { reason: string } {
  const axisNodes = children(plot).filter(n => n.namespaceUri === c && ["catAx", "valAx", "dateAx", "serAx"].includes(n.local));
  const cat = axisNodes.filter(n => n.local === "catAx"), val = axisNodes.filter(n => n.local === "valAx");
  if (cat.length !== 1 || val.length !== 1 || axisNodes.length !== 2) return { reason: "Only one category/value axis pair is mapped; date/series/secondary axes remain source-only." };
  const ids = axisNodes.map(axis => {
    const id = r.one(axis, c, "axId", true)!;
    return r.number(id, "val", 0, 4294967295);
  });
  if (new Set(ids).size !== 2) invalidStructure("Chart axis identities collide.", r.part);
  const references = children(family).filter(n => n.namespaceUri === c && n.local === "axId").map(n => r.number(n, "val", 0, 4294967295));
  if (references.length !== 2 || new Set(references).size !== 2 || references.some(id => !ids.includes(id))) invalidStructure("Chart axis references do not resolve to its category/value pair.", r.part);
  for (const [index, axis] of axisNodes.entries()) {
    const cross = r.one(axis, c, "crossAx", true)!;
    if (r.number(cross, "val", 0, 4294967295) !== ids[1 - index]) invalidStructure("Chart crossing-axis identity is inconsistent.", r.part);
  }
  const category = axisStyle(cat[0]!, c, r), value = valueAxis(val[0]!, c, r);
  if ("reason" in value) return value;
  return { axes: { category, value } };
}

/** Bind each family to a disjoint, reciprocal axis pair before projecting. */
export function comboAxes(plot: XmlElement, bar: XmlElement, line: XmlElement, c: string, r: ContentReader): { axes: ComboChartAxesInput } | { reason: string } {
  const axisNodes = children(plot).filter(n => n.namespaceUri === c && ["catAx", "valAx", "dateAx", "serAx"].includes(n.local));
  if (axisNodes.length !== 4 || axisNodes.filter(node => node.local === "catAx").length !== 2 ||
    axisNodes.filter(node => node.local === "valAx").length !== 2) return { reason: "Combo charts require two category/value axis pairs." };
  const id = (node: XmlElement, local: string) => r.number(r.one(node, c, local, true)!, "val", 0, 4294967295);
  const byId = new Map(axisNodes.map(node => [id(node, "axId"), node]));
  if (byId.size !== 4) invalidStructure("Chart axis identities collide.", r.part);
  const used = new Set<number>();
  const pair = (family: XmlElement) => {
    const refs = children(family).filter(node => node.namespaceUri === c && node.local === "axId").map(node => r.number(node, "val", 0, 4294967295));
    if (refs.length !== 2 || new Set(refs).size !== 2 || refs.some(key => !byId.has(key))) invalidStructure("Combo axis references do not resolve.", r.part);
    if (refs.some(key => used.has(key))) return undefined;
    for (const key of refs) used.add(key);
    const [a, b] = refs.map(key => byId.get(key)!);
    if (id(a!, "crossAx") !== refs[1] || id(b!, "crossAx") !== refs[0]) return undefined;
    if (a!.local === b!.local) return undefined;
    return a!.local === "catAx" ? { category: a!, value: b! } : { category: b!, value: a! };
  };
  const primary = pair(bar), secondary = pair(line);
  if (!primary || !secondary) return { reason: "Combo charts require distinct reciprocal category/value axis pairs." };
  // PowerPoint rewrites the hidden secondary category axis from top to bottom.
  if (chartScalar(primary.category, "axPos", c, r) !== "b" || chartScalar(primary.value, "axPos", c, r) !== "l" ||
    !["t", "b"].includes(chartScalar(secondary.category, "axPos", c, r) ?? "") || chartScalar(secondary.value, "axPos", c, r) !== "r" ||
    !chartBoolean(secondary.category, "delete", c, r, false) || chartScalar(secondary.value, "crosses", c, r) !== "max") {
    return { reason: "Combo chart axis positions or secondary category visibility are outside the supported profile." };
  }
  for (const node of axisNodes) {
    const crossing = chartScalar(node, "crosses", c, r);
    if (r.one(node, c, "crossesAt") || (node !== secondary.value && crossing !== undefined && crossing !== "autoZero")) {
      return { reason: "Custom combo axis crossings remain source-only." };
    }
    const scaling = r.one(node, c, "scaling");
    if (scaling && (![undefined, "minMax"].includes(chartScalar(scaling, "orientation", c, r)) || r.one(scaling, c, "logBase"))) {
      return { reason: "Reversed or logarithmic combo axes remain source-only." };
    }
  }
  if ([primary.value, secondary.value].some(axis => chartScalar(axis, "crossBetween", c, r) !== "between")) {
    return { reason: "Combo axes must place both families at shared category centers." };
  }
  const category = axisStyle(primary.category, c, r), value = valueAxis(primary.value, c, r);
  const secondaryValue = valueAxis(secondary.value, c, r);
  if ("reason" in value) return value;
  if ("reason" in secondaryValue) return secondaryValue;
  r.known(secondary.category);
  return { axes: { category, value, secondaryValue } };
}

function valueAxis(axis: XmlElement, c: string, r: ContentReader): ValueAxisInput | { reason: string } {
  const value = axisStyle(axis, c, r);
  const scaling = r.one(axis, c, "scaling");
  if (scaling) r.known(scaling);
  const min = scaling ? chartScalar(scaling, "min", c, r) : undefined;
  const max = scaling ? chartScalar(scaling, "max", c, r) : undefined;
  const step = chartScalar(axis, "majorUnit", c, r);
  const bounds = [min, max, step].map(v => v === undefined ? undefined : chartNumber(v));
  if ([min, max, step].some(v => v !== undefined && chartNumberUnderflows(v))) return { reason: "A nonzero value-axis bound underflows Core's numeric range." };
  if ([min, max, step].some((v, i) => v !== undefined && bounds[i] === undefined)) invalidStructure("Chart axis bounds require finite numeric values.", r.part);
  if (bounds.every((n): n is number => n !== undefined)) {
    const [lo, hi, unit] = bounds as [number, number, number];
    if (hi <= lo || unit <= 0 || !Number.isFinite(hi - lo) || (hi - lo) / unit > 10000) return { reason: "Value-axis scale is invalid or exceeds a bounded Core tick profile." };
    return { ...value, scale: { min: lo, max: hi, majorUnit: unit } };
  }
  if ([min, max, step].some(v => v !== undefined)) r.note(axis, "Partial explicit axis bounds cannot map to Core's complete manual scale; Core auto scale is a projection value.");
  return value;
}

export function scatterAxes(plot: XmlElement, family: XmlElement, c: string, r: ContentReader): { axes: ScatterChartAxesInput } | { reason: string } {
  const nodes = children(plot).filter(n => n.namespaceUri === c && ["catAx", "valAx", "dateAx", "serAx"].includes(n.local));
  if (nodes.length !== 2 || nodes.some(n => n.local !== "valAx")) return { reason: "Scatter requires exactly two value axes; other axis configurations remain source-only." };
  // PowerPoint opens and saves legacy signed 32-bit identities from some
  // producers. Compare their numeric values exactly; do not reinterpret a
  // negative value as its unsigned two's-complement spelling.
  const axisIdentityNodes = [...nodes.flatMap(n => [r.one(n, c, "axId", true)!, r.one(n, c, "crossAx", true)!]), ...children(family).filter(n => n.namespaceUri === c && n.local === "axId")];
  if (axisIdentityNodes.some(n => { const value = r.attr(n, "val")?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "") ?? "", number = Number(value); return !/^[+-]?\d+$/.test(value) || !Number.isSafeInteger(number) || number < -2147483648 || number > 4294967295; })) return { reason: "Scatter axis identities are outside the bounded signed/unsigned 32-bit profile; source remains unchanged." };
  if (axisIdentityNodes.some(n => Number(r.attr(n, "val")) < 0)) r.note(family, "Legacy signed scatter axis identities are compared directly and preserved as source text; cross-producer compatibility is not assumed.");
  const ids = nodes.map(n => r.number(r.one(n, c, "axId", true)!, "val", -2147483648, 4294967295));
  if (new Set(ids).size !== 2) invalidStructure("Chart axis identities collide.", r.part);
  const refs = children(family).filter(n => n.namespaceUri === c && n.local === "axId").map(n => r.number(n, "val", -2147483648, 4294967295));
  if (refs.length !== 2 || new Set(refs).size !== 2 || refs.some(id => !ids.includes(id))) invalidStructure("Scatter axis references do not resolve to its value axes.", r.part);
  for (const [i, node] of nodes.entries()) if (r.number(r.one(node, c, "crossAx", true)!, "val", -2147483648, 4294967295) !== ids[1 - i]) invalidStructure("Chart crossing-axis identity is inconsistent.", r.part);
  const x = nodes.filter(n => ["b", "t"].includes(chartScalar(n, "axPos", c, r) ?? ""));
  const y = nodes.filter(n => ["l", "r"].includes(chartScalar(n, "axPos", c, r) ?? ""));
  if (x.length !== 1 || y.length !== 1) return { reason: "Scatter axis positions do not identify one horizontal and one vertical value axis." };
  const horizontal = valueAxis(x[0]!, c, r), vertical = valueAxis(y[0]!, c, r);
  if ("reason" in horizontal) return horizontal;
  if ("reason" in vertical) return vertical;
  return { axes: { x: horizontal, y: vertical } };
}
