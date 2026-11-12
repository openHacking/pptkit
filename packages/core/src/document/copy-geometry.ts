import type { PresentationElementInput } from "../types/element.js";
import type { PresentationSize } from "../types/geometry.js";

/** Uniformly fit the source page in the destination, centered in the unused axis.
 * Only authored geometry in the page coordinate system is changed. */
export function fitCopiedGeometry(elements: PresentationElementInput[], source: PresentationSize, target: PresentationSize): void {
  if (source.unit !== "pt" || target.unit !== "pt" || !Number.isFinite(source.width) || !Number.isFinite(source.height) || source.width <= 0 || source.height <= 0) throw new TypeError("Copy geometry requires positive finite point dimensions.");
  const scale = Math.min(target.width / source.width, target.height / source.height);
  const x = (target.width - source.width * scale) / 2, y = (target.height - source.height * scale) / 2;
  if (!Number.isFinite(scale) || scale <= 0 || !Number.isFinite(x) || !Number.isFinite(y)) throw new RangeError("Copy geometry scale must be positive and finite.");
  for (const element of elements) {
    if (element.box) {
      element.box = { ...element.box, x: element.box.x * scale + x, y: element.box.y * scale + y, width: element.box.width * scale, ...(element.box.height === undefined ? {} : { height: element.box.height * scale }) };
    }
    if (element.type === "connector") {
      const convert = <T extends { x: number; y: number }>(point: T): T => ({ ...point, x: point.x * scale + x, y: point.y * scale + y });
      if (!("elementId" in element.start)) element.start = convert(element.start);
      if (!("elementId" in element.end)) element.end = convert(element.end);
      if (element.route) element.route = element.route.map(convert);
    }
    if (element.type === "table") {
      element.columns = element.columns.map(width => width * scale);
      element.rows = element.rows.map(row => row.height === undefined ? row : { ...row, height: row.height * scale });
    }
  }
}
