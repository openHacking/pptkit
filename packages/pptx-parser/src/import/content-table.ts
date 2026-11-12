import type { Box, ElementAction, NativeTableStyleFlags, StrokeStyleInput, TableCellEdge, TableCellInput, TableCellStyleInput, TableElementInput, TableRowInput } from "@pptkit/core";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";
import type { SourceTableStyle } from "./table-style-inheritance.js";
import { contentText, type ContentTextRun } from "./content-text.js";

export interface ContentTableTextRun extends ContentTextRun {
  readonly cell: XmlElement;
  readonly rowIndex: number;
  readonly cellIndex: number;
  readonly columnIndex: number;
}

type Fields = Pick<TableElementInput, "type" | "columns" | "rows" | "nativeFirstRowStyle" | "nativeTableStyleFlags" | "nativeTableStylePreset">;
type Merge = { row: number; column: number; rows: number; columns: number };

/** Convert physical DrawingML grid cells into Core's compact, anchor-only rows. */
export async function contentTable(
  node: XmlElement, box: Box, r: ContentReader, checkpoint: () => Promise<void>,
  readAction: (node: XmlElement) => ElementAction | undefined,
  recordRun?: (run: ContentTableTextRun) => void,
  styleFactory?: (properties: XmlElement | undefined) => Promise<SourceTableStyle>,
  fieldTextProfile?: "slide-number-v1" | "date-and-number-v1" | "date-time-and-number-v1",
): Promise<{ fields: Fields; previewFallback: boolean } | { reason: string }> {
  const graphic = r.one(node, r.a, "graphic"), data = graphic ? r.one(graphic, r.a, "graphicData") : undefined;
  const uri = data ? attribute(data, "uri") : undefined;
  if (!data || !["http://schemas.openxmlformats.org/drawingml/2006/table", "http://purl.oclc.org/ooxml/drawingml/table"].includes(uri ?? "")) return { reason: "This graphic frame is not the supported DrawingML table family." };
  r.known(graphic!); r.attr(data, "uri");
  const table = r.one(data, r.a, "tbl", true)!; r.known(table);
  const grid = r.one(table, r.a, "tblGrid", true)!; r.known(grid);
  const gridColumns = children(grid).filter(n => n.namespaceUri === r.a && n.local === "gridCol");
  const sourceRows = children(table).filter(n => n.namespaceUri === r.a && n.local === "tr");
  if (!gridColumns.length || !sourceRows.length) invalidStructure("A table requires grid columns and rows.", r.part);
  const columns: number[] = [], heights: number[] = [];
  let totalWidth = 0, totalHeight = 0;
  for (const column of gridColumns) {
    await checkpoint();
    const width = r.number(column, "w", -27273042329600, 27273042316900);
    if (width <= 0) return { reason: "Nonpositive table columns have no positive Core column mapping." };
    totalWidth += width; columns.push(width / 12700);
  }
  for (const row of sourceRows) {
    await checkpoint();
    const height = r.number(row, "h", -27273042329600, 27273042316900);
    if (height <= 0) return { reason: "Nonpositive table rows have no positive Core row mapping." };
    totalHeight += height; heights.push(height / 12700);
  }
  if (!Number.isSafeInteger(totalWidth) || !Number.isSafeInteger(totalHeight)) return { reason: "Table grid sums exceed exact integer-EMU projection precision." };
  if (!box.width || !box.height) return { reason: "Zero-area graphic-frame geometry has no current table preview mapping." };
  const previewFallback = Math.abs(box.width * 12700 - totalWidth) > 1 || Math.abs(box.height * 12700 - totalHeight) > 1;
  if (previewFallback) r.note(grid, "Frame extent differs from authored grid sums; both values are retained, but Core preview does not resolve native table sizing.");
  const properties = r.one(table, r.a, "tblPr");
  if (properties) r.known(properties); // Remaining style flags and inherited styling stay in the audit.
  const styleFlag = (name: string): boolean | undefined => {
    const value = properties ? r.attr(properties, name)?.trim() : undefined;
    if (value === undefined || value === "0" || value === "false") return false;
    if (value === "1" || value === "true") return true;
    r.note(properties!, `Unrecognized ${name} flag remains source-only.`);
    return undefined;
  };
  const nativeFirstRowStyle = styleFlag("firstRow");
  const nativeTableStyleFlags: NativeTableStyleFlags = {};
  for (const [field, attributeName] of [
    ["firstColumn", "firstCol"], ["lastColumn", "lastCol"], ["lastRow", "lastRow"],
    ["bandedRows", "bandRow"], ["bandedColumns", "bandCol"],
  ] as const) {
    const value = styleFlag(attributeName);
    if (value !== undefined) nativeTableStyleFlags[field] = value;
  }

  // Validate the complete physical grid before parsing actions or publishing cells.
  // Occupancy costs O(columns); declared spans cannot allocate a rectangular grid.
  const occupied: (Merge | undefined)[] = columns.map(() => undefined);
  const anchors: { node: XmlElement; merge: Merge }[][] = [];
  for (const [rowIndex, row] of sourceRows.entries()) {
    await checkpoint();
    const cells = children(row).filter(n => n.namespaceUri === r.a && n.local === "tc");
    if (cells.length !== columns.length) return { reason: "Table rows must contain one physical cell per grid column; sparse or excess-cell encodings remain source-only." };
    const rowAnchors: { node: XmlElement; merge: Merge }[] = [];
    for (const [columnIndex, cell] of cells.entries()) {
      await checkpoint(); r.known(cell);
      const horizontal = r.bool(cell, "hMerge"), vertical = r.bool(cell, "vMerge");
      const columnSpan = r.number(cell, "gridSpan", 1, 2147483647, 1), rowSpan = r.number(cell, "rowSpan", 1, 2147483647, 1);
      let owner = occupied[columnIndex];
      if (owner && owner.row + owner.rows <= rowIndex) { occupied[columnIndex] = undefined; owner = undefined; }
      if (owner) {
        if (horizontal !== (columnIndex > owner.column) || vertical !== (rowIndex > owner.row)) invalidStructure("Table merge continuation flags disagree with the anchor rectangle.", r.part);
        const expectedColumns = columnIndex === owner.column ? owner.columns : 1;
        const expectedRows = rowIndex === owner.row ? owner.rows : 1;
        if ((columnSpan !== 1 && columnSpan !== expectedColumns) || (rowSpan !== 1 && rowSpan !== expectedRows)) invalidStructure("Table merge continuation spans disagree with the anchor rectangle.", r.part);
        if (await continuationHasContent(cell, r, checkpoint)) return { reason: "A merged continuation contains text, fields, breaks, explicit text formatting/actions or unsupported text markup that Core anchor-only cells cannot retain; the whole table remains source-only." };
        continue;
      }
      if (horizontal || vertical) invalidStructure("Table merge continuation has no preceding anchor.", r.part);
      if (columnIndex + columnSpan > columns.length || rowIndex + rowSpan > sourceRows.length) invalidStructure("Table span exceeds the declared grid.", r.part);
      const merge = { row: rowIndex, column: columnIndex, rows: rowSpan, columns: columnSpan };
      for (let c = columnIndex; c < columnIndex + columnSpan; c++) {
        const previous = occupied[c];
        if (previous && previous.row + previous.rows > rowIndex) invalidStructure("Overlapping table merge rectangles.", r.part);
        occupied[c] = merge;
      }
      rowAnchors.push({ node: cell, merge });
    }
    anchors.push(rowAnchors);
  }
  const sourceStyle = styleFactory ? await styleFactory(properties) : undefined;
  if (sourceStyle?.conditional) r.note(table, "Conditional table style regions use the experimental direct-region projection order; native overlap precedence is not certified. Merged cells additionally use an anchor/boundary approximation.");
  const rows: TableRowInput[] = [];
  const borderStyles = await mergedBorderStyles(sourceRows, anchors, columns.length, r, checkpoint, sourceStyle);
  for (const [index, row] of anchors.entries()) {
    await checkpoint();
    const cells: TableCellInput[] = [];
    for (const { node: cell, merge } of row) {
      await checkpoint();
      const body = r.one(cell, r.a, "txBody");
      const inherited = await sourceStyle?.cell(cell, index, merge.column, merge.rows, merge.columns, sourceRows.length, columns.length);
      const content = body ? (await contentText(body, r, checkpoint, readAction, false, inherited?.text,
        recordRun ? run => recordRun({ ...run, cell, rowIndex: index, cellIndex: cells.length, columnIndex: merge.column }) : undefined, fieldTextProfile)).content : "";
      if (!body) recordRun?.({ text: "", paragraphIndex: 0, runIndex: 0, cell, rowIndex: index, cellIndex: cells.length, columnIndex: merge.column });
      const edges = borderStyles.get(merge)!, style = await cellStyle(cell, body, r, edges, inherited?.style, sourceStyle);
      inherited?.record(style, edges);
      cells.push({ content, rowSpan: merge.rows, colSpan: merge.columns, style });
    }
    rows.push({ height: heights[index]!, cells });
  }
  return { fields: { type: "table", columns, rows, ...(sourceStyle?.nativePreset === undefined ? {} : { nativeTableStylePreset: sourceStyle.nativePreset }), ...(Object.keys(nativeTableStyleFlags).length ? { nativeTableStyleFlags } : {}), ...(nativeFirstRowStyle === undefined ? {} : { nativeFirstRowStyle }) }, previewFallback: previewFallback || Boolean(sourceStyle?.conditional) || sourceStyle?.nativePreset === "medium-2-accent-2" };
}

async function continuationHasContent(cell: XmlElement, r: ContentReader, checkpoint: () => Promise<void>): Promise<boolean> {
  const body = r.one(cell, r.a, "txBody");
  if (body && children(body).filter(child => child.namespaceUri === r.a && child.local === "p").length > 1) return true;
  const stack: { node: XmlElement; parent?: string }[] = body ? [{ node: body }] : [];
  const parentOf: Record<string, string> = { bodyPr: "txBody", lstStyle: "txBody", p: "txBody", r: "p", pPr: "p", endParaRPr: "p", rPr: "r", t: "r" };
  while (stack.length) {
    await checkpoint();
    const { node: current, parent } = stack.pop()!;
    if (current.namespaceUri !== r.a || !["txBody", "bodyPr", "lstStyle", "p", "r", "t", "rPr", "pPr", "endParaRPr"].includes(current.local)) return true;
    if (parent !== undefined && parentOf[current.local] !== parent) return true;
    if (current.local !== "t" && current.children.some(n => (n.kind === "text" || n.kind === "cdata") && n.value.trim().length > 0)) return true;
    if (current.namespaceUri === r.a && ["bodyPr", "lstStyle", "rPr", "pPr", "endParaRPr"].includes(current.local)) {
      // Empty continuation paragraphs emitted by this writer have only the
      // default insertion language. Other declarations or child elements can
      // carry formatting or hyperlinks that an anchor-only cell cannot retain.
      if (children(current).length || current.attributes.some(attr => attr.namespaceUri !== "http://www.w3.org/2000/xmlns/" && !(current.local === "endParaRPr" && attr.namespaceUri === "" && attr.local === "lang" && attr.value === "en-US"))) return true;
    } else if (current.attributes.some(attr => attr.namespaceUri !== "http://www.w3.org/2000/xmlns/" && !(current.local === "t" && attr.namespaceUri === "http://www.w3.org/XML/1998/namespace" && attr.local === "space" && attr.value === "preserve"))) {
      return true;
    }
    if (current.namespaceUri === r.a && current.local === "t" && current.children.some(n => (n.kind === "text" || n.kind === "cdata") && n.value.length > 0)) return true;
    for (const child of children(current)) stack.push({ node: child, parent: current.local });
  }
  return false;
}

type Borders = Partial<Record<TableCellEdge, StrokeStyleInput>>;
/** Stream physical borders with O(columns + anchors) state, never a source rectangle allocation. */
async function mergedBorderStyles(sourceRows: readonly XmlElement[], anchors: readonly { node: XmlElement; merge: Merge }[][], columns: number, r: ContentReader, checkpoint: () => Promise<void>, sourceStyle?: SourceTableStyle): Promise<Map<Merge, Borders>> {
  type Observation = { seen: boolean; key?: string; first?: StrokeStyleInput; equal: boolean; conflict?: XmlElement };
  const entries = new Map<Merge, Record<TableCellEdge, Observation>>(), active: (Merge | undefined)[] = new Array(columns);
  const anchorBorders = new Map<Merge, Borders>(), continuationBorders = new Set<Merge>();
  const names = { left: "lnL", right: "lnR", top: "lnT", bottom: "lnB" } as const;
  for (let row = 0; row < sourceRows.length; row++) {
    for (const { merge } of anchors[row]!) {
      entries.set(merge, { left: { seen: false, equal: true }, right: { seen: false, equal: true }, top: { seen: false, equal: true }, bottom: { seen: false, equal: true } });
      for (let column = merge.column; column < merge.column + merge.columns; column++) { await checkpoint(); active[column] = merge; }
    }
    const cells = children(sourceRows[row]!).filter(n => n.namespaceUri === r.a && n.local === "tc");
    for (let column = 0; column < columns; column++) {
      await checkpoint(); const cell = cells[column]!, merge = active[column]!;
      const properties = r.one(cell, r.a, "tcPr");
      const anchor = row === merge.row && column === merge.column;
      if (anchor) anchorBorders.set(merge, {});
      else if (properties && children(properties).some(n => n.namespaceUri === r.a && Object.values(names).includes(n.local as typeof names[TableCellEdge]))) continuationBorders.add(merge);
      for (const side of ["left", "right", "top", "bottom"] as const) {
        await checkpoint();
        const border = properties ? sourceStyle ? await sourceStyle.stroke(properties, names[side]) : r.stroke(properties, names[side]) : undefined;
        if (anchor && border !== undefined) { const copy = { ...border, paint: border.paint ?? { type: "none" as const } }; sourceStyle?.copyStroke(border, copy); anchorBorders.get(merge)![side] = copy; }
        const outer = side === "left" ? column === merge.column : side === "right" ? column === merge.column + merge.columns - 1 : side === "top" ? row === merge.row : row === merge.row + merge.rows - 1;
        if (!outer) {
          if (!anchor && border?.paint?.type === "solid" && (border.paint.opacity ?? 1) > 0 && (border.width ?? 1) > 0) r.note(properties!, "A visible physical line inside a merged cell has no Core logical-edge counterpart.");
          continue;
        }
        const observation = entries.get(merge)![side], key = JSON.stringify(border);
        if (!observation.seen) { observation.seen = true; if (key !== undefined) observation.key = key; if (border !== undefined) observation.first = border; }
        else if (observation.key !== key && observation.equal) {
          observation.equal = false;
          observation.conflict = properties ?? cell;
        }
      }
    }
  }
  const result = new Map<Merge, Borders>();
  for (const [merge, observations] of entries) {
    // Retain the established anchor-border projection for older packages that
    // declare no continuation lines. Explicit physical continuation styles use
    // the complete outer-side checks below; no missing segment is invented.
    if ((merge.rows > 1 || merge.columns > 1) && !continuationBorders.has(merge)) { result.set(merge, anchorBorders.get(merge)!); continue; }
    const borders: Borders = {};
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const observation = observations[side];
      if (observation.equal && observation.first !== undefined) { const copy = { ...observation.first, paint: observation.first.paint ?? { type: "none" as const } }; sourceStyle?.copyStroke(observation.first, copy); borders[side] = copy; }
      if (!observation.equal && observation.conflict) {
        const reason = `Unequal or partially missing physical ${side} border segments cannot be represented by one logical cell edge; that edge remains source-only.`;
        if (sourceStyle?.inherited) borders[side] = sourceStyle.blockStroke(observation.conflict, reason);
        else r.note(observation.conflict, reason);
      }
    }
    result.set(merge, borders);
  }
  return result;
}

async function cellStyle(cell: XmlElement, body: XmlElement | undefined, r: ContentReader, edges: Borders, inherited?: TableCellStyleInput, sourceStyle?: SourceTableStyle): Promise<TableCellStyleInput> {
  const properties = r.one(cell, r.a, "tcPr"), frame = body ? r.one(body, r.a, "bodyPr") : undefined;
  const style: TableCellStyleInput = { fill: { type: "none" }, stroke: { paint: { type: "none" } }, ...inherited };
  const margin: Partial<{ left: number; right: number; top: number; bottom: number }> = {};
  // Body insets are accepted for existing PPTKit output; standard tcPr values
  // override them. Other frame semantics remain unconsumed and reported.
  for (const [parent, names] of [[frame, ["lIns", "rIns", "tIns", "bIns"]], [properties, ["marL", "marR", "marT", "marB"]]] as const) {
    if (!parent) continue;
    r.known(parent);
    for (const [i, name] of names.entries()) if (attribute(parent, name) !== undefined) {
      const value = r.number(parent, name, -2147483648, 2147483647);
      if (value >= 0) margin[(["left", "right", "top", "bottom"] as const)[i]!] = value / 12700;
      else r.note(parent, "Negative table text insets have no nonnegative Core cell-margin mapping.");
    }
    const alignment = r.enumValue(parent, "anchor", { t: "top", ctr: "middle", b: "bottom" } as const);
    if (alignment) style.verticalAlign = alignment;
  }
  style.margin = margin;
  if (properties) {
    const direction = r.enumValue(properties, "vert", { horz: "horizontal", vert: "rotate90", vert270: "rotate270", eaVert: "eastAsianVertical", mongolianVert: "mongolianVertical", wordArtVert: "stackedVertical", wordArtVertRtl: "stackedVerticalRtl" } as const);
    if (direction !== undefined) style.textDirection = direction;
    const fill = sourceStyle ? await sourceStyle.paint(properties) : r.paint(properties);
    if (fill !== undefined) style.fill = fill;
    const merged = { ...style.borders, ...edges };
    const borders = [merged.left, merged.right, merged.top, merged.bottom];
    if (borders.every(border => border !== undefined) && borders.every(border => JSON.stringify(border) === JSON.stringify(borders[0]))) {
      style.stroke = { ...borders[0], paint: borders[0]!.paint ?? { type: "none" } };
      delete style.borders;
    } else if (borders.some(border => border !== undefined)) style.borders = merged;
  } else if (Object.keys(edges).length) {
    // A merge's anchor may lack tcPr while continuation outer edges are direct.
    style.borders = { ...style.borders, ...edges };
  }
  return style;
}
