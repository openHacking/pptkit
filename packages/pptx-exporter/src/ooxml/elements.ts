import type {
  ElementAction,
  NormalizedElement,
  NormalizedPlaceholderDefinition,
  NormalizedTableCell,
  NormalizedTextFrameStyle,
  NormalizedTextParagraph,
  PlaceholderKind,
} from "@pptkit/core";
import type { LayoutConnectorElement, LayoutElement, LayoutGroupElement } from "@pptkit/layout";
import type { ExportWarning } from "../types/export.js";
import { colorXml, emu, escapeXml, groupTransformXml, paintXml, paragraphAlignment, strokeXml, transformXml } from "./xml.js";

export interface ElementXmlContext {
  slideId?: string;
  warnings: ExportWarning[];
  placeholders: ReadonlyMap<string, NormalizedPlaceholderDefinition>;
  nextObjectId(): number;
  imageRelationship(assetId: string): string | undefined;
  actionRelationship(action: ElementAction): string | undefined;
}

function actionXml(action: ElementAction | undefined, context: ElementXmlContext): string {
  if (action === undefined) return "";
  const id = context.actionRelationship(action);
  if (id === undefined) return "";
  return `<a:hlinkClick r:id="${id}"${action.tooltip === undefined ? "" : ` tooltip="${escapeXml(action.tooltip)}"`}/>`;
}

function accessibility(element: LayoutElement): string {
  return element.accessibility.description ?? element.name;
}

function placeholderType(kind: PlaceholderKind): string {
  if (kind === "subtitle") return "subTitle";
  if (kind === "image") return "pic";
  if (kind === "slideNumber") return "sldNum";
  if (kind === "table") return "tbl";
  return kind;
}

function nonVisualProperties(element: LayoutElement, context: ElementXmlContext, kind: "shape" | "picture" | "connector" | "graphic"): string {
  const id = context.nextObjectId();
  const placeholder = element.placeholderKey === undefined ? undefined : context.placeholders.get(element.placeholderKey);
  const placeholderXml = placeholder === undefined ? "" : `<p:ph type="${placeholderType(placeholder.kind)}" idx="${[...context.placeholders.keys()].indexOf(placeholder.key) + 1}"/>`;
  const cNvPr = `<p:cNvPr id="${id}" name="${escapeXml(element.name)}" descr="${escapeXml(accessibility(element))}">${actionXml(element.action, context)}</p:cNvPr>`;
  if (kind === "picture") return `<p:nvPicPr>${cNvPr}<p:cNvPicPr preferRelativeResize="0"/><p:nvPr>${placeholderXml}</p:nvPr></p:nvPicPr>`;
  if (kind === "connector") return `<p:nvCxnSpPr>${cNvPr}<p:cNvCxnSpPr/><p:nvPr>${placeholderXml}</p:nvPr></p:nvCxnSpPr>`;
  if (kind === "graphic") return `<p:nvGraphicFramePr>${cNvPr}<p:cNvGraphicFramePr/><p:nvPr>${placeholderXml}</p:nvPr></p:nvGraphicFramePr>`;
  return `<p:nvSpPr>${cNvPr}<p:cNvSpPr${element.type === "text" ? ' txBox="1"' : ""}/><p:nvPr>${placeholderXml}</p:nvPr></p:nvSpPr>`;
}

function fontTypeface(value: NormalizedTextParagraph["runs"][number]["style"]["fontFamily"]): string {
  if (typeof value === "string") return value;
  return value.theme === "heading" ? "+mj-lt" : "+mn-lt";
}

function bulletXml(bullet: NormalizedTextParagraph["style"]["bullet"]): string {
  if (bullet.type === "none") return "<a:buNone/>";
  if (bullet.type === "bullet") return `<a:buChar char="${escapeXml(bullet.character)}"/>`;
  const mapping = { arabicPeriod: "arabicPeriod", arabicParen: "arabicParenR", alphaLowerPeriod: "alphaLcPeriod", alphaUpperPeriod: "alphaUcPeriod" } as const;
  return `<a:buAutoNum type="${mapping[bullet.style]}" startAt="${bullet.startAt}"/>`;
}

function runXml(run: NormalizedTextParagraph["runs"][number], context: ElementXmlContext, opacity: number): string {
  const style = run.style;
  const typeface = escapeXml(fontTypeface(style.fontFamily));
  const properties = `lang="${escapeXml(style.language)}" sz="${Math.round(style.fontSize * 100)}"${style.bold ? ' b="1"' : ""}${style.italic ? ' i="1"' : ""}${style.underline ? ' u="sng"' : ""}${style.strike ? ' strike="sngStrike"' : ""}`;
  return `<a:r><a:rPr ${properties}><a:solidFill>${colorXml(style.color, opacity)}</a:solidFill><a:latin typeface="${typeface}"/><a:ea typeface="${typeface}"/><a:cs typeface="${typeface}"/>${actionXml(run.action, context)}</a:rPr><a:t>${escapeXml(run.text)}</a:t></a:r>`;
}

export function textParagraphsXml(paragraphs: NormalizedTextParagraph[], context: ElementXmlContext, opacity = 1): string {
  return paragraphs.map((paragraph) => {
    const style = paragraph.style;
    const pPr = `<a:pPr algn="${paragraphAlignment(style.align)}" marL="${emu(style.indent)}" indent="-${emu(style.hanging)}"><a:lnSpc><a:spcPct val="${Math.round(style.lineSpacing * 100000)}"/></a:lnSpc><a:spcBef><a:spcPts val="${Math.round(style.spaceBefore * 100)}"/></a:spcBef><a:spcAft><a:spcPts val="${Math.round(style.spaceAfter * 100)}"/></a:spcAft>${bulletXml(style.bullet)}</a:pPr>`;
    return `<a:p>${pPr}${paragraph.runs.map((run) => runXml(run, context, opacity)).join("")}<a:endParaRPr lang="en-US"/></a:p>`;
  }).join("");
}

function bodyPropertiesXml(style: NormalizedTextFrameStyle): string {
  const anchor = style.verticalAlign === "middle" ? "ctr" : style.verticalAlign === "bottom" ? "b" : "t";
  const fit = style.autoFit.mode === "none" ? "<a:noAutofit/>" : style.autoFit.mode === "resize" ? "<a:spAutoFit/>" : `<a:normAutofit fontScale="${Math.round(style.autoFit.fontScale * 100000)}"/>`;
  return `<a:bodyPr wrap="${style.wrap ? "square" : "none"}" anchor="${anchor}" lIns="${emu(style.margin.left)}" tIns="${emu(style.margin.top)}" rIns="${emu(style.margin.right)}" bIns="${emu(style.margin.bottom)}">${fit}</a:bodyPr>`;
}

function textBodyXml(content: NormalizedTextParagraph[], frame: NormalizedTextFrameStyle, context: ElementXmlContext, opacity: number): string {
  return `<p:txBody>${bodyPropertiesXml(frame)}<a:lstStyle/>${textParagraphsXml(content, context, opacity)}</p:txBody>`;
}

function textXml(element: Extract<LayoutElement, { type: "text" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const opacity = inheritedOpacity * element.opacity;
  return `<p:sp>${nonVisualProperties(element, context, "shape")}<p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>${textBodyXml(element.content, element.frame, context, opacity)}</p:sp>`;
}

function imageXml(element: Extract<LayoutElement, { type: "image" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const relationshipId = context.imageRelationship(element.assetId);
  if (relationshipId === undefined) return "";
  const crop = element.fit === "crop" ? `<a:srcRect l="${Math.round(element.crop.left * 100000)}" t="${Math.round(element.crop.top * 100000)}" r="${Math.round(element.crop.right * 100000)}" b="${Math.round(element.crop.bottom * 100000)}"/>` : "";
  const alpha = inheritedOpacity * element.opacity < 1 ? `<a:alphaModFix amt="${Math.round(inheritedOpacity * element.opacity * 100000)}"/>` : "";
  return `<p:pic>${nonVisualProperties(element, context, "picture")}<p:blipFill><a:blip r:embed="${relationshipId}">${alpha}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function shapeGeometry(shape: Extract<LayoutElement, { type: "shape" }>["shape"]): string {
  return { rect: "rect", roundRect: "roundRect", ellipse: "ellipse", triangle: "triangle", diamond: "diamond", arrow: "rightArrow", chevron: "chevron" }[shape];
}

function shapeXml(element: Extract<LayoutElement, { type: "shape" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const opacity = inheritedOpacity * element.opacity;
  const text = element.text === undefined ? "" : textBodyXml(element.text.content, element.text.frame, context, opacity);
  return `<p:sp>${nonVisualProperties(element, context, "shape")}<p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="${shapeGeometry(element.shape)}"><a:avLst/></a:prstGeom>${paintXml(element.style.fill, opacity)}${strokeXml(element.style.stroke, opacity)}</p:spPr>${text}</p:sp>`;
}

function connectorXml(element: LayoutConnectorElement, context: ElementXmlContext, inheritedOpacity: number): string {
  const points = [element.start, ...element.route, element.end];
  const width = Math.max(element.box.width, 0.001);
  const height = Math.max(element.box.height, 0.001);
  // DrawingML requires a positive shape extent. A horizontal or vertical
  // connector can legitimately have a zero-sized bounding-box dimension,
  // but emitting cy/cx="0" makes PowerPoint repair the package on open.
  const transformBox = {
    ...element.box,
    width,
    height,
  };
  const path = points.map((point, index) => {
    const x = emu(point.x - element.box.x);
    const y = emu(point.y - element.box.y);
    return index === 0 ? `<a:moveTo><a:pt x="${x}" y="${y}"/></a:moveTo>` : `<a:lnTo><a:pt x="${x}" y="${y}"/></a:lnTo>`;
  }).join("");
  const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="${emu(width)}" h="${emu(height)}" fill="none">${path}</a:path></a:pathLst></a:custGeom>`;
  // Custom geometry is valid on a regular shape but is not accepted inside
  // PowerPoint's connector container. Keep the connector's resolved path and
  // styling, while using the interoperable shape container for serialization.
  return `<p:sp>${nonVisualProperties(element, context, "shape")}<p:spPr>${transformXml(transformBox, element.transform)}${geometry}${strokeXml(element.style, inheritedOpacity * element.opacity)}</p:spPr></p:sp>`;
}

function groupXml(element: LayoutGroupElement, context: ElementXmlContext, inheritedOpacity: number): string {
  const id = context.nextObjectId();
  const opacity = inheritedOpacity * element.opacity;
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapeXml(element.name)}" descr="${escapeXml(accessibility(element))}">${actionXml(element.action, context)}</p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml(element.box, element.transform, element.coordinateSize.width, element.coordinateSize.height)}</p:grpSpPr>${element.children.map((child) => elementXml(child, context, opacity)).join("")}</p:grpSp>`;
}

function emptyTableCell(attributes: string): string {
  return `<a:tc${attributes}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>`;
}

function tableCellXml(cell: NormalizedTableCell, context: ElementXmlContext, opacity: number, attributes = ""): string {
  const style = cell.style;
  const border = ["lnL", "lnR", "lnT", "lnB"].map((tag) => strokeXml(style.stroke, opacity, `a:${tag}`)).join("");
  const anchor = style.verticalAlign === "middle" ? "ctr" : style.verticalAlign === "bottom" ? "b" : "t";
  return `<a:tc${attributes}><a:txBody><a:bodyPr lIns="${emu(style.margin.left)}" tIns="${emu(style.margin.top)}" rIns="${emu(style.margin.right)}" bIns="${emu(style.margin.bottom)}" anchor="${anchor}"/><a:lstStyle/>${textParagraphsXml(cell.content, context, opacity)}</a:txBody><a:tcPr>${paintXml(style.fill, opacity)}${border}</a:tcPr></a:tc>`;
}

function tableRowsXml(element: Extract<LayoutElement, { type: "table" }>, context: ElementXmlContext, opacity: number): string {
  const active = new Map<number, { remaining: number; horizontalOffset: number }>();
  return element.rows.map((row) => {
    const cells: string[] = [];
    let inputIndex = 0;
    for (let column = 0; column < element.columns.length; column += 1) {
      const merge = active.get(column);
      if (merge !== undefined) {
        cells.push(emptyTableCell(` vMerge="1"${merge.horizontalOffset > 0 ? ' hMerge="1"' : ""}`));
        merge.remaining -= 1;
        if (merge.remaining === 0) active.delete(column);
        continue;
      }
      const cell = row.cells[inputIndex++];
      if (cell === undefined) {
        cells.push(emptyTableCell(""));
        continue;
      }
      const attributes = `${cell.colSpan > 1 ? ` gridSpan="${cell.colSpan}"` : ""}${cell.rowSpan > 1 ? ` rowSpan="${cell.rowSpan}"` : ""}`;
      cells.push(tableCellXml(cell, context, opacity, attributes));
      for (let offset = 0; offset < cell.colSpan; offset += 1) {
        if (cell.rowSpan > 1) active.set(column + offset, { remaining: cell.rowSpan - 1, horizontalOffset: offset });
        if (offset > 0) {
          column += 1;
          cells.push(emptyTableCell(' hMerge="1"'));
        }
      }
    }
    return `<a:tr h="${emu(row.height)}">${cells.join("")}</a:tr>`;
  }).join("");
}

function tableXml(element: Extract<LayoutElement, { type: "table" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const opacity = inheritedOpacity * element.opacity;
  return `<p:graphicFrame>${nonVisualProperties(element, context, "graphic")}<p:xfrm><a:off x="${emu(element.box.x)}" y="${emu(element.box.y)}"/><a:ext cx="${emu(element.box.width)}" cy="${emu(element.box.height)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${element.columns.map((width) => `<a:gridCol w="${emu(width)}"/>`).join("")}</a:tblGrid>${tableRowsXml(element, context, opacity)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

export function elementXml(element: LayoutElement, context: ElementXmlContext, inheritedOpacity = 1): string {
  if (element.hidden) return "";
  if (element.type === "text") return textXml(element, context, inheritedOpacity);
  if (element.type === "image") return imageXml(element, context, inheritedOpacity);
  if (element.type === "shape") return shapeXml(element, context, inheritedOpacity);
  if (element.type === "connector") return connectorXml(element, context, inheritedOpacity);
  if (element.type === "group") return groupXml(element, context, inheritedOpacity);
  if (element.type === "table") return tableXml(element, context, inheritedOpacity);
  context.warnings.push({ code: "unsupported-element", message: `Unsupported element type was omitted.`, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}) });
  return "";
}

export function placeholderShapeXml(placeholder: NormalizedPlaceholderDefinition, index: number, context: ElementXmlContext): string {
  const id = context.nextObjectId();
  const type = placeholderType(placeholder.kind);
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(placeholder.key)}"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="${type}" idx="${index + 1}"/></p:nvPr></p:nvSpPr><p:spPr>${transformXml(placeholder.box, { rotation: 0, flipH: false, flipV: false })}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody>${bodyPropertiesXml(placeholder.textStyle.frame)}<a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}
