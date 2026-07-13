import type { NormalizedElement } from "@pptkit/core";
import type { ExportWarning } from "../types/export.js";
import { colorValue, emu, escapeXml, paragraphAlignment } from "./xml.js";

export function elementXml(element: NormalizedElement, index: number, warnings: ExportWarning[], slideId: string): string {
  if (element.type === "text") return textXml(element, index);
  if (element.type === "shape") return shapeXml(element, index);
  warnings.push({ code: "unsupported-element", message: `Unsupported element type "${element.type}" was omitted.`, slideId, elementIndex: index });
  return "";
}

export function imageXml(element: Extract<NormalizedElement, { type: "image" }>, relationshipId: string, index: number): string {
  const name = escapeXml(element.altText ?? `Image ${index + 1}`);
  return `<p:pic><p:nvPicPr><p:cNvPr id="${index + 2}" name="${name}" descr="${name}"/><p:cNvPicPr preferRelativeResize="0"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${emu(element.box.x)}" y="${emu(element.box.y)}"/><a:ext cx="${emu(element.box.width)}" cy="${emu(element.box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function textXml(element: Extract<NormalizedElement, { type: "text" }>, index: number): string {
  const style = element.style;
  const paragraphs = element.text.split(/\r\n|\r|\n/).map((text) => paragraphXml(text, style));
  return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="Text ${index + 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(element.box.x)}" y="${emu(element.box.y)}"/><a:ext cx="${emu(element.box.width)}" cy="${emu(element.box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody>${bodyPropertiesXml(style)}<a:lstStyle/>${paragraphs.join("")}</p:txBody></p:sp>`;
}

function bodyPropertiesXml(style: Extract<NormalizedElement, { type: "text" }>["style"]): string {
  if (style.autoFit === undefined) return "<a:bodyPr/>";
  if (style.autoFit.mode === "none") {
    return '<a:bodyPr wrap="square" anchor="t"><a:noAutofit/></a:bodyPr>';
  }
  const fontScale = style.autoFit.fontScale === undefined
    ? ""
    : ` fontScale="${Math.round(style.autoFit.fontScale * 100000)}"`;
  return `<a:bodyPr wrap="square" anchor="t"><a:normAutofit${fontScale}/></a:bodyPr>`;
}

function paragraphXml(text: string, style: Extract<NormalizedElement, { type: "text" }>["style"]): string {
  const lineSpacing = style.lineSpacing === undefined
    ? ""
    : `<a:lnSpc><a:spcPct val="${Math.round(style.lineSpacing * 100000)}"/></a:lnSpc>`;
  const fontFamily = style.fontFamily === undefined
    ? ""
    : `<a:latin typeface="${escapeXml(style.fontFamily)}"/><a:ea typeface="${escapeXml(style.fontFamily)}"/><a:cs typeface="${escapeXml(style.fontFamily)}"/>`;
  return `<a:p><a:pPr algn="${paragraphAlignment(style.align)}">${lineSpacing}</a:pPr><a:r><a:rPr lang="en-US" sz="${Math.round((style.fontSize ?? 18) * 100)}"${style.fontWeight === "bold" ? " b=\"1\"" : ""}><a:solidFill><a:srgbClr val="${colorValue(style.color ?? "000000")}"/></a:solidFill>${fontFamily}</a:rPr><a:t>${escapeXml(text)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>`;
}

function shapeXml(element: Extract<NormalizedElement, { type: "shape" }>, index: number): string {
  const style = element.style;
  const geometry = element.shape === "ellipse" ? "ellipse" : element.shape === "line" ? "line" : "rect";
  const fill = style.fill ? `<a:solidFill><a:srgbClr val="${colorValue(style.fill)}"/></a:solidFill>` : "<a:noFill/>";
  const stroke = style.stroke ? `<a:ln w="${emu(style.strokeWidth ?? 1)}"><a:solidFill><a:srgbClr val="${colorValue(style.stroke)}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="Shape ${index + 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(element.box.x)}" y="${emu(element.box.y)}"/><a:ext cx="${emu(element.box.width)}" cy="${emu(element.box.height)}"/></a:xfrm><a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>${fill}${stroke}</p:spPr></p:sp>`;
}
