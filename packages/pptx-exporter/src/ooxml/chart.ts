import type { NormalizedChartTextStyle, NormalizedPaint, NormalizedStrokeStyle } from "@pptkit/core";
import type { LayoutChartElement } from "@pptkit/layout";
import { colorValue, escapeXml, paintXml, strokeXml } from "./xml.js";
import { spreadsheetColumn } from "./chart-workbook.js";

function ratio(value: number, total: number): string {
  return (total <= 0 ? 0 : Math.max(0, Math.min(1, value / total))).toFixed(6).replace(/0+$/, "").replace(/\.$/, "");
}

function textProperties(style: NormalizedChartTextStyle): string {
  return `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round(style.fontSize * 100)}" b="${style.bold ? 1 : 0}" i="${style.italic ? 1 : 0}"><a:solidFill><a:srgbClr val="${colorValue(style.color)}"/></a:solidFill><a:latin typeface="${escapeXml(style.fontFamily)}"/><a:ea typeface="${escapeXml(style.fontFamily)}"/><a:cs typeface="${escapeXml(style.fontFamily)}"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

function shapeProperties(fill: NormalizedPaint, stroke?: NormalizedStrokeStyle): string {
  return `<c:spPr>${paintXml(fill)}${stroke === undefined ? "<a:ln><a:noFill/></a:ln>" : strokeXml(stroke)}</c:spPr>`;
}

function manualLayout(chart: LayoutChartElement, box: { x: number; y: number; width: number; height: number }): string {
  return `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="${ratio(box.x, chart.box.width)}"/><c:y val="${ratio(box.y, chart.box.height)}"/><c:w val="${ratio(box.width, chart.box.width)}"/><c:h val="${ratio(box.height, chart.box.height)}"/></c:manualLayout></c:layout>`;
}

function stringReference(formula: string, values: string[]): string {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`).join("");
  return `<c:strRef><c:f>${formula}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

function numberReference(formula: string, values: number[]): string {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("");
  return `<c:numRef><c:f>${formula}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`;
}

function markerXml(chart: Extract<LayoutChartElement, { chartType: "line" }>, seriesIndex: number): string {
  const marker = chart.series[seriesIndex]?.marker;
  if (marker === false || marker === undefined) return `<c:marker><c:symbol val="none"/></c:marker>`;
  const shape = marker.shape === "plus" ? "plus" : marker.shape;
  return `<c:marker><c:symbol val="${shape}"/><c:size val="${Math.max(2, Math.min(72, Math.round(marker.size)))}"/>${shapeProperties(marker.fill, marker.stroke)}</c:marker>`;
}

function seriesXml(chart: LayoutChartElement): string {
  const lastRow = chart.categories.length + 1;
  return chart.series.map((series, seriesIndex) => {
    const column = spreadsheetColumn(seriesIndex + 2);
    const title = `<c:tx>${stringReference(`Sheet1!$${column}$1`, [series.name])}</c:tx>`;
    const categories = `<c:cat>${stringReference(`Sheet1!$A$2:$A$${lastRow}`, chart.categories)}</c:cat>`;
    const values = `<c:val>${numberReference(`Sheet1!$${column}$2:$${column}$${lastRow}`, series.values)}</c:val>`;
    let style: string;
    let postData = "";
    if (chart.chartType === "line") {
      const lineSeries = chart.series[seriesIndex];
      style = `${lineSeries?.line === false || lineSeries === undefined ? "" : `<c:spPr>${strokeXml(lineSeries.line)}</c:spPr>`}${markerXml(chart, seriesIndex)}`;
      // CT_LineSer requires smooth after the category and value references.
      // PowerPoint repairs the presentation when this node appears before them.
      postData = `<c:smooth val="0"/>`;
    } else if (chart.chartType === "pie") {
      style = (chart.series[seriesIndex]?.pointColors ?? []).map((pointColor, pointIndex) => `<c:dPt><c:idx val="${pointIndex}"/>${shapeProperties({ type: "solid", color: pointColor, opacity: 1 })}</c:dPt>`).join("");
    } else {
      style = shapeProperties({ type: "solid", color: series.color, opacity: 1 });
    }
    return `<c:ser><c:idx val="${seriesIndex}"/><c:order val="${seriesIndex}"/>${title}${style}${categories}${values}${postData}</c:ser>`;
  }).join("");
}

function tickMark(value: "none" | "inside" | "outside" | "cross"): string {
  if (value === "inside") return "in";
  if (value === "outside") return "out";
  return value;
}

function hiddenAxisStroke(stroke: NormalizedStrokeStyle, visible: boolean): NormalizedStrokeStyle {
  return visible ? stroke : { ...stroke, paint: { type: "none" } };
}

function axesXml(chart: Exclude<LayoutChartElement, { chartType: "pie" }>): string {
  const horizontal = chart.chartType === "bar" && chart.orientation === "horizontal";
  const category = chart.axes.category;
  const value = chart.axes.value;
  const scale = chart.chartLayout.valueScale;
  const categoryGridlines = category.majorGridlines === false ? "" : `<c:majorGridlines><c:spPr>${strokeXml(category.majorGridlines)}</c:spPr></c:majorGridlines>`;
  const gridlines = value.majorGridlines === false ? "" : `<c:majorGridlines><c:spPr>${strokeXml(value.majorGridlines)}</c:spPr></c:majorGridlines>`;
  const catAxis = `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="${horizontal ? "maxMin" : "minMax"}"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? "l" : "b"}"/>${categoryGridlines}<c:majorTickMark val="${category.visible ? tickMark(category.majorTick) : "none"}"/><c:minorTickMark val="none"/><c:tickLblPos val="${category.visible && category.labels ? "nextTo" : "none"}"/><c:spPr>${strokeXml(hiddenAxisStroke(category.line, category.visible))}</c:spPr>${textProperties(category.labelStyle)}<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const crossBetween = chart.chartLayout.categoryAxisPlacement === "onCategory" ? "midCat" : "between";
  const valAxis = `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:max val="${scale.max}"/><c:min val="${scale.min}"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? "b" : "l"}"/>${gridlines}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="${value.visible ? tickMark(value.majorTick) : "none"}"/><c:minorTickMark val="none"/><c:tickLblPos val="${value.visible && value.labels ? "nextTo" : "none"}"/><c:spPr>${strokeXml(hiddenAxisStroke(value.line, value.visible))}</c:spPr>${textProperties(value.labelStyle)}<c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="${crossBetween}"/><c:majorUnit val="${scale.majorUnit}"/></c:valAx>`;
  return `${catAxis}${valAxis}`;
}

function titleXml(chart: LayoutChartElement): string {
  if (chart.title === undefined) return "";
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="${Math.round(chart.titleStyle.fontSize * 100)}" b="${chart.titleStyle.bold ? 1 : 0}" i="${chart.titleStyle.italic ? 1 : 0}"><a:solidFill><a:srgbClr val="${colorValue(chart.titleStyle.color)}"/></a:solidFill><a:latin typeface="${escapeXml(chart.titleStyle.fontFamily)}"/></a:rPr><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

function legendXml(chart: LayoutChartElement): string {
  if (!chart.legend.visible) return "";
  const position = chart.legend.position === "bottom" ? "b" : "r";
  return `<c:legend><c:legendPos val="${position}"/><c:overlay val="0"/>${textProperties(chart.legend.textStyle)}</c:legend>`;
}

export function chartPartXml(chart: LayoutChartElement): string {
  const series = seriesXml(chart);
  let chartBody: string;
  if (chart.chartType === "bar") {
    const grouping = chart.grouping === "percentStacked" ? "percentStacked" : chart.grouping;
    const overlap = chart.grouping === "clustered" ? -Math.round(chart.seriesGap) : 100;
    chartBody = `<c:barChart><c:barDir val="${chart.orientation === "horizontal" ? "bar" : "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}<c:gapWidth val="${Math.round(chart.categoryGap)}"/><c:overlap val="${overlap}"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axesXml(chart)}`;
  } else if (chart.chartType === "line") {
    chartBody = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/><c:smooth val="0"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axesXml(chart)}`;
  } else {
    chartBody = `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  const plot = `<c:plotArea>${manualLayout(chart, chart.chartLayout.plotBox)}${chartBody}${shapeProperties(chart.style.plotArea)}</c:plotArea>`;
  const chartXml = `<c:chart>${titleXml(chart)}${plot}${legendXml(chart)}<c:plotVisOnly val="1"/><c:dispBlanksAs val="zero"/><c:showDLblsOverMax val="0"/></c:chart>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>${chartXml}${shapeProperties(chart.style.chartArea)}<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}
