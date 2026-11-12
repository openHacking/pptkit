import type { NormalizedChartMarkerStyle, NormalizedChartTextStyle, NormalizedPaint, NormalizedStrokeStyle } from "@pptkit/core";
import type { LayoutChartElement } from "@pptkit/layout";
import { colorValue, escapeXml, paintXml, strokeXml } from "./xml.js";
import { spreadsheetColumn } from "./chart-workbook.js";
import { chartStringReference, chartNumberReference } from "./chart-reference.js";
import { scatterSeriesXml } from "./scatter-series.js";

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

function markerXml(marker: false | NormalizedChartMarkerStyle | undefined): string {
  if (marker === false || marker === undefined) return `<c:marker><c:symbol val="none"/></c:marker>`;
  const shape = marker.shape === "plus" ? "plus" : marker.shape;
  return `<c:marker><c:symbol val="${shape}"/><c:size val="${Math.max(2, Math.min(72, Math.round(marker.size)))}"/>${shapeProperties(marker.fill, marker.stroke)}</c:marker>`;
}

function seriesXml(chart: Exclude<LayoutChartElement, { chartType: "scatter" | "bubble" }>, comboKind?: "bar" | "line"): string {
  const lastRow = chart.categories.length + 1;
  return chart.series.map((series, seriesIndex) => {
    if (comboKind && (!("kind" in series) || series.kind !== comboKind)) return "";
    const column = spreadsheetColumn(seriesIndex + 2);
    const title = `<c:tx>${chartStringReference(`Sheet1!$${column}$1`, [series.name])}</c:tx>`;
    const categories = `<c:cat>${chartStringReference(`Sheet1!$A$2:$A$${lastRow}`, chart.categories)}</c:cat>`;
    const values = `<c:val>${chartNumberReference(`Sheet1!$${column}$2:$${column}$${lastRow}`, series.values)}</c:val>`;
    let style: string;
    let postData = "";
    if ("line" in series) {
      style = `${series.line === false ? "" : `<c:spPr>${strokeXml(series.line)}</c:spPr>`}${markerXml(series.marker)}`;
      // CT_LineSer requires smooth after the category and value references.
      // PowerPoint repairs the presentation when this node appears before them.
      postData = `<c:smooth val="0"/>`;
    } else if ((chart.chartType === "pie" || chart.chartType === "doughnut")) {
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

function axesXml(chart: Exclude<LayoutChartElement, { chartType: "pie" | "doughnut" | "scatter" | "bubble" }>, secondary = false): string {
  const horizontal = chart.chartType === "bar" && chart.orientation === "horizontal";
  const category = chart.axes.category;
  const value = secondary && chart.chartType === "combo" ? chart.axes.secondaryValue : chart.axes.value;
  const scale = secondary ? chart.chartLayout.secondaryValueScale! : chart.chartLayout.valueScale;
  const categoryId = secondary ? 3 : 1, valueId = secondary ? 4 : 2;
  const categoryGridlines = secondary || category.majorGridlines === false ? "" : `<c:majorGridlines><c:spPr>${strokeXml(category.majorGridlines)}</c:spPr></c:majorGridlines>`;
  const gridlines = value.majorGridlines === false ? "" : `<c:majorGridlines><c:spPr>${strokeXml(value.majorGridlines)}</c:spPr></c:majorGridlines>`;
  const catAxis = `<c:catAx><c:axId val="${categoryId}"/><c:scaling><c:orientation val="${horizontal ? "maxMin" : "minMax"}"/></c:scaling><c:delete val="${secondary ? 1 : 0}"/><c:axPos val="${secondary ? "t" : horizontal ? "l" : "b"}"/>${categoryGridlines}<c:majorTickMark val="${category.visible ? tickMark(category.majorTick) : "none"}"/><c:minorTickMark val="none"/><c:tickLblPos val="${category.visible && category.labels ? "nextTo" : "none"}"/><c:spPr>${strokeXml(hiddenAxisStroke(category.line, category.visible))}</c:spPr>${textProperties(category.labelStyle)}<c:crossAx val="${valueId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const crossBetween = chart.chartLayout.categoryAxisPlacement === "onCategory" ? "midCat" : "between";
  const valAxis = `<c:valAx><c:axId val="${valueId}"/><c:scaling><c:orientation val="minMax"/><c:max val="${scale.max}"/><c:min val="${scale.min}"/></c:scaling><c:delete val="0"/><c:axPos val="${secondary ? "r" : horizontal ? "b" : "l"}"/>${gridlines}<c:numFmt formatCode="${chart.chartType === "area" && chart.grouping === "percentStacked" ? "0%" : "General"}" sourceLinked="${chart.chartType === "area" && chart.grouping === "percentStacked" ? "0" : "1"}"/><c:majorTickMark val="${value.visible ? tickMark(value.majorTick) : "none"}"/><c:minorTickMark val="none"/><c:tickLblPos val="${value.visible && value.labels ? "nextTo" : "none"}"/><c:spPr>${strokeXml(hiddenAxisStroke(value.line, value.visible))}</c:spPr>${textProperties(value.labelStyle)}<c:crossAx val="${categoryId}"/><c:crosses val="${secondary ? "max" : "autoZero"}"/><c:crossBetween val="${crossBetween}"/><c:majorUnit val="${scale.majorUnit}"/></c:valAx>`;
  return `${catAxis}${valAxis}`;
}

function scatterAxesXml(chart: Extract<LayoutChartElement, { chartType: "scatter" | "bubble" }>): string {
  const serialize = (id: number, cross: number, position: "b" | "l", axis: typeof chart.axes.x, scale: typeof chart.chartLayout.valueScale) => {
    const grid = axis.majorGridlines === false ? "" : `<c:majorGridlines><c:spPr>${strokeXml(axis.majorGridlines)}</c:spPr></c:majorGridlines>`;
    return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/><c:max val="${scale.max}"/><c:min val="${scale.min}"/></c:scaling><c:delete val="0"/><c:axPos val="${position}"/>${grid}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="${axis.visible ? tickMark(axis.majorTick) : "none"}"/><c:minorTickMark val="none"/><c:tickLblPos val="${axis.visible && axis.labels ? "nextTo" : "none"}"/><c:spPr>${strokeXml(hiddenAxisStroke(axis.line,axis.visible))}</c:spPr>${textProperties(axis.labelStyle)}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/><c:majorUnit val="${scale.majorUnit}"/></c:valAx>`;
  };
  return serialize(1,2,"b",chart.axes.x,chart.chartLayout.xScale!)+serialize(2,1,"l",chart.axes.y,chart.chartLayout.valueScale);
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

export function chartPartXml(chart: LayoutChartElement, bubbleSeries?: string): string {
  if (chart.chartType === "bubble" && bubbleSeries === undefined) throw new Error("Bubble series must be prepared with its workbook.");
  const series = chart.chartType === "combo" ? "" : chart.chartType === "bubble" ? bubbleSeries : chart.chartType === "scatter" ? scatterSeriesXml(chart.series) : seriesXml(chart);
  let chartBody: string;
  if (chart.chartType === "scatter") {
    chartBody = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>${scatterAxesXml(chart)}`;
  } else if (chart.chartType === "bubble") {
    chartBody = `<c:bubbleChart><c:varyColors val="0"/>${series}<c:bubble3D val="0"/><c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>${scatterAxesXml(chart)}`;
  } else if (chart.chartType === "combo") {
    chartBody = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml(chart, "bar")}<c:gapWidth val="150"/><c:overlap val="0"/><c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml(chart, "line")}<c:marker val="1"/><c:smooth val="0"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>${axesXml(chart)}${axesXml(chart, true)}`;
  } else if (chart.chartType === "bar") {
    const grouping = chart.grouping === "percentStacked" ? "percentStacked" : chart.grouping;
    const overlap = chart.grouping === "clustered" ? -Math.round(chart.seriesGap) : 100;
    chartBody = `<c:barChart><c:barDir val="${chart.orientation === "horizontal" ? "bar" : "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}<c:gapWidth val="${Math.round(chart.categoryGap)}"/><c:overlap val="${overlap}"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axesXml(chart)}`;
  } else if (chart.chartType === "line") {
    chartBody = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/><c:smooth val="0"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axesXml(chart)}`;
  } else if (chart.chartType === "area") {
    chartBody = `<c:areaChart><c:grouping val="${chart.grouping}"/><c:varyColors val="0"/>${series}<c:axId val="1"/><c:axId val="2"/></c:areaChart>${axesXml(chart)}`;
  } else if (chart.chartType === "doughnut") {
    chartBody = `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="${chart.firstSliceAngle}"/><c:holeSize val="${chart.holeSize}"/></c:doughnutChart>`;
  } else {
    chartBody = `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  const plot = `<c:plotArea>${manualLayout(chart, chart.chartLayout.plotBox)}${chartBody}${shapeProperties(chart.style.plotArea)}</c:plotArea>`;
  const blanks = chart.chartType === "bubble" ? "zero" : chart.blankDisplay ?? ((chart.chartType === "scatter" ? chart.series.some(series => series.xValues.includes(null) || series.yValues.includes(null)) : chart.series.some(series => series.values.includes(null))) ? "gap" : "zero");
  const chartXml = `<c:chart>${titleXml(chart)}${plot}${legendXml(chart)}<c:plotVisOnly val="1"/><c:dispBlanksAs val="${blanks}"/><c:showDLblsOverMax val="0"/></c:chart>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>${chartXml}${shapeProperties(chart.style.chartArea)}<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}
