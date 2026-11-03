import {
  pxToEmu,
  resolvedColorXml,
  xmlEscape,
  type ColorValue,
  type ThemeContext,
} from "@react-pptx/core";
import type { NativeLink } from "@react-pptx/core";
import { strToU8, zipSync } from "fflate";

import type {
  ChartAxis,
  ChartData,
  ChartDataLabels,
  ChartKind,
  ChartLine,
  ChartPoint,
  ChartProps,
  ChartSeries,
} from "./chart.ts";

const chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const drawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const categoryAxisId = 111111;
const valueAxisId = 222222;
const sheet = "Sheet1";

type XmlAttributes = Readonly<
  Record<string, string | number | boolean | undefined>
>;
type XmlNode = XmlElement | XmlRaw | string;
type Child = XmlNode | undefined | false | null | readonly Child[];
type XmlRaw = { readonly raw: string };
type XmlElement = {
  readonly name: string;
  readonly attributes: XmlAttributes;
  readonly children: readonly XmlNode[];
};

function flatten(children: readonly Child[], into: XmlNode[]): XmlNode[] {
  for (const child of children) {
    if (child === undefined || child === false || child === null) continue;
    if (typeof child === "string" || !Array.isArray(child))
      into.push(child as XmlNode);
    else flatten(child, into);
  }
  return into;
}

function el(
  name: string,
  attributes: XmlAttributes | null = null,
  ...children: Child[]
): XmlElement {
  return {
    name,
    attributes: attributes ?? {},
    children: flatten(children, []),
  };
}

function raw(value: string | undefined): XmlRaw | undefined {
  return value === undefined ? undefined : { raw: value };
}

function write(node: XmlNode, into: string[]): void {
  if (typeof node === "string") {
    into.push(xmlEscape(node));
    return;
  }
  if ("raw" in node) {
    into.push(node.raw);
    return;
  }
  into.push("<", node.name);
  for (const [key, value] of Object.entries(node.attributes)) {
    if (value === undefined) continue;
    const rendered =
      typeof value === "boolean" ? (value ? "1" : "0") : String(value);
    into.push(" ", key, '="', xmlEscape(rendered), '"');
  }
  if (node.children.length === 0) {
    into.push("/>");
    return;
  }
  into.push(">");
  for (const child of node.children) write(child, into);
  into.push("</", node.name, ">");
}

function serialize(root: XmlElement): string {
  const into = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'];
  write(root, into);
  return into.join("");
}

function serializeFragment(root: XmlElement): string {
  const into: string[] = [];
  write(root, into);
  return into.join("");
}

function columnName(index: number): string {
  let name = "";
  for (let rest = index + 1; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    name = String.fromCharCode(65 + ((rest - 1) % 26)) + name;
  }
  return name;
}

function strRef(formula: string, values: readonly string[]): XmlElement {
  return el(
    "c:strRef",
    null,
    el("c:f", null, formula),
    el(
      "c:strCache",
      null,
      el("c:ptCount", { val: values.length }),
      values.map((value, index) =>
        el("c:pt", { idx: index }, el("c:v", null, value)),
      ),
    ),
  );
}

function numRef(
  formula: string,
  values: readonly (number | null)[],
): XmlElement {
  return el(
    "c:numRef",
    null,
    el("c:f", null, formula),
    el(
      "c:numCache",
      null,
      el("c:formatCode", null, "General"),
      el("c:ptCount", { val: values.length }),
      values.flatMap((value, index) =>
        value === null
          ? []
          : [el("c:pt", { idx: index }, el("c:v", null, String(value)))],
      ),
    ),
  );
}

function color(value: ColorValue, theme: ThemeContext): XmlRaw {
  return { raw: resolvedColorXml(theme.resolveColor(value)) };
}

function chartLine(
  value: ChartLine | undefined,
  theme: ThemeContext,
): XmlElement | undefined {
  if (value === undefined) return undefined;
  return el(
    "a:ln",
    { w: pxToEmu(value.width ?? 1) },
    el("a:solidFill", null, color(value.color ?? "dk1", theme)),
  );
}

function shapeProperties(
  value: "none" | ChartLine | undefined,
  theme: ThemeContext,
): XmlElement | undefined {
  if (value === undefined) return undefined;
  return el(
    "c:spPr",
    null,
    value === "none"
      ? el("a:ln", null, el("a:noFill"))
      : chartLine(value, theme),
  );
}

function transparent(): XmlElement {
  return el("c:spPr", null, el("a:noFill"), el("a:ln", null, el("a:noFill")));
}

export type ChartTextStyleXml = {
  readonly chart: string;
  readonly categoryAxis?: string;
  readonly valueAxis?: string;
  readonly dataLabels?: string;
  readonly pointLabels?: ReadonlyMap<string, string>;
};

function dataLabels(
  kind: ChartKind,
  labels: ChartDataLabels | false | undefined,
  series: ChartSeries,
  seriesIndex: number,
  textStyles: ChartTextStyleXml,
): XmlElement | undefined {
  if (labels === undefined || labels === false) return undefined;
  const positions = {
    center: "ctr",
    insideEnd: "inEnd",
    insideBase: "inBase",
    outsideEnd: "outEnd",
    bestFit: "bestFit",
  } as const;
  const hasPosition =
    kind === "column" ||
    kind === "bar" ||
    kind === "pie" ||
    kind === "doughnut";
  const settings = (textStyle: string | undefined) => [
    labels.numberFormat === undefined
      ? undefined
      : el("c:numFmt", {
          formatCode: labels.numberFormat,
          sourceLinked: false,
        }),
    transparent(),
    raw(textStyle),
    labels.position === undefined || !hasPosition
      ? undefined
      : el("c:dLblPos", { val: positions[labels.position] }),
    el("c:showLegendKey", { val: false }),
    el("c:showVal", { val: labels.value ?? false }),
    el("c:showCatName", { val: labels.category ?? false }),
    el("c:showSerName", { val: false }),
    el("c:showPercent", { val: labels.percent ?? false }),
    el("c:showBubbleSize", { val: false }),
  ];
  const pointLabels = Object.entries<ChartPoint>(series.points ?? {}).flatMap(
    ([pointIndex, point]) => {
      const style = textStyles.pointLabels?.get(`${seriesIndex}:${pointIndex}`);
      if (point.labelColor === undefined || style === undefined) return [];
      return [
        el(
          "c:dLbl",
          null,
          el("c:idx", { val: Number(pointIndex) }),
          settings(style),
        ),
      ];
    },
  );
  return el("c:dLbls", null, pointLabels, settings(textStyles.dataLabels));
}

function chartSeries(
  kind: ChartKind,
  data: ChartData,
  item: ChartSeries,
  index: number,
  labels: ChartDataLabels | false | undefined,
  textStyles: ChartTextStyleXml,
  explosion: number | undefined,
  theme: ThemeContext,
): XmlElement {
  const rows = data.categories.length + 1;
  const column = columnName(index + 1);
  const bar = kind === "column" || kind === "bar";
  const lineChart = kind === "line";
  const pie = kind === "pie" || kind === "doughnut";
  const itemShape =
    item.color === undefined
      ? undefined
      : lineChart
        ? el(
            "c:spPr",
            null,
            el("a:ln", null, el("a:solidFill", null, color(item.color, theme))),
          )
        : el("c:spPr", null, el("a:solidFill", null, color(item.color, theme)));
  const points = Object.entries<ChartPoint>(item.points ?? {}).flatMap(
    ([pointIndex, point]) => {
      if (point.color === undefined) return [];
      return [
        el(
          "c:dPt",
          null,
          el("c:idx", { val: Number(pointIndex) }),
          bar ? el("c:invertIfNegative", { val: false }) : undefined,
          bar ? el("c:bubble3D", { val: false }) : undefined,
          el(
            "c:spPr",
            null,
            lineChart
              ? el(
                  "a:ln",
                  null,
                  el("a:solidFill", null, color(point.color, theme)),
                )
              : el("a:solidFill", null, color(point.color, theme)),
          ),
        ),
      ];
    },
  );
  return el(
    "c:ser",
    null,
    el("c:idx", { val: index }),
    el("c:order", { val: index }),
    el("c:tx", null, strRef(`${sheet}!$${column}$1`, [item.name])),
    itemShape,
    bar ? el("c:invertIfNegative", { val: false }) : undefined,
    lineChart
      ? el("c:marker", null, el("c:symbol", { val: "none" }))
      : undefined,
    pie && explosion !== undefined
      ? el("c:explosion", { val: Math.round(explosion) })
      : undefined,
    points,
    dataLabels(kind, labels, item, index, textStyles),
    el("c:cat", null, strRef(`${sheet}!$A$2:$A$${rows}`, data.categories)),
    el(
      "c:val",
      null,
      numRef(`${sheet}!$${column}$2:$${column}$${rows}`, item.values),
    ),
    lineChart ? el("c:smooth", { val: false }) : undefined,
  );
}

function grouping(
  stacking: "stacked" | "percent" | undefined,
  fallback: string,
): string {
  if (stacking === "stacked") return "stacked";
  if (stacking === "percent") return "percentStacked";
  return fallback;
}

function chartPlot(
  props: ChartProps,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement {
  const explosion =
    props.kind === "pie" || props.kind === "doughnut"
      ? props.explosion
      : undefined;
  const items = props.data.series.map((item, index) =>
    chartSeries(
      props.kind,
      props.data,
      item,
      index,
      props.dataLabels,
      textStyles,
      explosion,
      theme,
    ),
  );
  const axes = [
    el("c:axId", { val: categoryAxisId }),
    el("c:axId", { val: valueAxisId }),
  ];
  switch (props.kind) {
    case "column":
    case "bar":
      return el(
        "c:barChart",
        null,
        el("c:barDir", { val: props.kind === "bar" ? "bar" : "col" }),
        el("c:grouping", { val: grouping(props.stacking, "clustered") }),
        el("c:varyColors", { val: false }),
        items,
        el("c:gapWidth", { val: Math.round(props.gapWidth ?? 150) }),
        props.stacking !== undefined
          ? el("c:overlap", { val: 100 })
          : props.overlap === undefined
            ? undefined
            : el("c:overlap", { val: Math.round(props.overlap) }),
        axes,
      );
    case "line":
      return el(
        "c:lineChart",
        null,
        el("c:grouping", { val: grouping(props.stacking, "standard") }),
        el("c:varyColors", { val: false }),
        items,
        el("c:marker", { val: true }),
        axes,
      );
    case "area":
      return el(
        "c:areaChart",
        null,
        el("c:grouping", { val: grouping(props.stacking, "standard") }),
        el("c:varyColors", { val: false }),
        items,
        axes,
      );
    case "pie":
    case "doughnut":
      return el(
        props.kind === "pie" ? "c:pieChart" : "c:doughnutChart",
        null,
        el("c:varyColors", { val: true }),
        items,
        el("c:firstSliceAng", { val: props.firstSliceAngle ?? 0 }),
        props.kind === "doughnut"
          ? el("c:holeSize", { val: props.holeSize ?? 50 })
          : undefined,
      );
  }
}

function lineXml(
  value: boolean | ChartLine | undefined,
  theme: ThemeContext,
): XmlElement | undefined {
  if (value === undefined || value === true) return undefined;
  if (value === false)
    return el("c:spPr", null, el("a:ln", null, el("a:noFill")));
  return shapeProperties(value, theme);
}

function gridlines(
  value: ChartAxis | undefined,
  theme: ThemeContext,
): XmlElement | undefined {
  if (value?.gridlines === undefined || value.gridlines === false)
    return undefined;
  return el(
    "c:majorGridlines",
    null,
    value.gridlines === true
      ? undefined
      : shapeProperties(value.gridlines, theme),
  );
}

function axes(
  props: ChartProps,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement[] {
  if (props.kind === "pie" || props.kind === "doughnut") return [];
  const category = props.categoryAxis;
  const value = props.valueAxis;
  const horizontal = props.kind === "bar";
  return [
    el(
      "c:catAx",
      null,
      el("c:axId", { val: categoryAxisId }),
      el("c:scaling", null, el("c:orientation", { val: "minMax" })),
      el("c:delete", { val: category?.visible === false }),
      el("c:axPos", { val: horizontal ? "l" : "b" }),
      gridlines(category, theme),
      el("c:numFmt", { formatCode: "General", sourceLinked: true }),
      el("c:majorTickMark", { val: "none" }),
      el("c:minorTickMark", { val: "none" }),
      el("c:tickLblPos", { val: "nextTo" }),
      lineXml(category?.line, theme),
      raw(textStyles.categoryAxis),
      el("c:crossAx", { val: valueAxisId }),
      el("c:crosses", { val: "autoZero" }),
      el("c:auto", { val: true }),
      el("c:lblAlgn", { val: "ctr" }),
      el("c:lblOffset", { val: 100 }),
      el("c:noMultiLvlLbl", { val: false }),
    ),
    el(
      "c:valAx",
      null,
      el("c:axId", { val: valueAxisId }),
      el(
        "c:scaling",
        null,
        el("c:orientation", { val: "minMax" }),
        value?.max === undefined ? undefined : el("c:max", { val: value.max }),
        value?.min === undefined ? undefined : el("c:min", { val: value.min }),
      ),
      el("c:delete", { val: value?.visible === false }),
      el("c:axPos", { val: horizontal ? "b" : "l" }),
      gridlines(value, theme),
      el("c:numFmt", {
        formatCode: value?.numberFormat ?? "General",
        sourceLinked: value?.numberFormat === undefined,
      }),
      el("c:majorTickMark", { val: "none" }),
      el("c:minorTickMark", { val: "none" }),
      el("c:tickLblPos", { val: "nextTo" }),
      lineXml(value?.line, theme),
      raw(textStyles.valueAxis),
      el("c:crossAx", { val: categoryAxisId }),
      el("c:crosses", { val: "autoZero" }),
      el("c:crossBetween", {
        val: props.kind === "area" ? "midCat" : "between",
      }),
      value?.majorUnit === undefined
        ? undefined
        : el("c:majorUnit", { val: value.majorUnit }),
    ),
  ];
}

const legendPositions = {
  top: "t",
  bottom: "b",
  left: "l",
  right: "r",
} as const;

function plotLayout(props: ChartProps): XmlElement {
  const area = props.plotArea;
  if (area === undefined) return el("c:layout");
  const fraction = (value: number, total: number) =>
    Math.round((value / total) * 1e6) / 1e6;
  return el(
    "c:layout",
    null,
    el(
      "c:manualLayout",
      null,
      el("c:layoutTarget", { val: "inner" }),
      el("c:xMode", { val: "edge" }),
      el("c:yMode", { val: "edge" }),
      el("c:x", { val: fraction(area.left, props.width ?? 1) }),
      el("c:y", { val: fraction(area.top, props.height ?? 1) }),
      el("c:w", { val: fraction(area.width, props.width ?? 1) }),
      el("c:h", { val: fraction(area.height, props.height ?? 1) }),
    ),
  );
}

export function chartPartXml(
  props: ChartProps,
  relatedRelationshipId: string,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): string {
  return serialize(
    el(
      "c:chartSpace",
      {
        "xmlns:c": chartNamespace,
        "xmlns:a": drawingNamespace,
        "xmlns:r": relationshipNamespace,
      },
      el("c:date1904", { val: false }),
      el("c:lang", { val: "en-US" }),
      el("c:roundedCorners", { val: false }),
      el(
        "c:chart",
        null,
        el("c:autoTitleDeleted", { val: true }),
        el(
          "c:plotArea",
          null,
          plotLayout(props),
          chartPlot(props, textStyles, theme),
          axes(props, textStyles, theme),
          transparent(),
        ),
        props.legend === undefined || props.legend === false
          ? undefined
          : el(
              "c:legend",
              null,
              el("c:legendPos", { val: legendPositions[props.legend] }),
              el("c:overlay", { val: false }),
            ),
        el("c:plotVisOnly", { val: true }),
        el("c:dispBlanksAs", { val: "gap" }),
      ),
      transparent(),
      raw(textStyles.chart),
      el(
        "c:externalData",
        { "r:id": relatedRelationshipId },
        el("c:autoUpdate", { val: false }),
      ),
    ),
  );
}

export function chartFrameXml(
  id: number,
  name: string,
  box: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
  relationshipId: string,
  link: NativeLink | undefined,
): string {
  const linkAttributes = {
    ...(link?.relationshipId === undefined
      ? {}
      : { "r:id": link.relationshipId }),
    ...(link?.action === undefined ? {} : { action: link.action }),
  };
  return serializeFragment(
    el(
      "p:graphicFrame",
      null,
      el(
        "p:nvGraphicFramePr",
        null,
        el(
          "p:cNvPr",
          { id, name },
          link === undefined ? undefined : el("a:hlinkClick", linkAttributes),
        ),
        el("p:cNvGraphicFramePr"),
        el("p:nvPr"),
      ),
      el(
        "p:xfrm",
        null,
        el("a:off", { x: pxToEmu(box.x), y: pxToEmu(box.y) }),
        el("a:ext", { cx: pxToEmu(box.width), cy: pxToEmu(box.height) }),
      ),
      el(
        "a:graphic",
        null,
        el(
          "a:graphicData",
          { uri: chartNamespace },
          el("c:chart", {
            "xmlns:c": chartNamespace,
            "xmlns:r": relationshipNamespace,
            "r:id": relationshipId,
          }),
        ),
      ),
    ),
  );
}

function inlineString(reference: string, value: string): XmlElement {
  return el(
    "c",
    { r: reference, t: "inlineStr" },
    el("is", null, el("t", null, value)),
  );
}

export function chartWorkbook(data: ChartData): Uint8Array {
  const spreadsheet =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const officeRelationships = relationshipNamespace;
  const header = el(
    "row",
    { r: 1 },
    inlineString("A1", " "),
    data.series.map((item, index) =>
      inlineString(`${columnName(index + 1)}1`, item.name),
    ),
  );
  const rows = data.categories.map((category, row) =>
    el(
      "row",
      { r: row + 2 },
      inlineString(`A${row + 2}`, category),
      data.series.flatMap((item, index) => {
        const value = item.values[row];
        return value === null || value === undefined
          ? []
          : [
              el(
                "c",
                { r: `${columnName(index + 1)}${row + 2}` },
                el("v", null, String(value)),
              ),
            ];
      }),
    ),
  );
  const parts: Record<string, XmlElement> = {
    "[Content_Types].xml": el(
      "Types",
      { xmlns: "http://schemas.openxmlformats.org/package/2006/content-types" },
      el("Default", {
        Extension: "rels",
        ContentType: "application/vnd.openxmlformats-package.relationships+xml",
      }),
      el("Default", { Extension: "xml", ContentType: "application/xml" }),
      el("Override", {
        PartName: "/xl/workbook.xml",
        ContentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
      }),
      el("Override", {
        PartName: "/xl/worksheets/sheet1.xml",
        ContentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
      }),
      el("Override", {
        PartName: "/xl/styles.xml",
        ContentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
      }),
    ),
    "_rels/.rels": el(
      "Relationships",
      { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
      el("Relationship", {
        Id: "rId1",
        Type: `${officeRelationships}/officeDocument`,
        Target: "xl/workbook.xml",
      }),
    ),
    "xl/workbook.xml": el(
      "workbook",
      { xmlns: spreadsheet, "xmlns:r": officeRelationships },
      el(
        "sheets",
        null,
        el("sheet", { name: sheet, sheetId: 1, "r:id": "rId1" }),
      ),
    ),
    "xl/_rels/workbook.xml.rels": el(
      "Relationships",
      { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
      el("Relationship", {
        Id: "rId1",
        Type: `${officeRelationships}/worksheet`,
        Target: "worksheets/sheet1.xml",
      }),
      el("Relationship", {
        Id: "rId2",
        Type: `${officeRelationships}/styles`,
        Target: "styles.xml",
      }),
    ),
    "xl/worksheets/sheet1.xml": el(
      "worksheet",
      { xmlns: spreadsheet },
      el("sheetData", null, header, rows),
    ),
    "xl/styles.xml": el(
      "styleSheet",
      { xmlns: spreadsheet },
      el(
        "fonts",
        { count: 1 },
        el("font", null, el("sz", { val: 11 }), el("name", { val: "Calibri" })),
      ),
      el(
        "fills",
        { count: 2 },
        el("fill", null, el("patternFill", { patternType: "none" })),
        el("fill", null, el("patternFill", { patternType: "gray125" })),
      ),
      el(
        "borders",
        { count: 1 },
        el(
          "border",
          null,
          el("left"),
          el("right"),
          el("top"),
          el("bottom"),
          el("diagonal"),
        ),
      ),
      el(
        "cellStyleXfs",
        { count: 1 },
        el("xf", { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }),
      ),
      el(
        "cellXfs",
        { count: 1 },
        el("xf", { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }),
      ),
    ),
  };
  const mtime = new Date("2020-01-01T00:00:00Z");
  return zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([path, xml]) => [
        path,
        [strToU8(serialize(xml)), { mtime }],
      ]),
    ),
  );
}
