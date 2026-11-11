import {
  pxToEmu,
  resolvedColorXml,
  xmlEscape,
  type ColorValue,
  type NativeLink,
  type ThemeContext,
} from "@react-pptx/core/capability";
import { strToU8, zipSync } from "fflate";

import { chartDataLabelPositions, mergeDataLabels } from "./chart-labels.ts";
import type {
  ChartAxis,
  ChartDataLabels,
  ChartKind,
  ChartLine,
  ChartMarker,
  ChartPoint,
} from "./chart-types.ts";
import type {
  ParsedChart,
  ParsedChartAxis,
  ParsedChartGroup,
  ParsedChartSeries,
} from "./chart.ts";

const chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const drawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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

const presetDashValues = {
  solid: "solid",
  dash: "dash",
  dot: "dot",
  dashDot: "dashDot",
  longDash: "lgDash",
  longDashDot: "lgDashDot",
  longDashDotDot: "lgDashDotDot",
} as const;

function chartLine(
  value: ChartLine | undefined,
  theme: ThemeContext,
): XmlElement | undefined {
  if (value === undefined) return undefined;
  return el(
    "a:ln",
    { w: pxToEmu(value.width ?? 1) },
    el("a:solidFill", null, color(value.color ?? "dk1", theme)),
    value.dash === undefined
      ? undefined
      : el("a:prstDash", { val: presetDashValues[value.dash] }),
  );
}

function markerShapeProperties(
  value: ChartMarker,
  theme: ThemeContext,
): XmlElement | undefined {
  if (value.fill === undefined && value.line === undefined) return undefined;
  return el(
    "c:spPr",
    null,
    value.fill === undefined
      ? undefined
      : el("a:solidFill", null, color(value.fill, theme)),
    value.line === undefined ? undefined : chartLine(value.line, theme),
  );
}

function chartMarker(
  value: boolean | ChartMarker | undefined,
  theme: ThemeContext,
): XmlElement {
  if (value === undefined || value === false) {
    return el("c:marker", null, el("c:symbol", { val: "none" }));
  }
  if (value === true) {
    return el("c:marker", null, el("c:symbol", { val: "circle" }));
  }
  return el(
    "c:marker",
    null,
    el("c:symbol", { val: value.symbol ?? "circle" }),
    value.size === undefined ? undefined : el("c:size", { val: value.size }),
    markerShapeProperties(value, theme),
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
  readonly title?: string;
  readonly axisTitles?: ReadonlyMap<string, string>;
  readonly categoryAxis?: string;
  readonly valueAxis?: string;
  readonly groupDataLabels?: ReadonlyMap<number, string>;
  readonly pointLabels?: ReadonlyMap<string, string>;
};

function dataLabels(
  group: ParsedChartGroup,
  series: ParsedChartSeries,
  groupIndex: number,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement | undefined {
  const seriesLabels = mergeDataLabels(group.dataLabels, series.dataLabels);
  const hasPosition =
    group.kind === "column" ||
    group.kind === "bar" ||
    group.kind === "line" ||
    group.kind === "pie";
  const hasVisibleContent = (labels: ChartDataLabels) =>
    labels.value === true ||
    labels.percent === true ||
    labels.category === true;
  const leaderLineSettings = (labels: ChartDataLabels) => {
    if (labels.leaderLines === undefined) return [];
    const visible = hasVisibleContent(labels);
    if (labels.leaderLines === false) {
      // Office requires a visible sibling whenever showLeaderLines is present, including
      // the false value. Omitting the element keeps the labels disabled without the
      // invalid empty display group.
      return visible ? [el("c:showLeaderLines", { val: false })] : [];
    }
    if (!visible) return [];
    return [
      el("c:showLeaderLines", { val: true }),
      el(
        "c:leaderLines",
        null,
        typeof labels.leaderLines === "object"
          ? el("c:spPr", null, chartLine(labels.leaderLines, theme))
          : undefined,
      ),
    ];
  };
  const settings = (
    labels: ChartDataLabels,
    textStyle: string | undefined,
    includeLeaderLines: boolean,
  ) => [
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
      : el("c:dLblPos", {
          val: chartDataLabelPositions[labels.position],
        }),
    el("c:showLegendKey", { val: false }),
    el("c:showVal", { val: labels.value ?? false }),
    el("c:showCatName", { val: labels.category ?? false }),
    el("c:showSerName", { val: false }),
    group.kind === "pie" || group.kind === "doughnut"
      ? el("c:showPercent", { val: labels.percent ?? false })
      : undefined,
    el("c:showBubbleSize", { val: false }),
    includeLeaderLines ? leaderLineSettings(labels) : undefined,
  ];
  const groupStyle = textStyles.groupDataLabels?.get(groupIndex);
  const pointBase = seriesLabels === false ? {} : seriesLabels;
  const pointLabels = Object.entries<ChartPoint>(series.points ?? {}).flatMap(
    ([pointIndex, point]) => {
      if (point.dataLabels === undefined && point.labelColor === undefined) {
        return [];
      }
      const index = Number(pointIndex);
      const effective = mergeDataLabels(pointBase, point.dataLabels);
      if (effective === false) {
        return [
          el(
            "c:dLbl",
            null,
            el("c:idx", { val: index }),
            el("c:delete", { val: true }),
          ),
        ];
      }
      if (effective === undefined) return [];
      if (point.dataLabels === undefined && !hasVisibleContent(effective)) {
        return [];
      }
      const style =
        textStyles.pointLabels?.get(`${series.seriesIndex}:${pointIndex}`) ??
        groupStyle;
      return [
        el(
          "c:dLbl",
          null,
          el("c:idx", { val: index }),
          settings(effective, style, false),
        ),
      ];
    },
  );
  const base =
    seriesLabels === false
      ? pointLabels.length > 0
        ? {}
        : false
      : (seriesLabels ?? (pointLabels.length > 0 ? {} : undefined));
  if (base === undefined) return undefined;
  if (base === false) {
    return el("c:dLbls", null, el("c:delete", { val: true }));
  }
  return el("c:dLbls", null, pointLabels, settings(base, groupStyle, true));
}

function seriesShape(
  kind: ChartKind,
  itemColor: ColorValue | undefined,
  line: ChartLine | undefined,
  theme: ThemeContext,
): XmlElement | undefined {
  if (kind === "line") {
    const stroke =
      line === undefined
        ? itemColor === undefined
          ? undefined
          : { color: itemColor }
        : {
            ...line,
            ...(line.color === undefined && itemColor !== undefined
              ? { color: itemColor }
              : {}),
          };
    return stroke === undefined
      ? undefined
      : el("c:spPr", null, chartLine(stroke, theme));
  }
  if (itemColor === undefined && line === undefined) return undefined;
  return el(
    "c:spPr",
    null,
    itemColor === undefined
      ? undefined
      : el("a:solidFill", null, color(itemColor, theme)),
    line === undefined ? undefined : chartLine(line, theme),
  );
}

function pointShape(
  kind: ChartKind,
  point: ChartPoint,
  theme: ThemeContext,
): XmlElement | undefined {
  const line =
    kind === "line"
      ? point.line === undefined
        ? point.color === undefined
          ? undefined
          : { color: point.color }
        : {
            ...point.line,
            ...(point.line.color === undefined && point.color !== undefined
              ? { color: point.color }
              : {}),
          }
      : point.line;
  if (kind === "line") {
    return line === undefined
      ? undefined
      : el("c:spPr", null, chartLine(line, theme));
  }
  if (point.color === undefined && line === undefined) return undefined;
  return el(
    "c:spPr",
    null,
    point.color === undefined
      ? undefined
      : el("a:solidFill", null, color(point.color, theme)),
    line === undefined ? undefined : chartLine(line, theme),
  );
}

type SeriesCells = {
  /** One-based worksheet category column for this series domain. */
  readonly categoryColumn: number;
  /** One-based worksheet value column; also holds the series name in row 1. */
  readonly valueColumn: number;
  readonly rows: number;
};

/** Derives every workbook/cache formula range from the same resolved series domain. */
function seriesCells(
  group: ParsedChartGroup,
  series: ParsedChartSeries,
): SeriesCells {
  if (group.data.kind === "independent-categories") {
    return {
      categoryColumn: series.seriesIndex * 2 + 1,
      valueColumn: series.seriesIndex * 2 + 2,
      rows: series.categories.length,
    };
  }
  return {
    categoryColumn: 1,
    valueColumn: series.seriesIndex + 2,
    rows: series.categories.length,
  };
}

function chartSeries(
  group: ParsedChartGroup,
  series: ParsedChartSeries,
  groupIndex: number,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement {
  const cells = seriesCells(group, series);
  const categoryColumn = columnName(cells.categoryColumn - 1);
  const valueColumn = columnName(cells.valueColumn - 1);
  const rows = cells.rows + 1;
  const kind: ChartKind = group.kind;
  const bar = kind === "column" || kind === "bar";
  const lineChart = kind === "line";
  const pie = kind === "pie" || kind === "doughnut";
  const itemShape = seriesShape(kind, series.color, series.line, theme);
  const seriesExplosion = series.explosion ?? group.explosion;
  const points = Object.entries<ChartPoint>(series.points ?? {}).flatMap(
    ([pointIndex, point]) => {
      const shape = pointShape(kind, point, theme);
      const marker =
        lineChart && point.marker !== undefined
          ? chartMarker(point.marker, theme)
          : undefined;
      if (
        shape === undefined &&
        marker === undefined &&
        point.explosion === undefined
      ) {
        return [];
      }
      return [
        el(
          "c:dPt",
          null,
          el("c:idx", { val: Number(pointIndex) }),
          bar ? el("c:invertIfNegative", { val: false }) : undefined,
          bar ? el("c:bubble3D", { val: false }) : undefined,
          marker,
          point.explosion === undefined
            ? undefined
            : el("c:explosion", { val: Math.round(point.explosion) }),
          shape,
        ),
      ];
    },
  );
  return el(
    "c:ser",
    null,
    el("c:idx", { val: series.seriesIndex }),
    el("c:order", { val: series.seriesIndex }),
    el("c:tx", null, strRef(`${sheet}!$${valueColumn}$1`, [series.name])),
    itemShape,
    bar ? el("c:invertIfNegative", { val: false }) : undefined,
    lineChart ? chartMarker(series.marker, theme) : undefined,
    pie && seriesExplosion !== undefined
      ? el("c:explosion", { val: Math.round(seriesExplosion) })
      : undefined,
    points,
    dataLabels(group, series, groupIndex, textStyles, theme),
    el(
      "c:cat",
      null,
      strRef(
        `${sheet}!$${categoryColumn}$2:$${categoryColumn}$${rows}`,
        series.categories,
      ),
    ),
    el(
      "c:val",
      null,
      numRef(
        `${sheet}!$${valueColumn}$2:$${valueColumn}$${rows}`,
        series.values,
      ),
    ),
    lineChart ? el("c:smooth", { val: series.smooth ?? false }) : undefined,
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

function chartGroup(
  group: ParsedChartGroup,
  groupIndex: number,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement {
  const items = group.series.map((item) =>
    chartSeries(group, item, groupIndex, textStyles, theme),
  );
  const axes =
    group.axes === undefined
      ? []
      : [
          el("c:axId", { val: group.axes.category.id }),
          el("c:axId", { val: group.axes.value.id }),
        ];
  switch (group.kind) {
    case "column":
    case "bar":
      return el(
        "c:barChart",
        null,
        el("c:barDir", { val: group.kind === "bar" ? "bar" : "col" }),
        el("c:grouping", { val: grouping(group.stacking, "clustered") }),
        el("c:varyColors", { val: false }),
        items,
        el("c:gapWidth", { val: Math.round(group.gapWidth ?? 150) }),
        group.overlap !== undefined
          ? el("c:overlap", { val: Math.round(group.overlap) })
          : group.stacking === undefined
            ? undefined
            : el("c:overlap", { val: 100 }),
        axes,
      );
    case "line":
      return el(
        "c:lineChart",
        null,
        el("c:grouping", { val: grouping(group.stacking, "standard") }),
        el("c:varyColors", { val: false }),
        items,
        el("c:marker", { val: true }),
        axes,
      );
    case "area":
      return el(
        "c:areaChart",
        null,
        el("c:grouping", { val: grouping(group.stacking, "standard") }),
        el("c:varyColors", { val: false }),
        items,
        axes,
      );
    case "pie":
    case "doughnut":
      return el(
        group.kind === "pie" ? "c:pieChart" : "c:doughnutChart",
        null,
        el("c:varyColors", { val: true }),
        items,
        el("c:firstSliceAng", { val: group.firstSliceAngle ?? 0 }),
        group.kind === "doughnut"
          ? el("c:holeSize", { val: group.holeSize ?? 50 })
          : undefined,
      );
    default: {
      const exhaustiveCheck: never = group.kind;
      throw new Error(
        `Unsupported chart kind: ${JSON.stringify(exhaustiveCheck)}`,
      );
    }
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
  kind: "major" | "minor",
  theme: ThemeContext,
): XmlElement | undefined {
  const setting = kind === "major" ? value?.gridlines : value?.minorGridlines;
  if (setting === undefined || setting === false) return undefined;
  return el(
    kind === "major" ? "c:majorGridlines" : "c:minorGridlines",
    null,
    setting === true ? undefined : shapeProperties(setting, theme),
  );
}

function axisTitle(
  spec: ChartAxis,
  textStyles: ChartTextStyleXml,
  name: string,
): XmlRaw | undefined {
  return raw(textStyles.axisTitles?.get(name));
}

function parsedAxis(
  axes: readonly ParsedChartAxis[],
  role: "category" | "value",
): ParsedChartAxis | undefined {
  return axes.find((axis) => axis.role === role);
}

function axisScaling(spec: ChartAxis, valueAxis: boolean): XmlElement {
  const valueSpec = valueAxis
    ? (spec as Extract<ChartAxis, { kind: "value" }>)
    : undefined;
  return el(
    "c:scaling",
    null,
    valueSpec?.logBase === undefined
      ? undefined
      : el("c:logBase", { val: valueSpec.logBase }),
    el("c:orientation", { val: spec.reversed === true ? "maxMin" : "minMax" }),
    valueSpec?.max === undefined
      ? undefined
      : el("c:max", { val: valueSpec.max }),
    valueSpec?.min === undefined
      ? undefined
      : el("c:min", { val: valueSpec.min }),
  );
}

function axisCrossing(spec: ChartAxis): XmlElement {
  return spec.crossesAt === undefined
    ? el("c:crosses", { val: spec.crosses ?? "autoZero" })
    : el("c:crossesAt", { val: spec.crossesAt });
}

function axes(
  parsed: ParsedChart,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement[] {
  if (parsed.groups[0]?.axes === undefined) return [];
  const category = parsedAxis(parsed.axes, "category");
  const value = parsedAxis(parsed.axes, "value");
  if (category === undefined || value === undefined) {
    throw new Error(
      "Cartesian chart is missing resolved category or value axis",
    );
  }
  const categorySpec = category.spec as Extract<
    ChartAxis,
    { kind: "category" }
  >;
  const valueSpec = value.spec as Extract<ChartAxis, { kind: "value" }>;
  const horizontal = parsed.groups[0].kind === "bar";
  const firstKind = parsed.groups[0].kind;
  return [
    el(
      "c:catAx",
      null,
      el("c:axId", { val: category.id }),
      axisScaling(categorySpec, false),
      el("c:delete", { val: categorySpec.visible === false }),
      el("c:axPos", { val: horizontal ? "l" : "b" }),
      gridlines(categorySpec, "major", theme),
      gridlines(categorySpec, "minor", theme),
      axisTitle(categorySpec, textStyles, category.name),
      el("c:numFmt", { formatCode: "General", sourceLinked: true }),
      el("c:majorTickMark", { val: categorySpec.majorTickMark ?? "none" }),
      el("c:minorTickMark", { val: categorySpec.minorTickMark ?? "none" }),
      el("c:tickLblPos", {
        val: categorySpec.tickLabelPosition ?? "nextTo",
      }),
      lineXml(categorySpec.line, theme),
      raw(textStyles.categoryAxis),
      el("c:crossAx", { val: value.id }),
      axisCrossing(categorySpec),
      el("c:auto", { val: true }),
      el("c:lblAlgn", { val: "ctr" }),
      el("c:lblOffset", { val: 100 }),
      el("c:noMultiLvlLbl", { val: false }),
    ),
    el(
      "c:valAx",
      null,
      el("c:axId", { val: value.id }),
      axisScaling(valueSpec, true),
      el("c:delete", { val: valueSpec.visible === false }),
      el("c:axPos", { val: horizontal ? "b" : "l" }),
      gridlines(valueSpec, "major", theme),
      gridlines(valueSpec, "minor", theme),
      axisTitle(valueSpec, textStyles, value.name),
      el("c:numFmt", {
        formatCode: valueSpec.numberFormat ?? "General",
        sourceLinked: valueSpec.numberFormat === undefined,
      }),
      el("c:majorTickMark", { val: valueSpec.majorTickMark ?? "none" }),
      el("c:minorTickMark", { val: valueSpec.minorTickMark ?? "none" }),
      el("c:tickLblPos", {
        val: valueSpec.tickLabelPosition ?? "nextTo",
      }),
      lineXml(valueSpec.line, theme),
      raw(textStyles.valueAxis),
      el("c:crossAx", { val: category.id }),
      axisCrossing(valueSpec),
      el("c:crossBetween", {
        val: firstKind === "area" ? "midCat" : "between",
      }),
      valueSpec.majorUnit === undefined
        ? undefined
        : el("c:majorUnit", { val: valueSpec.majorUnit }),
      valueSpec.minorUnit === undefined
        ? undefined
        : el("c:minorUnit", { val: valueSpec.minorUnit }),
    ),
  ];
}

const legendPositions = {
  top: "t",
  bottom: "b",
  left: "l",
  right: "r",
} as const;

function plotLayout(
  parsed: ParsedChart,
  box: { readonly width: number; readonly height: number },
): XmlElement {
  const area = parsed.plotArea;
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
      el("c:x", { val: fraction(area.left, box.width) }),
      el("c:y", { val: fraction(area.top, box.height) }),
      el("c:w", { val: fraction(area.width, box.width) }),
      el("c:h", { val: fraction(area.height, box.height) }),
    ),
  );
}

export function chartPartXml(
  parsed: ParsedChart,
  box: { readonly width: number; readonly height: number },
  relatedRelationshipId: string,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): string {
  const chartLang =
    (typeof parsed.style?.lang === "string" ? parsed.style.lang : undefined) ??
    theme.textStyle.lang ??
    "en-US";
  return serialize(
    el(
      "c:chartSpace",
      {
        "xmlns:c": chartNamespace,
        "xmlns:a": drawingNamespace,
        "xmlns:r": relationshipNamespace,
      },
      el("c:date1904", { val: false }),
      el("c:lang", { val: chartLang }),
      el("c:roundedCorners", { val: false }),
      el(
        "c:chart",
        null,
        raw(textStyles.title),
        el("c:autoTitleDeleted", { val: parsed.title === undefined }),
        el(
          "c:plotArea",
          null,
          plotLayout(parsed, box),
          parsed.groups.map((group, groupIndex) =>
            chartGroup(group, groupIndex, textStyles, theme),
          ),
          axes(parsed, textStyles, theme),
          transparent(),
        ),
        parsed.legend === undefined || parsed.legend === false
          ? undefined
          : el(
              "c:legend",
              null,
              el("c:legendPos", { val: legendPositions[parsed.legend] }),
              el("c:overlay", { val: false }),
            ),
        el("c:plotVisOnly", { val: true }),
        parsed.displayBlanksAs === undefined
          ? undefined
          : el("c:dispBlanksAs", { val: parsed.displayBlanksAs }),
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

export function chartWorkbook(parsed: ParsedChart): Uint8Array {
  const spreadsheet =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const officeRelationships = relationshipNamespace;
  const independent = parsed.groups[0]?.data.kind === "independent-categories";
  const header = el(
    "row",
    { r: 1 },
    inlineString("A1", " "),
    independent
      ? parsed.series.map((item) => {
          const { valueColumn } = seriesCells(
            parsed.groups[item.groupIndex]!,
            item,
          );
          return inlineString(`${columnName(valueColumn - 1)}1`, item.name);
        })
      : parsed.series.map((item) =>
          inlineString(`${columnName(item.seriesIndex + 1)}1`, item.name),
        ),
  );
  const rows = independent
    ? Array.from(
        {
          length: Math.max(
            0,
            ...parsed.series.map((item) => item.categories.length),
          ),
        },
        (_, row) =>
          el(
            "row",
            { r: row + 2 },
            parsed.series.flatMap((item) => {
              const category = item.categories[row];
              if (category === undefined) return [];
              const { categoryColumn, valueColumn } = seriesCells(
                parsed.groups[item.groupIndex]!,
                item,
              );
              const value = item.values[row];
              return [
                inlineString(
                  `${columnName(categoryColumn - 1)}${row + 2}`,
                  category,
                ),
                ...(value === null || value === undefined
                  ? []
                  : [
                      el(
                        "c",
                        { r: `${columnName(valueColumn - 1)}${row + 2}` },
                        el("v", null, String(value)),
                      ),
                    ]),
              ];
            }),
          ),
      )
    : parsed.categories.map((category, row) =>
        el(
          "row",
          { r: row + 2 },
          inlineString(`A${row + 2}`, category),
          parsed.series.flatMap((item) => {
            const value = item.values[row];
            return value === null || value === undefined
              ? []
              : [
                  el(
                    "c",
                    { r: `${columnName(item.seriesIndex + 1)}${row + 2}` },
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
