import type {
  ColorValue,
  NativeLink,
  ThemeContext,
} from "@react-pptx/core/capability";
import { pxToEmu, resolvedColorXml, xmlEscape } from "@react-pptx/emit";
import { strToU8, zipSync } from "fflate";

import { chartDataLabelPositions, mergeDataLabels } from "./chart-labels.ts";
import type {
  ParsedAxisCrossing,
  ParsedChart,
  ParsedChartAxis,
  ParsedChartGroup,
  ParsedChartSeries,
  ParsedScatterSeries,
} from "./chart-model.ts";
import type {
  ChartAxis,
  ChartDataLabelSettings,
  ChartErrorBars,
  ChartKind,
  ChartLine,
  ChartMarker,
  ChartPoint,
  ChartTrendline,
} from "./chart-types.ts";

const chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const drawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const sheet = "Sheet1";
// Microsoft chart extensions used only by independent doughnut category labels.
const c15Namespace = "http://schemas.microsoft.com/office/drawing/2012/chart";
const categoryRangeExtensionUri = "{02D57815-91ED-43cb-92C2-25804820EDAC}";
const dataLabelRangeExtensionUri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}";

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

type ChartSeriesCells = {
  readonly nameColumn: number;
  readonly category?: {
    readonly column: number;
    readonly pointCount: number;
  };
  readonly value?: {
    readonly column: number;
  };
  readonly x?: {
    readonly column: number;
    readonly pointCount: number;
    readonly labelFormats: ScatterXLabelFormats;
  };
  readonly y?: {
    readonly column: number;
    readonly pointCount: number;
  };
};

type ScatterXLabelFormats = {
  readonly baseFormatCode: string;
  readonly pointFormatCodes: readonly string[];
};

type ChartCells = {
  readonly categoryColumns: ReadonlyMap<
    string,
    { readonly column: number; readonly categories: readonly string[] }
  >;
  readonly axislessCategoryColumns: ReadonlyMap<
    number,
    { readonly column: number; readonly categories: readonly string[] }
  >;
  readonly independentCategoryColumns: ReadonlyMap<
    number,
    { readonly column: number; readonly categories: readonly string[] }
  >;
  readonly series: ReadonlyMap<number, ChartSeriesCells>;
  readonly projectedScatterXAxes: ReadonlySet<number>;
  readonly nextColumn: number;
};

function scatterXFormatCode(
  labels: false | ChartDataLabelSettings | undefined,
): string {
  return labels !== undefined && labels !== false && labels.xValue === true
    ? (labels.numberFormat ?? "General")
    : "General";
}

/**
 * Assigns workbook columns and scatter X label formats for every parsed series once.
 *
 * Category axes own the leading columns because several groups may share one ordered domain;
 * each remaining category series owns one value column, and every scatter series owns its own
 * X and Y columns. The chart cache, embedded workbook, and affected default X axis read this
 * mapping so their coordinate formats stay aligned.
 */
function chartCells(parsed: ParsedChart): ChartCells {
  const categoryColumns = new Map<
    string,
    { column: number; categories: readonly string[] }
  >();
  const axislessCategoryColumns = new Map<
    number,
    { column: number; categories: readonly string[] }
  >();
  const independentCategoryColumns = new Map<
    number,
    { column: number; categories: readonly string[] }
  >();
  let nextColumn = 0;
  for (const [groupIndex, group] of parsed.groups.entries()) {
    if (group.kind === "scatter") continue;
    if (group.data.kind === "independent-categories") continue;
    if (group.kind === "pie" || group.kind === "doughnut") {
      // Pie and doughnut groups have no axis identity; only one group may exist.
      axislessCategoryColumns.set(groupIndex, {
        column: nextColumn,
        categories: group.data.categories,
      });
      nextColumn += 1;
      continue;
    }
    const name = group.axes.category.name;
    if (categoryColumns.has(name)) continue;
    categoryColumns.set(name, {
      column: nextColumn,
      categories: group.data.categories,
    });
    nextColumn += 1;
  }
  if (
    categoryColumns.size === 0 &&
    axislessCategoryColumns.size === 0 &&
    parsed.groups.some((group) => group.kind === "scatter")
  ) {
    // Preserve the established scatter workbook layout: column A stays blank and the first
    // series starts in B. Mixed charts begin series columns after their category columns.
    nextColumn = 1;
  }
  const series = new Map<number, ChartSeriesCells>();
  const projectedScatterXAxes = new Set<number>();
  for (const [groupIndex, group] of parsed.groups.entries()) {
    if (group.kind === "scatter") {
      for (const item of group.series) {
        const seriesLabels = mergeDataLabels(group.dataLabels, item.dataLabels);
        const pointBase = seriesLabels === false ? {} : seriesLabels;
        const labelFormats: ScatterXLabelFormats = {
          baseFormatCode: scatterXFormatCode(seriesLabels),
          pointFormatCodes: item.x.map((value, index) => {
            const pointLabels = mergeDataLabels(
              pointBase,
              item.points?.[index]?.dataLabels,
            );
            const formatCode = scatterXFormatCode(pointLabels);
            if (
              value !== null &&
              value !== undefined &&
              formatCode !== "General"
            ) {
              projectedScatterXAxes.add(group.axes.x.id);
            }
            return formatCode;
          }),
        };
        series.set(item.seriesIndex, {
          nameColumn: nextColumn,
          x: {
            column: nextColumn,
            pointCount: item.x.length,
            labelFormats,
          },
          y: { column: nextColumn + 1, pointCount: item.y.length },
        });
        nextColumn += 2;
      }
      continue;
    }
    if (group.data.kind === "independent-categories") {
      for (const item of group.series) {
        const categoryColumn = nextColumn;
        const valueColumn = nextColumn + 1;
        independentCategoryColumns.set(item.seriesIndex, {
          column: categoryColumn,
          categories: item.categories,
        });
        series.set(item.seriesIndex, {
          nameColumn: valueColumn,
          category: {
            column: categoryColumn,
            pointCount: item.categories.length,
          },
          value: { column: valueColumn },
        });
        nextColumn += 2;
      }
      continue;
    }
    const category =
      group.kind === "pie" || group.kind === "doughnut"
        ? axislessCategoryColumns.get(groupIndex)
        : categoryColumns.get(group.axes.category.name);
    if (category === undefined) {
      throw new Error(
        "Category chart group is missing its workbook category column",
      );
    }
    for (const item of group.series) {
      series.set(item.seriesIndex, {
        nameColumn: nextColumn,
        category: {
          column: category.column,
          pointCount: group.data.categories.length,
        },
        value: { column: nextColumn },
      });
      nextColumn += 1;
    }
  }
  return {
    categoryColumns,
    axislessCategoryColumns,
    independentCategoryColumns,
    series,
    projectedScatterXAxes,
    nextColumn,
  };
}

function rangeFormula(column: number, pointCount: number): string {
  const name = columnName(column);
  return `${sheet}!$${name}$2:$${name}$${pointCount + 1}`;
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
  labelFormats?: ScatterXLabelFormats,
): XmlElement {
  const baseFormatCode = labelFormats?.baseFormatCode ?? "General";
  return el(
    "c:numRef",
    null,
    el("c:f", null, formula),
    el(
      "c:numCache",
      null,
      el("c:formatCode", null, baseFormatCode),
      el("c:ptCount", { val: values.length }),
      values.flatMap((value, index) =>
        value === null
          ? []
          : [
              el(
                "c:pt",
                {
                  idx: index,
                  formatCode:
                    labelFormats?.pointFormatCodes[index] === baseFormatCode
                      ? undefined
                      : labelFormats?.pointFormatCodes[index],
                },
                el("c:v", null, String(value)),
              ),
            ],
      ),
    ),
  );
}

/** Literal numeric data source; zero-observation custom error bars use zero points. */
function numLit(values: readonly (number | null)[]): XmlElement {
  return el(
    "c:numLit",
    null,
    el("c:formatCode", null, "General"),
    el("c:ptCount", { val: values.length }),
    values.flatMap((value, index) =>
      value === null
        ? []
        : [el("c:pt", { idx: index }, el("c:v", null, String(value)))],
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

type ErrorBarDirection = "value" | "x" | "y";

type ErrorBarChannel = {
  readonly seriesIndex: number;
  /** Author-facing series axis; category custom channels always use value. */
  readonly channelDirection: ErrorBarDirection;
  readonly side: "plus" | "minus";
  readonly column: number;
  readonly pointCount: number;
  readonly values: readonly (number | null)[];
  readonly header: string;
};

type ErrorBarChannelMap = ReadonlyMap<string, ErrorBarChannel>;

function errorBarChannelKey(
  seriesIndex: number,
  channelDirection: ErrorBarDirection,
  side: "plus" | "minus",
): string {
  return `${seriesIndex}:${channelDirection}:${side}`;
}

/**
 * Allocates one workbook column per declared custom error-bar side after the unique
 * `chartCells()` data columns. Those columns can include multiple category and XY ranges,
 * so this function appends after `cells.nextColumn` instead of assuming a single A column.
 * Both the chart writer and the workbook writer consume this mapping so a formula and its
 * cells cannot drift apart. A category custom series always has the channel identity `value`;
 * the writer decides separately whether the native element stores x or y errDir. Null magnitudes
 * keep their original slot through numRef indices.
 */
function errorBarChannels(
  parsed: ParsedChart,
  startColumn: number,
): ErrorBarChannelMap {
  let nextColumn = startColumn;
  const channels = new Map<string, ErrorBarChannel>();
  const add = (
    seriesIndex: number,
    seriesName: string,
    channelDirection: ErrorBarDirection,
    bars: ChartErrorBars | undefined,
    pointCount: number,
    axisLabel: string,
  ): void => {
    if (bars === undefined || bars.type !== "custom") return;
    for (const side of ["plus", "minus"] as const) {
      const values = bars[side];
      if (values === undefined) continue;
      const header = `${seriesName}${axisLabel === "" ? "" : ` ${axisLabel}`} ${side === "plus" ? "+" : "-"}`;
      channels.set(errorBarChannelKey(seriesIndex, channelDirection, side), {
        seriesIndex,
        channelDirection,
        side,
        column: nextColumn,
        pointCount,
        values,
        header,
      });
      nextColumn += 1;
    }
  };
  for (const series of parsed.series) {
    if (series.errorBars === undefined) continue;
    if ("x" in series) {
      add(
        series.seriesIndex,
        series.name,
        "x",
        series.errorBars.x,
        series.x.length,
        "X",
      );
      add(
        series.seriesIndex,
        series.name,
        "y",
        series.errorBars.y,
        series.y.length,
        "Y",
      );
    } else {
      add(
        series.seriesIndex,
        series.name,
        "value",
        series.errorBars,
        series.values.length,
        "",
      );
    }
  }
  return channels;
}

function errorBarFormula(channel: ErrorBarChannel): string {
  const column = columnName(channel.column - 1);
  return `${sheet}!$${column}$2:$${column}$${channel.pointCount + 1}`;
}

const errorBarValueTypes = {
  custom: "cust",
  fixed: "fixedVal",
  percentage: "percentage",
  standardDeviation: "stdDev",
  standardError: "stdErr",
} as const;

/**
 * Writes one native error-bar element.
 *
 * `channelDirection` selects the author-facing series axis used by the workbook map; category
 * custom bars always use `value` even when their native `errDir` is `y`. `nativeDirection` is
 * the serialized c:errDir: scatters emit their X/Y axis, line emits y, and bar/column/area rely
 * on the documented application default by omitting the element.
 */
function errorBarXml(
  bars: ChartErrorBars,
  channelDirection: ErrorBarDirection,
  nativeDirection: "x" | "y" | undefined,
  seriesIndex: number,
  channels: ErrorBarChannelMap,
  theme: ThemeContext,
): XmlElement {
  const channel = (side: "plus" | "minus"): ErrorBarChannel => {
    const result = channels.get(
      errorBarChannelKey(seriesIndex, channelDirection, side),
    );
    if (result === undefined) {
      throw new Error("Missing workbook channel for a custom error bar");
    }
    return result;
  };
  let plus: XmlElement | undefined;
  let minus: XmlElement | undefined;
  if (bars.type === "custom") {
    if (bars.plus !== undefined) {
      const values = channel("plus");
      plus = el(
        "c:plus",
        null,
        values.values.length === 0
          ? numLit(values.values)
          : numRef(errorBarFormula(values), values.values),
      );
    }
    if (bars.minus !== undefined) {
      const values = channel("minus");
      minus = el(
        "c:minus",
        null,
        values.values.length === 0
          ? numLit(values.values)
          : numRef(errorBarFormula(values), values.values),
      );
    }
  }
  return el(
    "c:errBars",
    null,
    nativeDirection === undefined
      ? undefined
      : el("c:errDir", { val: nativeDirection }),
    el("c:errBarType", { val: bars.include ?? "both" }),
    el("c:errValType", { val: errorBarValueTypes[bars.type] }),
    bars.endCap === false ? el("c:noEndCap", { val: true }) : undefined,
    plus,
    minus,
    bars.type === "custom" || bars.type === "standardError"
      ? undefined
      : el("c:val", { val: bars.value }),
    bars.line === undefined
      ? undefined
      : el("c:spPr", null, chartLine(bars.line, theme)),
  );
}

type CategoryLabelRange = {
  /** Workbook range matching the series-owned ordered categories. */
  readonly formula: string;
  readonly categories: readonly string[];
};

export type ChartTextStyleXml = {
  readonly chart: string;
  readonly title?: string;
  readonly axisTitles?: ReadonlyMap<string, string>;
  readonly axisStyles?: ReadonlyMap<string, string>;
  readonly groupDataLabels?: ReadonlyMap<number, string>;
  readonly pointLabels?: ReadonlyMap<string, string>;
};

function categoryLabelFieldExtension(enabled: boolean): XmlElement {
  return el(
    "c:extLst",
    null,
    el(
      "c:ext",
      { uri: dataLabelRangeExtensionUri, "xmlns:c15": c15Namespace },
      el("c15:showDataLabelsRange", { val: enabled }),
    ),
  );
}

function categoryRangeExtension(range: CategoryLabelRange): XmlElement {
  return el(
    "c:extLst",
    null,
    el(
      "c:ext",
      { uri: categoryRangeExtensionUri, "xmlns:c15": c15Namespace },
      el(
        "c15:datalabelsRange",
        null,
        el("c15:f", null, range.formula),
        el(
          "c15:dlblRangeCache",
          null,
          el("c:ptCount", { val: range.categories.length }),
          range.categories.map((category, index) =>
            el("c:pt", { idx: index }, el("c:v", null, category)),
          ),
        ),
      ),
    ),
  );
}

function categoryLabelsEnabled(
  labels: false | ChartDataLabelSettings | undefined,
): boolean {
  return labels !== undefined && labels !== false && labels.category === true;
}

function dataLabels(
  group: ParsedChartGroup,
  series: ParsedChartSeries | ParsedScatterSeries,
  groupIndex: number,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
  categoryRange?: CategoryLabelRange,
): XmlElement | undefined {
  const seriesLabels = mergeDataLabels(group.dataLabels, series.dataLabels);
  const hasPosition =
    group.kind === "column" ||
    group.kind === "bar" ||
    group.kind === "line" ||
    group.kind === "pie" ||
    group.kind === "scatter";
  const hasVisibleContent = (labels: ChartDataLabelSettings) =>
    labels.value === true ||
    labels.percent === true ||
    labels.category === true ||
    labels.xValue === true;
  const leaderLineSettings = (labels: ChartDataLabelSettings) => {
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
    labels: ChartDataLabelSettings,
    textStyle: string | undefined,
    includeLeaderLines: boolean,
    categoryField: boolean,
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
    // Scatter showCatName displays the horizontal X value; category charts keep category labels.
    el("c:showCatName", {
      val:
        categoryRange === undefined
          ? ((group.kind === "scatter" ? labels.xValue : labels.category) ??
            false)
          : false,
    }),
    el("c:showSerName", { val: false }),
    // Persist percentage labels as disabled where the chart model cannot enable them.
    el("c:showPercent", {
      val:
        group.kind === "pie" || group.kind === "doughnut"
          ? (labels.percent ?? false)
          : false,
    }),
    el("c:showBubbleSize", { val: false }),
    includeLeaderLines ? leaderLineSettings(labels) : undefined,
    categoryField
      ? categoryLabelFieldExtension(categoryLabelsEnabled(labels))
      : undefined,
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
          settings(effective, style, false, categoryRange !== undefined),
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
  return el(
    "c:dLbls",
    null,
    pointLabels,
    settings(base, groupStyle, true, categoryRange !== undefined),
  );
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

function trendlineXml(value: ChartTrendline, theme: ThemeContext): XmlElement {
  const type =
    value.type === "movingAverage"
      ? "movingAvg"
      : value.type === "polynomial"
        ? "poly"
        : value.type === "exponential"
          ? "exp"
          : value.type === "logarithmic"
            ? "log"
            : value.type;
  return el(
    "c:trendline",
    null,
    value.line === undefined
      ? undefined
      : el("c:spPr", null, chartLine(value.line, theme)),
    el("c:trendlineType", { val: type }),
    // CT_Trendline order: order, period, forecast, intercept, R-squared, equation.
    value.type === "polynomial"
      ? el("c:order", { val: value.order })
      : undefined,
    value.type === "movingAverage"
      ? el("c:period", { val: value.period })
      : undefined,
    value.forward === undefined
      ? undefined
      : el("c:forward", { val: value.forward }),
    value.backward === undefined
      ? undefined
      : el("c:backward", { val: value.backward }),
    value.type === "linear" ||
      value.type === "exponential" ||
      value.type === "polynomial"
      ? value.intercept === undefined
        ? undefined
        : el("c:intercept", { val: value.intercept })
      : undefined,
    value.type === "movingAverage"
      ? undefined
      : value.displayRSquared === undefined
        ? undefined
        : el("c:dispRSqr", { val: value.displayRSquared }),
    value.type === "movingAverage"
      ? undefined
      : value.displayEquation === undefined
        ? undefined
        : el("c:dispEq", { val: value.displayEquation }),
  );
}

type SeriesCells = {
  /** Zero-based worksheet category column, matching columnName. */
  readonly categoryColumn: number;
  /** Zero-based worksheet value column; also holds the series name in row 1. */
  readonly valueColumn: number;
  readonly rows: number;
};

/** Maps an independent-ring series to its columns for c15 label-range detection. */
function seriesCells(
  group: ParsedChartGroup,
  series: ParsedChartSeries,
): SeriesCells {
  if (
    group.kind !== "scatter" &&
    group.data.kind === "independent-categories"
  ) {
    return {
      categoryColumn: series.seriesIndex * 2,
      valueColumn: series.seriesIndex * 2 + 1,
      rows: series.categories.length,
    };
  }
  return {
    categoryColumn: 0,
    valueColumn: series.seriesIndex + 1,
    rows: series.categories.length,
  };
}

function categoryLabelRange(
  group: Extract<ParsedChartGroup, { kind: Exclude<ChartKind, "scatter"> }>,
  series: ParsedChartSeries,
  cells: SeriesCells,
): CategoryLabelRange | undefined {
  if (
    group.kind !== "doughnut" ||
    group.data.kind !== "independent-categories" ||
    series.categories.length === 0
  ) {
    return undefined;
  }
  const seriesLabels = mergeDataLabels(group.dataLabels, series.dataLabels);
  const pointBase = seriesLabels === false ? {} : seriesLabels;
  const seriesNeedsLabels = categoryLabelsEnabled(seriesLabels);
  const pointNeedsLabels = Object.values(series.points ?? {}).some((point) =>
    categoryLabelsEnabled(mergeDataLabels(pointBase, point.dataLabels)),
  );
  if (!seriesNeedsLabels && !pointNeedsLabels) return undefined;
  const categoryColumn = columnName(cells.categoryColumn);
  return {
    formula: `${sheet}!$${categoryColumn}$2:$${categoryColumn}$${cells.rows + 1}`,
    categories: series.categories,
  };
}

export function chartUsesCategoryLabelRange(parsed: ParsedChart): boolean {
  for (const group of parsed.groups) {
    if (group.kind === "scatter") continue;
    if (
      group.kind !== "doughnut" ||
      group.data.kind !== "independent-categories"
    ) {
      continue;
    }
    for (const series of group.series) {
      if (
        categoryLabelRange(group, series, seriesCells(group, series)) !==
        undefined
      ) {
        return true;
      }
    }
  }
  return false;
}

function categoryChartSeries(
  group: Extract<ParsedChartGroup, { kind: Exclude<ChartKind, "scatter"> }>,
  series: ParsedChartSeries,
  seriesCells: ChartSeriesCells,
  groupIndex: number,
  textStyles: ChartTextStyleXml,
  channels: ErrorBarChannelMap,
  theme: ThemeContext,
): XmlElement {
  const category = seriesCells.category;
  const value = seriesCells.value;
  if (category === undefined || value === undefined) {
    throw new Error("Category chart series is missing workbook columns");
  }
  const categoryRange = categoryLabelRange(group, series, {
    categoryColumn: category.column,
    valueColumn: value.column,
    rows: series.categories.length,
  });
  const kind: ChartKind = group.kind;
  const bar = kind === "column" || kind === "bar";
  const lineChart = kind === "line";
  const pie = kind === "pie" || kind === "doughnut";
  const itemShape = seriesShape(kind, series.color, series.line, theme);
  const seriesExplosion =
    series.explosion ??
    (group.kind === "pie" || group.kind === "doughnut"
      ? group.explosion
      : undefined);
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
    el(
      "c:tx",
      null,
      strRef(`${sheet}!$${columnName(seriesCells.nameColumn)}$1`, [
        series.name,
      ]),
    ),
    itemShape,
    bar ? el("c:invertIfNegative", { val: false }) : undefined,
    lineChart ? chartMarker(series.marker, theme) : undefined,
    pie && seriesExplosion !== undefined
      ? el("c:explosion", { val: Math.round(seriesExplosion) })
      : undefined,
    points,
    dataLabels(group, series, groupIndex, textStyles, theme, categoryRange),
    (series.trendlines ?? []).map((trendline) =>
      trendlineXml(trendline, theme),
    ),
    series.errorBars === undefined
      ? undefined
      : errorBarXml(
          series.errorBars,
          "value",
          group.kind === "line" ? "y" : undefined,
          series.seriesIndex,
          channels,
          theme,
        ),
    // A true zero-observation series omits both optional data-source wrappers. An empty
    // wrapper is not equivalent, and the source must stay absent after native save.
    // Keep missing points and inherited labels intact even though WPS may label gaps as 0
    // or misrender empty rings beside filled ones.
    series.values.length === 0
      ? undefined
      : el(
          "c:cat",
          null,
          strRef(
            rangeFormula(category.column, category.pointCount),
            series.categories,
          ),
        ),
    series.values.length === 0
      ? undefined
      : el(
          "c:val",
          null,
          numRef(
            rangeFormula(value.column, category.pointCount),
            series.values,
          ),
        ),
    lineChart ? el("c:smooth", { val: series.smooth ?? false }) : undefined,
    categoryRange === undefined
      ? undefined
      : categoryRangeExtension(categoryRange),
  );
}

function scatterMarkerXml(
  group: Extract<ParsedChartGroup, { kind: "scatter" }>,
  series: ParsedScatterSeries,
  theme: ThemeContext,
): XmlElement {
  if (!group.markers || series.marker === false) {
    return el("c:marker", null, el("c:symbol", { val: "none" }));
  }
  if (series.marker === undefined || series.marker === true) {
    return chartMarker(
      {
        symbol: "circle",
        ...(series.color === undefined ? {} : { fill: series.color }),
      },
      theme,
    );
  }
  return chartMarker(
    {
      ...series.marker,
      ...(series.marker.fill === undefined && series.color !== undefined
        ? { fill: series.color }
        : {}),
    },
    theme,
  );
}

function scatterChartSeries(
  group: Extract<ParsedChartGroup, { kind: "scatter" }>,
  series: ParsedScatterSeries,
  seriesCells: ChartSeriesCells,
  groupIndex: number,
  textStyles: ChartTextStyleXml,
  channels: ErrorBarChannelMap,
  theme: ThemeContext,
): XmlElement {
  if (seriesCells.x === undefined || seriesCells.y === undefined) {
    throw new Error("Scatter chart series is missing workbook columns");
  }
  // MS-OE376 2.1.1642d: Office loads scatterStyle "marker" as "lineMarker" and
  // keeps a connecting line unless the series disables it. Markers-only intent
  // therefore needs an explicit series noFill rather than the group enum alone.
  const itemShape =
    group.line === undefined || group.line === false
      ? el("c:spPr", null, el("a:ln", null, el("a:noFill")))
      : seriesShape("line", series.color, series.line, theme);
  const points = Object.entries<ChartPoint>(series.points ?? {}).flatMap(
    ([pointIndex, point]) => {
      const shape = pointShape("scatter", point, theme);
      const marker =
        point.marker === undefined
          ? undefined
          : chartMarker(point.marker, theme);
      if (shape === undefined && marker === undefined) return [];
      return [
        el(
          "c:dPt",
          null,
          el("c:idx", { val: Number(pointIndex) }),
          marker,
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
    el(
      "c:tx",
      null,
      strRef(`${sheet}!$${columnName(seriesCells.nameColumn)}$1`, [
        series.name,
      ]),
    ),
    itemShape,
    scatterMarkerXml(group, series, theme),
    points,
    dataLabels(group, series, groupIndex, textStyles, theme),
    (series.trendlines ?? []).map((trendline) =>
      trendlineXml(trendline, theme),
    ),
    series.errorBars?.x === undefined
      ? undefined
      : errorBarXml(
          series.errorBars.x,
          "x",
          "x",
          series.seriesIndex,
          channels,
          theme,
        ),
    series.errorBars?.y === undefined
      ? undefined
      : errorBarXml(
          series.errorBars.y,
          "y",
          "y",
          series.seriesIndex,
          channels,
          theme,
        ),
    // Scatter X/Y are one paired data source; omit both wrappers for zero observations.
    series.x.length === 0
      ? undefined
      : el(
          "c:xVal",
          null,
          numRef(
            rangeFormula(seriesCells.x.column, seriesCells.x.pointCount),
            series.x,
            seriesCells.x.labelFormats,
          ),
        ),
    series.x.length === 0
      ? undefined
      : el(
          "c:yVal",
          null,
          numRef(
            rangeFormula(seriesCells.y.column, seriesCells.y.pointCount),
            series.y,
          ),
        ),
    group.line === "smooth" || group.line === "straight"
      ? el("c:smooth", { val: group.line === "smooth" })
      : undefined,
  );
}

function assertNever(value: never): never {
  throw new Error(`Unhandled chart model variant: ${JSON.stringify(value)}`);
}

function grouping(
  stacking: "clustered" | "standard" | "stacked" | "percent",
): string {
  switch (stacking) {
    case "clustered":
      return "clustered";
    case "standard":
      return "standard";
    case "stacked":
      return "stacked";
    case "percent":
      return "percentStacked";
    default:
      return assertNever(stacking);
  }
}

function scatterStyle(
  group: Extract<ParsedChartGroup, { kind: "scatter" }>,
): string {
  if (group.line === "smooth") {
    return group.markers ? "smoothMarker" : "smooth";
  }
  if (group.line === "straight") {
    return group.markers ? "lineMarker" : "line";
  }
  return "marker";
}

function chartGroup(
  group: ParsedChartGroup,
  groupIndex: number,
  cells: ChartCells,
  textStyles: ChartTextStyleXml,
  channels: ErrorBarChannelMap,
  theme: ThemeContext,
): XmlElement {
  const seriesCell = (seriesIndex: number): ChartSeriesCells => {
    const value = cells.series.get(seriesIndex);
    if (value === undefined) {
      throw new Error(
        `Chart series ${seriesIndex} is missing workbook columns`,
      );
    }
    return value;
  };
  switch (group.kind) {
    case "scatter": {
      const items = group.series.map((item) =>
        scatterChartSeries(
          group,
          item,
          seriesCell(item.seriesIndex),
          groupIndex,
          textStyles,
          channels,
          theme,
        ),
      );
      return el(
        "c:scatterChart",
        null,
        el("c:scatterStyle", { val: scatterStyle(group) }),
        el("c:varyColors", { val: false }),
        items,
        el("c:axId", { val: group.axes.x.id }),
        el("c:axId", { val: group.axes.y.id }),
      );
    }
    case "column":
    case "bar": {
      const items = group.series.map((item) =>
        categoryChartSeries(
          group,
          item,
          seriesCell(item.seriesIndex),
          groupIndex,
          textStyles,
          channels,
          theme,
        ),
      );
      const axes = [
        el("c:axId", { val: group.axes.category.id }),
        el("c:axId", { val: group.axes.value.id }),
      ];
      return el(
        "c:barChart",
        null,
        el("c:barDir", { val: group.kind === "bar" ? "bar" : "col" }),
        el("c:grouping", { val: grouping(group.stacking) }),
        el("c:varyColors", { val: false }),
        items,
        el("c:gapWidth", { val: Math.round(group.gapWidth) }),
        group.overlap !== undefined
          ? el("c:overlap", { val: Math.round(group.overlap) })
          : group.stacking === "clustered"
            ? undefined
            : el("c:overlap", { val: 100 }),
        axes,
      );
    }
    case "line": {
      const items = group.series.map((item) =>
        categoryChartSeries(
          group,
          item,
          seriesCell(item.seriesIndex),
          groupIndex,
          textStyles,
          channels,
          theme,
        ),
      );
      const axes = [
        el("c:axId", { val: group.axes.category.id }),
        el("c:axId", { val: group.axes.value.id }),
      ];
      return el(
        "c:lineChart",
        null,
        el("c:grouping", { val: grouping(group.stacking) }),
        el("c:varyColors", { val: false }),
        items,
        el("c:marker", { val: true }),
        axes,
      );
    }
    case "area": {
      const items = group.series.map((item) =>
        categoryChartSeries(
          group,
          item,
          seriesCell(item.seriesIndex),
          groupIndex,
          textStyles,
          channels,
          theme,
        ),
      );
      const axes = [
        el("c:axId", { val: group.axes.category.id }),
        el("c:axId", { val: group.axes.value.id }),
      ];
      return el(
        "c:areaChart",
        null,
        el("c:grouping", { val: grouping(group.stacking) }),
        el("c:varyColors", { val: false }),
        items,
        axes,
      );
    }
    case "pie":
    case "doughnut": {
      const items = group.series.map((item) =>
        categoryChartSeries(
          group,
          item,
          seriesCell(item.seriesIndex),
          groupIndex,
          textStyles,
          channels,
          theme,
        ),
      );
      return el(
        group.kind === "pie" ? "c:pieChart" : "c:doughnutChart",
        null,
        el("c:varyColors", { val: true }),
        items,
        el("c:firstSliceAng", { val: group.firstSliceAngle }),
        group.kind === "doughnut"
          ? el("c:holeSize", { val: group.holeSize })
          : undefined,
      );
    }
    default:
      return assertNever(group);
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

// Raw OOXML crossing owner: this field positions this axis inside its crossAx
// partner's domain. Author crossings are adapted to this partner owner in parse.
function axisCrossing(crossing: ParsedAxisCrossing): XmlElement {
  return crossing.kind === "crossesAt"
    ? el("c:crossesAt", { val: crossing.value })
    : el("c:crosses", { val: crossing.value });
}

const axisPositionCodes = {
  bottom: "b",
  top: "t",
  left: "l",
  right: "r",
} as const;

function axisCrossBetween(
  parsed: ParsedChart,
  axis: ParsedChartAxis,
): "between" | "midCat" {
  let categoryUse = false;
  let areaOnly = true;
  for (const group of parsed.groups) {
    if (
      group.kind === "scatter" ||
      group.kind === "pie" ||
      group.kind === "doughnut"
    ) {
      continue;
    }
    if (group.axes.value !== axis) continue;
    categoryUse = true;
    if (group.kind !== "area") areaOnly = false;
  }
  // Scatter value axes keep the previous midCat candidate; category combinations use between
  // unless every category group sharing the axis is an area chart.
  if (!categoryUse) return "midCat";
  return areaOnly ? "midCat" : "between";
}

function axisXml(
  parsed: ParsedChart,
  axis: ParsedChartAxis,
  projectedScatterXAxes: ReadonlySet<number>,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement {
  switch (axis.kind) {
    case "category": {
      const spec = axis.spec;
      return el(
        "c:catAx",
        null,
        el("c:axId", { val: axis.id }),
        axisScaling(spec, false),
        el("c:delete", { val: spec.visible === false }),
        el("c:axPos", { val: axisPositionCodes[axis.position] }),
        gridlines(spec, "major", theme),
        gridlines(spec, "minor", theme),
        axisTitle(spec, textStyles, axis.name),
        el("c:numFmt", {
          formatCode: "General",
          sourceLinked: true,
        }),
        el("c:majorTickMark", { val: spec.majorTickMark ?? "none" }),
        el("c:minorTickMark", { val: spec.minorTickMark ?? "none" }),
        el("c:tickLblPos", { val: spec.tickLabelPosition ?? "nextTo" }),
        lineXml(spec.line, theme),
        raw(textStyles.axisStyles?.get(axis.name)),
        el("c:crossAx", { val: axis.crossAxisId }),
        axisCrossing(axis.crossing),
        [
          el("c:auto", { val: true }),
          el("c:lblAlgn", { val: "ctr" }),
          el("c:lblOffset", { val: 100 }),
          el("c:noMultiLvlLbl", { val: false }),
        ],
      );
    }
    case "value": {
      const spec = axis.spec;
      // Source-formatted X cells also affect linked ticks, so keep an unspecified X axis at General.
      const isolateDefaultXFormat =
        spec.numberFormat === undefined && projectedScatterXAxes.has(axis.id);
      return el(
        "c:valAx",
        null,
        el("c:axId", { val: axis.id }),
        axisScaling(spec, true),
        el("c:delete", { val: spec.visible === false }),
        el("c:axPos", { val: axisPositionCodes[axis.position] }),
        gridlines(spec, "major", theme),
        gridlines(spec, "minor", theme),
        axisTitle(spec, textStyles, axis.name),
        el("c:numFmt", {
          formatCode: isolateDefaultXFormat
            ? "General"
            : (spec.numberFormat ?? "General"),
          sourceLinked:
            spec.numberFormat === undefined && !isolateDefaultXFormat,
        }),
        el("c:majorTickMark", { val: spec.majorTickMark ?? "none" }),
        el("c:minorTickMark", { val: spec.minorTickMark ?? "none" }),
        el("c:tickLblPos", { val: spec.tickLabelPosition ?? "nextTo" }),
        lineXml(spec.line, theme),
        raw(textStyles.axisStyles?.get(axis.name)),
        el("c:crossAx", { val: axis.crossAxisId }),
        axisCrossing(axis.crossing),
        [
          el("c:crossBetween", { val: axisCrossBetween(parsed, axis) }),
          spec.majorUnit === undefined
            ? undefined
            : el("c:majorUnit", { val: spec.majorUnit }),
          spec.minorUnit === undefined
            ? undefined
            : el("c:minorUnit", { val: spec.minorUnit }),
        ],
      );
    }
    default:
      return assertNever(axis);
  }
}

function axes(
  parsed: ParsedChart,
  cells: ChartCells,
  textStyles: ChartTextStyleXml,
  theme: ThemeContext,
): XmlElement[] {
  return parsed.axes.map((axis) =>
    axisXml(parsed, axis, cells.projectedScatterXAxes, textStyles, theme),
  );
}

const legendPositions = {
  top: "t",
  bottom: "b",
  left: "l",
  right: "r",
  topRight: "tr",
} as const;

/**
 * An independent-category doughnut legend explains the first declared ring's ordered
 * category positions. Native legend entries beyond that domain would otherwise describe
 * positions no first-ring category owns; delete only those trailing entries.
 */
function independentLegendEntries(parsed: ParsedChart): XmlElement[] {
  if (parsed.groups.length !== 1) return [];
  const group = parsed.groups[0]!;
  if (
    group.kind !== "doughnut" ||
    group.data.kind !== "independent-categories"
  ) {
    return [];
  }
  const first = group.series[0];
  if (first === undefined) return [];
  const firstCount = first.categories.length;
  const maximumCount = Math.max(
    firstCount,
    ...group.series.map((series) => series.categories.length),
  );
  return Array.from(
    { length: Math.max(0, maximumCount - firstCount) },
    (_, offset) =>
      el(
        "c:legendEntry",
        null,
        el("c:idx", { val: firstCount + offset }),
        el("c:delete", { val: true }),
      ),
  );
}

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
  const cells = chartCells(parsed);
  const channels = errorBarChannels(parsed, cells.nextColumn + 1);
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
            chartGroup(group, groupIndex, cells, textStyles, channels, theme),
          ),
          axes(parsed, cells, textStyles, theme),
          transparent(),
        ),
        parsed.legend === undefined || parsed.legend === false
          ? undefined
          : el(
              "c:legend",
              null,
              el("c:legendPos", { val: legendPositions[parsed.legend] }),
              independentLegendEntries(parsed),
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
  const cells = chartCells(parsed);
  const channels = [...errorBarChannels(parsed, cells.nextColumn + 1).values()];
  const xNumberFormats = [
    ...new Set(
      [...cells.series.values()].flatMap(
        (item) => item.x?.labelFormats.pointFormatCodes ?? [],
      ),
    ),
  ].filter((formatCode) => formatCode !== "General");
  const numberFormatIds = new Map(
    xNumberFormats.map((formatCode, index) => [formatCode, 164 + index]),
  );
  const numberFormatStyleIndices = new Map(
    xNumberFormats.map((formatCode, index) => [formatCode, index + 1]),
  );
  const independent =
    parsed.groups.length === 1 &&
    parsed.groups[0]!.kind === "doughnut" &&
    parsed.groups[0]!.data.kind === "independent-categories";
  const headerCells: XmlElement[] = [inlineString("A1", " ")];
  if (independent) {
    for (const item of parsed.series as readonly ParsedChartSeries[]) {
      const itemCells = cells.series.get(item.seriesIndex);
      if (itemCells?.value === undefined) {
        throw new Error(
          `Chart series ${item.seriesIndex} is missing workbook columns`,
        );
      }
      headerCells.push(
        inlineString(`${columnName(itemCells.value.column)}1`, item.name),
      );
    }
  } else {
    for (const { column } of [
      ...cells.categoryColumns.values(),
      ...cells.axislessCategoryColumns.values(),
    ]) {
      if (column !== 0) {
        headerCells.push(inlineString(`${columnName(column)}1`, " "));
      }
    }
    for (const group of parsed.groups) {
      for (const series of group.series) {
        const item = cells.series.get(series.seriesIndex);
        if (item === undefined) {
          throw new Error(
            `Chart series ${series.seriesIndex} is missing workbook columns`,
          );
        }
        headerCells.push(
          inlineString(`${columnName(item.nameColumn)}1`, series.name),
        );
        if (group.kind === "scatter" && item.y !== undefined) {
          headerCells.push(inlineString(`${columnName(item.y.column)}1`, "Y"));
        }
      }
    }
  }
  for (const channel of channels) {
    headerCells.push(
      inlineString(`${columnName(channel.column - 1)}1`, channel.header),
    );
  }
  const header = el("row", { r: 1 }, headerCells);
  let rowXml: XmlElement[];
  if (independent) {
    const rows: XmlElement[][] = [];
    const series = parsed.series as readonly ParsedChartSeries[];
    const rowCount = Math.max(
      0,
      ...series.map((item) => item.categories.length),
    );
    for (let row = 0; row < rowCount; row += 1) rows.push([]);
    for (const item of series) {
      const itemCells = cells.series.get(item.seriesIndex);
      if (itemCells?.category === undefined || itemCells.value === undefined) {
        throw new Error(
          `Chart series ${item.seriesIndex} is missing workbook columns`,
        );
      }
      for (const [row, category] of item.categories.entries()) {
        rows[row]!.push(
          inlineString(
            `${columnName(itemCells.category.column)}${row + 2}`,
            category,
          ),
        );
        const value = item.values[row];
        if (value !== null && value !== undefined) {
          rows[row]!.push(
            el(
              "c",
              { r: `${columnName(itemCells.value.column)}${row + 2}` },
              el("v", null, String(value)),
            ),
          );
        }
      }
    }
    rowXml = rows.map((items, index) => el("row", { r: index + 2 }, items));
  } else {
    const rows = new Map<number, XmlElement[]>();
    const addCell = (row: number, cell: XmlElement | undefined) => {
      if (cell === undefined) return;
      const items = rows.get(row);
      if (items === undefined) rows.set(row, [cell]);
      else items.push(cell);
    };
    const categoryRanges = [
      ...cells.categoryColumns.values(),
      ...cells.axislessCategoryColumns.values(),
    ];
    for (const { column, categories } of categoryRanges) {
      for (const [index, category] of categories.entries()) {
        const row = index + 2;
        addCell(row, inlineString(`${columnName(column)}${row}`, category));
      }
    }
    for (const group of parsed.groups) {
      for (const series of group.series) {
        const item = cells.series.get(series.seriesIndex);
        if (item === undefined) continue;
        if (group.kind === "scatter") {
          const scatter = series as ParsedScatterSeries;
          for (let index = 0; index < scatter.x.length; index += 1) {
            const row = index + 2;
            const x = scatter.x[index];
            const y = scatter.y[index];
            if (item.x !== undefined) {
              const reference = `${columnName(item.x.column)}${row}`;
              const formatCode =
                item.x.labelFormats.pointFormatCodes[index] ?? "General";
              const styleIndex = numberFormatStyleIndices.get(formatCode);
              if (x !== null && x !== undefined) {
                addCell(
                  row,
                  el(
                    "c",
                    {
                      r: reference,
                      s: styleIndex,
                    },
                    el("v", null, String(x)),
                  ),
                );
              } else if (styleIndex !== undefined) {
                // Keep the blank slot formatted for a future workbook edit without adding a value.
                addCell(row, el("c", { r: reference, s: styleIndex }));
              }
            }
            if (y !== null && y !== undefined && item.y !== undefined) {
              addCell(
                row,
                el(
                  "c",
                  { r: `${columnName(item.y.column)}${row}` },
                  el("v", null, String(y)),
                ),
              );
            }
          }
          continue;
        }
        const category = series as ParsedChartSeries;
        for (let index = 0; index < category.values.length; index += 1) {
          const value = category.values[index];
          if (
            value === null ||
            value === undefined ||
            item.value === undefined
          ) {
            continue;
          }
          const row = index + 2;
          addCell(
            row,
            el(
              "c",
              { r: `${columnName(item.value.column)}${row}` },
              el("v", null, String(value)),
            ),
          );
        }
      }
    }
    for (const channel of channels) {
      for (let index = 0; index < channel.pointCount; index += 1) {
        const value = channel.values[index];
        if (value === null || value === undefined) continue;
        addCell(
          index + 2,
          el(
            "c",
            { r: `${columnName(channel.column - 1)}${index + 2}` },
            el("v", null, String(value)),
          ),
        );
      }
    }
    rowXml = [...rows.entries()]
      .sort(([left], [right]) => left - right)
      .map(([row, items]) => el("row", { r: row }, items));
  }
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
      el("sheetData", null, header, rowXml),
    ),
    "xl/styles.xml": el(
      "styleSheet",
      { xmlns: spreadsheet },
      xNumberFormats.length === 0
        ? undefined
        : el(
            "numFmts",
            { count: xNumberFormats.length },
            xNumberFormats.map((formatCode) =>
              el("numFmt", {
                numFmtId: numberFormatIds.get(formatCode),
                formatCode,
              }),
            ),
          ),
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
        { count: xNumberFormats.length + 1 },
        el("xf", { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }),
        xNumberFormats.map((formatCode) =>
          el("xf", {
            numFmtId: numberFormatIds.get(formatCode),
            fontId: 0,
            fillId: 0,
            borderId: 0,
            xfId: 0,
            applyNumberFormat: true,
          }),
        ),
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
