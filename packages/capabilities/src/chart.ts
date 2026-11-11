import {
  CapabilityError,
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  xmlEscape,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityLayoutContract,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityOutput,
  type CapabilityParseContext,
  type CapabilityParseResult,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type ObjectCapabilityModule,
  type ThemeContext,
} from "@react-pptx/core/capability";
import { CommonNode } from "@react-pptx/core/document";

import type {
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisReferences,
  ChartCategoryAxis,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartGroupData,
  ChartKind,
  ChartLine,
  ChartLineDash,
  ChartPoint,
  ChartProps,
  ChartSeries,
  ChartTickLabelPosition,
  ChartTickMark,
  ChartTitle,
  ChartValueAxis,
  SharedCategoryData,
} from "./chart-types.ts";
import {
  chartFrameXml,
  chartPartXml,
  chartWorkbook,
  type ChartTextStyleXml,
} from "./chart-xml.ts";
import {
  fontPresentationContributions,
  type FontEmbeddingMode,
  type FontEmbeddingUse,
} from "./font-embedding.ts";
import { motionAnnotations } from "./motion-api.ts";
import {
  writeChartTextProperties,
  writeChartTitle,
  type TextResourceClaim,
} from "./text-service.ts";
import type { TextStyle } from "./text-service.ts";

const chartRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const packageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const chartContentType =
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const workbookContentType =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const maximumShapeExtentEmu = 2_147_483_647;
// DrawingML a:ln/@w upper limit is 20,116,800 EMU; validate converted value before serializing.
const maximumLineWidthEmu = 20_116_800;
// Current formal validation profile accepts rounded values in range 0–2,147,483,647; not a general OOXML upper limit.
const maximumExplosionForCurrentValidationProfile = 2_147_483_647;
const categoryAxisId = 111111;
const valueAxisId = 222222;

const chartOptions: Readonly<Record<ChartKind, readonly string[]>> = {
  column: ["stacking", "gapWidth", "overlap"],
  bar: ["stacking", "gapWidth", "overlap"],
  line: ["stacking"],
  area: ["stacking"],
  pie: ["explosion", "firstSliceAngle"],
  doughnut: ["explosion", "firstSliceAngle", "holeSize"],
};
const allGroupOptions = [...new Set(Object.values(chartOptions).flat())];
const chartTextStyleKeys = [
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "color",
  "lineHeight",
  "verticalAlign",
  "align",
  "spacingAfter",
  "letterSpacing",
  "textTransform",
  "textDecoration",
  "textShadow",
  "lang",
] as const;
const positionValues = new Set([
  "center",
  "insideEnd",
  "insideBase",
  "outsideEnd",
  "bestFit",
]);
const lineDashValues = new Set<ChartLineDash>([
  "solid",
  "dash",
  "dot",
  "dashDot",
  "longDash",
  "longDashDot",
  "longDashDotDot",
]);
const tickMarkValues = new Set<ChartTickMark>(["none", "in", "out", "cross"]);
const tickLabelPositionValues = new Set<ChartTickLabelPosition>([
  "nextTo",
  "high",
  "low",
  "none",
]);
const crossingValues = new Set<ChartAxisCrossing>(["autoZero", "max", "min"]);
const displayBlanksAsValues = new Set<ChartDisplayBlanksAs>([
  "gap",
  "zero",
  "span",
]);
const cartesianKinds = new Set<ChartKind>(["column", "bar", "line", "area"]);

function fail(code: string, message: string): never {
  throw new CapabilityError(code, message);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function onlyKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
  label: string,
): void {
  const unknown = Object.keys(value).find((key) => !allowed.includes(key));
  if (unknown !== undefined) {
    fail(
      "unknown-parameter",
      `${label} contains unimplemented parameter ${unknown}`,
    );
  }
}

function optionalName(value: unknown, label: "name"): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    fail("invalid-parameter", `${label} must be a non-empty string`);
  }
  return value;
}

function coordinate(value: unknown, label: "x" | "y"): number {
  if (value === undefined) return 0;
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < -1_000_000 ||
    value > 1_000_000
  ) {
    fail(
      "invalid-coordinate",
      `${label} must be a finite pixel value between -1000000 and 1000000`,
    );
  }
  return value;
}

function extent(value: unknown, label: "width" | "height"): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    fail(
      "invalid-coordinate",
      `${label} must be a finite positive pixel value`,
    );
  }
  const emu = pxToEmu(value);
  if (emu < 1 || emu > maximumShapeExtentEmu) {
    fail(
      "invalid-coordinate",
      `${label} converted value must be within closed interval of 1 to ${maximumShapeExtentEmu} EMU`,
    );
  }
  return value;
}

function textOnlyStyle(value: unknown): TextStyle | undefined {
  if (!isRecord(value)) return undefined;
  return Object.fromEntries(
    chartTextStyleKeys.flatMap((key) =>
      value[key] === undefined ? [] : [[key, value[key]]],
    ),
  );
}

function optionalColor(value: unknown, label: string): ColorValue | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value === "string" ||
    (isRecord(value) && ("slot" in value || "color" in value))
  ) {
    return value as ColorValue;
  }
  return fail("invalid-color", `${label} must be a color value or theme token`);
}

function finiteNumber(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail("invalid-parameter", `${label} must be a finite number`);
  }
  return value;
}

type RoundedRangeBoundary = "input-and-result" | "rounded-result";

function roundedRange(
  value: unknown,
  label: string,
  minimum: number,
  maximum: number,
  boundary: RoundedRangeBoundary = "input-and-result",
): number | undefined {
  if (value === undefined) return undefined;
  const number = finiteNumber(value, label);
  const rounded = Math.round(number);
  if (
    (boundary === "input-and-result" &&
      (number < minimum || number > maximum)) ||
    rounded < minimum ||
    rounded > maximum
  ) {
    fail(
      "invalid-chart-option",
      boundary === "rounded-result"
        ? `${label} rounded value must be within range of ${minimum} to ${maximum}`
        : `${label} must be within format range of ${minimum} to ${maximum}`,
    );
  }
  return rounded;
}

function xmlText(value: string, label: string): string {
  try {
    xmlEscape(value);
  } catch {
    fail(
      "invalid-xml-character",
      `${label} contains characters not allowed in XML 1.0`,
    );
  }
  return value;
}

function lineStyle(value: unknown, label: string): ChartLine | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) fail("invalid-style", `${label} must be an object`);
  onlyKeys(value, ["color", "width", "dash"], label);
  const color = optionalColor(value["color"], `${label}.color`);
  const width = value["width"];
  if (width !== undefined) {
    if (typeof width !== "number" || !Number.isFinite(width) || width <= 0) {
      fail(
        "invalid-style",
        `${label}.width must be a finite positive pixel value`,
      );
    }
    if (pxToEmu(width) > maximumLineWidthEmu) {
      fail(
        "invalid-style",
        `${label}.width converted value exceeds writable line width range`,
      );
    }
  }
  const dash = value["dash"];
  if (dash !== undefined && !lineDashValues.has(dash as ChartLineDash)) {
    fail(
      "invalid-style",
      `${label}.dash must be solid, dash, dot, dashDot, longDash, longDashDot, or longDashDotDot`,
    );
  }
  return {
    ...(color === undefined ? {} : { color }),
    ...(width === undefined ? {} : { width }),
    ...(dash === undefined ? {} : { dash: dash as ChartLineDash }),
  };
}

function chartTitleValue(
  value: unknown,
  label: string,
): ChartTitle | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") {
    if (value.trim() === "") {
      fail("invalid-parameter", `${label} must be a non-empty string`);
    }
    xmlText(value, label);
    return value;
  }
  if (!isRecord(value)) {
    fail(
      "invalid-parameter",
      `${label} must be a non-empty string or an object with text and optional style`,
    );
  }
  onlyKeys(value, ["text", "style"], label);
  const text = value["text"];
  if (typeof text !== "string" || text.trim() === "") {
    fail("invalid-parameter", `${label}.text must be a non-empty string`);
  }
  xmlText(text, `${label}.text`);
  const style = value["style"];
  if (style !== undefined) {
    if (!isRecord(style))
      fail("invalid-style", `${label}.style must be an object`);
    onlyKeys(style, chartTextStyleKeys, `${label}.style`);
  }
  return {
    text,
    ...(style === undefined ? {} : { style: style as TextStyle }),
  };
}

function axis(
  value: unknown,
  label: string,
  expected: "category" | "value",
): ChartAxis {
  if (!isRecord(value)) fail("invalid-axis", `${label} must be an object`);
  onlyKeys(
    value,
    [
      "kind",
      "visible",
      "min",
      "max",
      "majorUnit",
      "minorUnit",
      "numberFormat",
      "logBase",
      "gridlines",
      "minorGridlines",
      "line",
      "style",
      "title",
      "majorTickMark",
      "minorTickMark",
      "tickLabelPosition",
      "reversed",
      "crosses",
      "crossesAt",
    ],
    label,
  );
  const kind = value["kind"];
  if (kind !== expected) {
    fail(
      "invalid-chart-axis-role",
      `${label} must declare kind "${expected}" as referenced by its chart group`,
    );
  }
  if (expected === "category") {
    const unsupported = [
      "min",
      "max",
      "majorUnit",
      "minorUnit",
      "numberFormat",
      "logBase",
    ].find((key) => value[key] !== undefined);
    if (unsupported !== undefined) {
      fail(
        "invalid-axis",
        `${label}.${unsupported} is only applicable to value axis`,
      );
    }
  }
  const visible = value["visible"];
  if (visible !== undefined && typeof visible !== "boolean") {
    fail("invalid-axis", `${label}.visible must be a boolean`);
  }
  const min =
    expected === "value" && value["min"] !== undefined
      ? finiteNumber(value["min"], `${label}.min`)
      : undefined;
  const max =
    expected === "value" && value["max"] !== undefined
      ? finiteNumber(value["max"], `${label}.max`)
      : undefined;
  const majorUnit =
    expected === "value" && value["majorUnit"] !== undefined
      ? finiteNumber(value["majorUnit"], `${label}.majorUnit`)
      : undefined;
  const minorUnit =
    expected === "value" && value["minorUnit"] !== undefined
      ? finiteNumber(value["minorUnit"], `${label}.minorUnit`)
      : undefined;
  if (min !== undefined && max !== undefined && min >= max) {
    fail("invalid-axis", `${label}.min must be less than max`);
  }
  if (majorUnit !== undefined && majorUnit <= 0) {
    fail("invalid-axis", `${label}.majorUnit must be greater than 0`);
  }
  if (minorUnit !== undefined && minorUnit <= 0) {
    fail("invalid-axis", `${label}.minorUnit must be greater than 0`);
  }
  const logBase =
    expected === "value" && value["logBase"] !== undefined
      ? finiteNumber(value["logBase"], `${label}.logBase`)
      : undefined;
  if (logBase !== undefined && (logBase < 2 || logBase > 1000)) {
    fail("invalid-axis", `${label}.logBase must be within range of 2 to 1000`);
  }
  if (logBase !== undefined) {
    if (min !== undefined && min <= 0) {
      fail(
        "invalid-axis",
        `${label}.min must be greater than 0 for a logarithmic axis`,
      );
    }
    if (max !== undefined && max <= 0) {
      fail(
        "invalid-axis",
        `${label}.max must be greater than 0 for a logarithmic axis`,
      );
    }
  }
  const numberFormat = expected === "value" ? value["numberFormat"] : undefined;
  if (
    numberFormat !== undefined &&
    (typeof numberFormat !== "string" || numberFormat.trim() === "")
  ) {
    fail("invalid-axis", `${label}.numberFormat must be a non-empty string`);
  }
  if (typeof numberFormat === "string") {
    xmlText(numberFormat, `${label}.numberFormat`);
  }
  const gridlines = value["gridlines"];
  if (
    gridlines !== undefined &&
    typeof gridlines !== "boolean" &&
    !isRecord(gridlines)
  ) {
    fail("invalid-axis", `${label}.gridlines must be a boolean or line style`);
  }
  const minorGridlines = value["minorGridlines"];
  if (
    minorGridlines !== undefined &&
    typeof minorGridlines !== "boolean" &&
    !isRecord(minorGridlines)
  ) {
    fail(
      "invalid-axis",
      `${label}.minorGridlines must be a boolean or line style`,
    );
  }
  const line = value["line"];
  if (line !== undefined && typeof line !== "boolean" && !isRecord(line)) {
    fail("invalid-axis", `${label}.line must be a boolean or line style`);
  }
  const majorTickMark = value["majorTickMark"];
  if (
    majorTickMark !== undefined &&
    !tickMarkValues.has(majorTickMark as ChartTickMark)
  ) {
    fail(
      "invalid-axis",
      `${label}.majorTickMark must be none, in, out, or cross`,
    );
  }
  const minorTickMark = value["minorTickMark"];
  if (
    minorTickMark !== undefined &&
    !tickMarkValues.has(minorTickMark as ChartTickMark)
  ) {
    fail(
      "invalid-axis",
      `${label}.minorTickMark must be none, in, out, or cross`,
    );
  }
  const tickLabelPosition = value["tickLabelPosition"];
  if (
    tickLabelPosition !== undefined &&
    !tickLabelPositionValues.has(tickLabelPosition as ChartTickLabelPosition)
  ) {
    fail(
      "invalid-axis",
      `${label}.tickLabelPosition must be nextTo, high, low, or none`,
    );
  }
  const reversed = value["reversed"];
  if (reversed !== undefined && typeof reversed !== "boolean") {
    fail("invalid-axis", `${label}.reversed must be a boolean`);
  }
  const crosses = value["crosses"];
  if (
    crosses !== undefined &&
    !crossingValues.has(crosses as ChartAxisCrossing)
  ) {
    fail("invalid-axis", `${label}.crosses must be autoZero, max, or min`);
  }
  const rawCrossesAt = value["crossesAt"];
  let crossesAt: number | undefined;
  if (rawCrossesAt !== undefined) {
    crossesAt = finiteNumber(rawCrossesAt, `${label}.crossesAt`);
    if (expected === "category") {
      if (!Number.isSafeInteger(crossesAt) || crossesAt < 1) {
        fail(
          "invalid-axis",
          `${label}.crossesAt must be a positive integer category position`,
        );
      }
    } else if (logBase !== undefined && crossesAt <= 0) {
      fail(
        "invalid-axis",
        `${label}.crossesAt must be greater than 0 for a logarithmic axis`,
      );
    }
  }
  if (crosses !== undefined && crossesAt !== undefined) {
    fail(
      "invalid-axis",
      `${label} must declare crosses or crossesAt, not both`,
    );
  }
  const title = chartTitleValue(value["title"], `${label}.title`);
  const common = {
    ...(visible === undefined ? {} : { visible }),
    ...(gridlines === undefined
      ? {}
      : {
          gridlines:
            typeof gridlines === "boolean"
              ? gridlines
              : lineStyle(gridlines, `${label}.gridlines`)!,
        }),
    ...(minorGridlines === undefined
      ? {}
      : {
          minorGridlines:
            typeof minorGridlines === "boolean"
              ? minorGridlines
              : lineStyle(minorGridlines, `${label}.minorGridlines`)!,
        }),
    ...(line === undefined
      ? {}
      : {
          line:
            typeof line === "boolean"
              ? line
              : lineStyle(line, `${label}.line`)!,
        }),
    ...(value["style"] === undefined
      ? {}
      : { style: value["style"] as TextStyle }),
    ...(title === undefined ? {} : { title }),
    ...(majorTickMark === undefined
      ? {}
      : { majorTickMark: majorTickMark as ChartTickMark }),
    ...(minorTickMark === undefined
      ? {}
      : { minorTickMark: minorTickMark as ChartTickMark }),
    ...(tickLabelPosition === undefined
      ? {}
      : { tickLabelPosition: tickLabelPosition as ChartTickLabelPosition }),
    ...(reversed === undefined ? {} : { reversed }),
    ...(crosses === undefined ? {} : { crosses: crosses as ChartAxisCrossing }),
    ...(crossesAt === undefined ? {} : { crossesAt }),
  };
  if (expected === "category") {
    return { kind: expected, ...common } satisfies ChartCategoryAxis;
  }
  return {
    kind: expected,
    ...common,
    ...(min === undefined ? {} : { min }),
    ...(max === undefined ? {} : { max }),
    ...(majorUnit === undefined ? {} : { majorUnit }),
    ...(minorUnit === undefined ? {} : { minorUnit }),
    ...(numberFormat === undefined ? {} : { numberFormat }),
    ...(logBase === undefined ? {} : { logBase }),
  } satisfies ChartValueAxis;
}

function sharedCategoryData(value: unknown, label: string): SharedCategoryData {
  if (!isRecord(value)) fail("invalid-data", `${label} must be an object`);
  onlyKeys(value, ["kind", "categories"], label);
  const kind = value["kind"];
  if (kind === "independent-categories") {
    fail(
      "unsupported-chart-data",
      `${label}.kind "independent-categories" is not supported by the migrated chart kinds`,
    );
  }
  if (kind !== "shared-categories") {
    fail(
      "invalid-data",
      `${label}.kind must be "shared-categories" for this chart group`,
    );
  }
  const categories = value["categories"];
  if (!Array.isArray(categories) || categories.length === 0) {
    fail(
      "invalid-data",
      `${label}.categories must contain at least one category`,
    );
  }
  for (const [index, category] of categories.entries()) {
    if (typeof category !== "string") {
      fail("invalid-data", `${label}.categories[${index}] must be a string`);
    }
  }
  return {
    kind: "shared-categories",
    categories: categories as string[],
  };
}

function chartSeriesList(
  value: unknown,
  categories: readonly string[],
  kind: ChartKind,
  label: string,
): ChartSeries[] {
  if (!Array.isArray(value) || value.length === 0) {
    fail("invalid-data", `${label} must contain at least one series`);
  }
  return value.map((item, seriesIndex) => {
    const seriesLabel = `${label}[${seriesIndex}]`;
    if (!isRecord(item)) {
      fail("invalid-data", `${seriesLabel} must be an object`);
    }
    onlyKeys(item, ["name", "values", "color", "line", "points"], seriesLabel);
    const name = item["name"];
    const values = item["values"];
    if (typeof name !== "string" || name.trim() === "") {
      fail("invalid-data", `${seriesLabel}.name must be a non-empty string`);
    }
    if (!Array.isArray(values) || values.length !== categories.length) {
      fail(
        "invalid-data",
        `Series "${name}" has ${Array.isArray(values) ? values.length : 0} values for ${categories.length} categories`,
      );
    }
    for (let valueIndex = 0; valueIndex < values.length; valueIndex += 1) {
      const value = values[valueIndex];
      if (
        value !== null &&
        (typeof value !== "number" || !Number.isFinite(value))
      ) {
        fail(
          "invalid-data",
          `Series "${name}" value ${valueIndex + 1} must be a finite number or null`,
        );
      }
      if (
        (kind === "pie" || kind === "doughnut") &&
        value !== null &&
        value < 0
      ) {
        fail(
          "invalid-data",
          `Series "${name}" pie value ${valueIndex + 1} cannot be negative`,
        );
      }
    }
    const pointsValue = item["points"];
    let points: Record<number, ChartPoint> | undefined;
    if (pointsValue !== undefined) {
      if (!isRecord(pointsValue)) {
        fail(
          "invalid-data",
          `Series "${name}" points must be an object indexed by category position`,
        );
      }
      points = {};
      for (const [rawIndex, pointValue] of Object.entries(pointsValue)) {
        const pointIndex = Number(rawIndex);
        if (
          !/^\d+$/u.test(rawIndex) ||
          !Number.isSafeInteger(pointIndex) ||
          pointIndex < 0 ||
          pointIndex >= categories.length
        ) {
          fail(
            "invalid-data",
            `Series "${name}" contains out-of-bounds data point index ${rawIndex}`,
          );
        }
        if (!isRecord(pointValue)) {
          fail(
            "invalid-data",
            `Series "${name}" data point ${rawIndex} must be an object`,
          );
        }
        onlyKeys(
          pointValue,
          ["color", "line", "labelColor"],
          `Series "${name}" data point ${rawIndex}`,
        );
        const pointColor = optionalColor(
          pointValue["color"],
          `Series "${name}" data point ${rawIndex}.color`,
        );
        const pointLine = lineStyle(
          pointValue["line"],
          `Series "${name}" data point ${rawIndex}.line`,
        );
        const labelColor = optionalColor(
          pointValue["labelColor"],
          `Series "${name}" data point ${rawIndex}.labelColor`,
        );
        points[pointIndex] = {
          ...(pointColor === undefined ? {} : { color: pointColor }),
          ...(pointLine === undefined ? {} : { line: pointLine }),
          ...(labelColor === undefined ? {} : { labelColor }),
        };
      }
    }
    const color = optionalColor(item["color"], `Series "${name}".color`);
    const line = lineStyle(item["line"], `Series "${name}".line`);
    return {
      name,
      values: values as readonly (number | null)[],
      ...(color === undefined ? {} : { color }),
      ...(line === undefined ? {} : { line }),
      ...(points === undefined ? {} : { points }),
    };
  });
}

function labels(
  value: unknown,
  kind: ChartKind,
  stacking?: "stacked" | "percent",
): false | ChartDataLabels | undefined {
  if (value === undefined || value === false) return value;
  if (!isRecord(value))
    fail(
      "invalid-data-labels",
      "Chart group dataLabels must be an object or false",
    );
  onlyKeys(
    value,
    ["value", "percent", "category", "position", "numberFormat", "style"],
    "Chart group dataLabels",
  );
  for (const key of ["value", "percent", "category"] as const) {
    if (value[key] !== undefined && typeof value[key] !== "boolean") {
      fail(
        "invalid-data-labels",
        `Chart group dataLabels.${key} must be a boolean`,
      );
    }
  }
  const position = value["position"];
  if (position !== undefined && !positionValues.has(position as string)) {
    fail(
      "invalid-data-labels",
      typeof position === "string"
        ? `Unsupported data label position ${position}`
        : "Chart group dataLabels.position must be a supported string",
    );
  }
  if (
    position !== undefined &&
    kind !== "column" &&
    kind !== "bar" &&
    kind !== "pie"
  ) {
    const reason =
      kind === "doughnut"
        ? "Office does not allow dLblPos for doughnutChart series data labels"
        : kind === "area"
          ? "Office does not allow dLblPos for areaChart series data labels"
          : "Chart generation does not expose explicit data label positions for line charts";
    fail("invalid-data-labels", reason);
  }
  if (kind === "pie" && position === "insideBase") {
    fail(
      "invalid-data-labels",
      'Office does not allow dLblPos="inBase" for pieChart series data labels',
    );
  }
  if ((kind === "column" || kind === "bar") && position === "bestFit") {
    fail(
      "invalid-data-labels",
      'Office does not allow dLblPos="bestFit" for barChart series data labels',
    );
  }
  if (
    (kind === "column" || kind === "bar") &&
    stacking !== undefined &&
    position === "outsideEnd"
  ) {
    fail(
      "invalid-data-labels",
      'Office does not allow dLblPos="outEnd" for stacked barChart series data labels',
    );
  }
  const numberFormat = value["numberFormat"];
  if (
    numberFormat !== undefined &&
    (typeof numberFormat !== "string" || numberFormat.trim() === "")
  ) {
    fail(
      "invalid-data-labels",
      "Chart group dataLabels.numberFormat must be a non-empty string",
    );
  }
  if (typeof numberFormat === "string") {
    xmlText(numberFormat, "Chart group dataLabels.numberFormat");
  }
  return {
    ...(value["value"] === undefined
      ? {}
      : { value: value["value"] as boolean }),
    ...(value["percent"] === undefined
      ? {}
      : { percent: value["percent"] as boolean }),
    ...(value["category"] === undefined
      ? {}
      : { category: value["category"] as boolean }),
    ...(position === undefined
      ? {}
      : { position: position as NonNullable<ChartDataLabels["position"]> }),
    ...(numberFormat === undefined ? {} : { numberFormat: numberFormat }),
    ...(value["style"] === undefined
      ? {}
      : { style: value["style"] as TextStyle }),
  };
}

function plotArea(
  value: unknown,
  box?: { readonly width: number; readonly height: number },
): ChartProps["plotArea"] {
  if (value === undefined) return undefined;
  if (!isRecord(value))
    fail("invalid-plot-area", "Chart.plotArea must be an object");
  onlyKeys(value, ["left", "top", "width", "height"], "Chart.plotArea");
  const entries = ["left", "top", "width", "height"] as const;
  const result = Object.fromEntries(
    entries.map((key) => [
      key,
      finiteNumber(value[key], `Chart.plotArea.${key}`),
    ]),
  ) as Record<(typeof entries)[number], number>;
  if (
    result.left < 0 ||
    result.top < 0 ||
    result.width <= 0 ||
    result.height <= 0 ||
    (box !== undefined &&
      (result.left + result.width > box.width ||
        result.top + result.height > box.height))
  ) {
    fail(
      "invalid-plot-area",
      "Chart.plotArea must be inside chart box with positive width and height",
    );
  }
  return result;
}

function axisReferences(
  value: unknown,
  label: string,
): ChartAxisReferences | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value))
    fail("invalid-chart-axes", `${label} must be an object`);
  onlyKeys(value, ["category", "value"], label);
  const category = value["category"];
  const valueAxis = value["value"];
  if (
    typeof category !== "string" ||
    category.trim() === "" ||
    typeof valueAxis !== "string" ||
    valueAxis.trim() === ""
  ) {
    fail(
      "invalid-chart-axes",
      `${label} must name both a category and a value axis`,
    );
  }
  return { category, value: valueAxis };
}

type ParsedGroupDraft = {
  readonly kind: ChartKind;
  readonly data: ChartGroupData;
  readonly series: readonly ChartSeries[];
  readonly axes?: ChartAxisReferences;
  readonly dataLabels?: false | ChartDataLabels;
  readonly stacking?: "stacked" | "percent";
  readonly gapWidth?: number;
  readonly overlap?: number;
  readonly explosion?: number;
  readonly firstSliceAngle?: number;
  readonly holeSize?: number;
};

function chartGroupDrafts(value: unknown): ParsedGroupDraft[] {
  if (!Array.isArray(value) || value.length === 0) {
    fail(
      "invalid-chart-groups",
      "Chart.groups must contain at least one group",
    );
  }
  return value.map((item, groupIndex) => {
    const label = `Chart.groups[${groupIndex}]`;
    if (!isRecord(item)) {
      fail("invalid-chart-groups", `${label} must be an object`);
    }
    onlyKeys(
      item,
      ["kind", "data", "series", "axes", "dataLabels", ...allGroupOptions],
      label,
    );
    const kind = item["kind"];
    if (typeof kind !== "string" || !Object.hasOwn(chartOptions, kind)) {
      fail("unsupported-chart-kind", `Unsupported chart kind ${String(kind)}`);
    }
    const chartKind = kind as ChartKind;
    for (const option of allGroupOptions) {
      if (
        item[option] !== undefined &&
        !chartOptions[chartKind].includes(option)
      ) {
        fail(
          "invalid-chart-option",
          `${chartKind} chart group does not support ${option}`,
        );
      }
    }
    const data = sharedCategoryData(item["data"], `${label}.data`);
    const series = chartSeriesList(
      item["series"],
      data.categories,
      chartKind,
      `${label}.series`,
    );
    if (
      (chartKind === "pie" || chartKind === "doughnut") &&
      series.length !== 1
    ) {
      fail(
        "invalid-data",
        `${chartKind} chart currently only supports a single data series`,
      );
    }
    const stack = item["stacking"];
    if (stack !== undefined && stack !== "stacked" && stack !== "percent") {
      fail("invalid-stacking", `${label}.stacking must be stacked or percent`);
    }
    const labelsValue = labels(item["dataLabels"], chartKind, stack);
    const gapWidth = item["gapWidth"];
    if (
      gapWidth !== undefined &&
      (typeof gapWidth !== "number" ||
        !Number.isFinite(gapWidth) ||
        gapWidth < 0 ||
        gapWidth > 500)
    ) {
      fail("invalid-gap-width", `${label}.gapWidth must be between 0 and 500`);
    }
    const overlap = item["overlap"];
    if (
      overlap !== undefined &&
      (typeof overlap !== "number" ||
        !Number.isFinite(overlap) ||
        overlap < -100 ||
        overlap > 100)
    ) {
      fail("invalid-overlap", `${label}.overlap must be between -100 and 100`);
    }
    const explosion =
      chartKind === "pie" || chartKind === "doughnut"
        ? roundedRange(
            item["explosion"],
            `${label}.explosion`,
            0,
            maximumExplosionForCurrentValidationProfile,
            "rounded-result",
          )
        : undefined;
    const firstSliceAngle =
      chartKind === "pie" || chartKind === "doughnut"
        ? roundedRange(
            item["firstSliceAngle"],
            `${label}.firstSliceAngle`,
            0,
            360,
          )
        : undefined;
    const holeSize =
      chartKind === "doughnut"
        ? roundedRange(item["holeSize"], `${label}.holeSize`, 10, 90)
        : undefined;
    const axes = axisReferences(item["axes"], `${label}.axes`);
    if (!cartesianKinds.has(chartKind) && axes !== undefined) {
      fail(
        "invalid-chart-axes",
        `${chartKind} chart group does not use category or value axes`,
      );
    }
    return {
      kind: chartKind,
      data,
      series,
      ...(axes === undefined ? {} : { axes }),
      ...(labelsValue === undefined ? {} : { dataLabels: labelsValue }),
      ...(stack === undefined ? {} : { stacking: stack }),
      ...(gapWidth === undefined ? {} : { gapWidth }),
      ...(overlap === undefined ? {} : { overlap }),
      ...(explosion === undefined ? {} : { explosion }),
      ...(firstSliceAngle === undefined ? {} : { firstSliceAngle }),
      ...(holeSize === undefined ? {} : { holeSize }),
    };
  });
}

export type ParsedChartSeries = ChartSeries & {
  readonly groupIndex: number;
  readonly groupSeriesIndex: number;
  /** Derived once from group order and group-local series order. */
  readonly seriesIndex: number;
};

export type ParsedChartAxis = {
  readonly name: string;
  readonly role: "category" | "value";
  readonly id: number;
  readonly crossAxisId: number;
  readonly spec: ChartAxis;
};

export type ParsedChartGroup = {
  readonly kind: ChartKind;
  readonly data: ChartGroupData;
  readonly series: readonly ParsedChartSeries[];
  readonly axes?: {
    readonly category: ParsedChartAxis;
    readonly value: ParsedChartAxis;
  };
  readonly dataLabels?: false | ChartDataLabels;
  readonly stacking?: "stacked" | "percent";
  readonly gapWidth?: number;
  readonly overlap?: number;
  readonly explosion?: number;
  readonly firstSliceAngle?: number;
  readonly holeSize?: number;
};

export type ParsedChart = {
  readonly name?: string;
  readonly morphKey?: string;
  readonly link?: LinkTarget;
  readonly declaredX?: number;
  readonly declaredY?: number;
  readonly declaredWidth?: number;
  readonly declaredHeight?: number;
  readonly axes: readonly ParsedChartAxis[];
  readonly groups: readonly ParsedChartGroup[];
  /** Flattened series view with the same objects owned by their groups. */
  readonly series: readonly ParsedChartSeries[];
  readonly categories: readonly string[];
  readonly title?: ChartTitle;
  readonly displayBlanksAs?: ChartDisplayBlanksAs;
  readonly legend?: ChartProps["legend"];
  readonly plotArea?: ChartProps["plotArea"];
  readonly style?: TextStyle & LayoutStyle;
  readonly enter?: unknown;
  readonly emphasis?: unknown;
  readonly motion?: unknown;
  readonly contract: CapabilityLayoutContract;
  readonly rawProps: Readonly<Record<string, unknown>>;
};

function parsedAxes(
  value: unknown,
  groups: readonly ParsedGroupDraft[],
): {
  axes: ParsedChartAxis[];
  groupAxes: (ParsedChartGroup["axes"] | undefined)[];
} {
  const cartesian = groups.every((group) => cartesianKinds.has(group.kind));
  if (!cartesian) {
    if (groups.length !== 1) {
      fail(
        "unsupported-chart-group-combination",
        "Pie and doughnut charts support one chart group",
      );
    }
    if (value !== undefined) {
      fail(
        "invalid-chart-axes",
        "Pie and doughnut charts do not use category or value axes",
      );
    }
    return { axes: [], groupAxes: [undefined] };
  }
  if (new Set(groups.map(({ kind }) => kind)).size !== 1) {
    fail(
      "unsupported-chart-group-combination",
      "Mixed chart group kinds are not supported by this migration",
    );
  }

  let entries: { name: string; spec: ChartAxis }[];
  const explicit = value !== undefined;
  if (value === undefined) {
    entries = [
      { name: "category", spec: { kind: "category" } },
      { name: "value", spec: { kind: "value" } },
    ];
  } else {
    if (!isRecord(value)) fail("invalid-axis", "Chart.axes must be an object");
    entries = Object.entries(value).map(([name, spec]) => {
      if (name.trim() === "") {
        fail("invalid-axis", "Chart.axes names must be non-empty strings");
      }
      if (!isRecord(spec)) {
        fail("invalid-axis", `Chart.axes.${name} must be an object`);
      }
      const kind = spec["kind"];
      if (kind !== "category" && kind !== "value") {
        fail(
          "invalid-chart-axis-role",
          `Chart.axes.${name}.kind must be "category" or "value"`,
        );
      }
      return {
        name,
        spec: axis(spec, `Chart.axes.${name}`, kind),
      };
    });
    if (entries.length === 0) {
      fail(
        "invalid-axis",
        "Chart.axes must define a category and a value axis",
      );
    }
  }

  const category = entries.find(({ spec }) => spec.kind === "category");
  const valueAxis = entries.find(({ spec }) => spec.kind === "value");
  if (category === undefined || valueAxis === undefined) {
    fail(
      "invalid-axis",
      "Chart.axes must define exactly one category axis and one value axis",
    );
  }
  if (
    entries.some(
      ({ name }) => name !== category.name && name !== valueAxis.name,
    )
  ) {
    fail(
      "unsupported-chart-axis",
      "This migration supports one category axis and one value axis per chart",
    );
  }
  const categoryParsed: ParsedChartAxis = {
    name: category.name,
    role: "category",
    id: categoryAxisId,
    crossAxisId: valueAxisId,
    spec: category.spec,
  };
  const valueParsed: ParsedChartAxis = {
    name: valueAxis.name,
    role: "value",
    id: valueAxisId,
    crossAxisId: categoryAxisId,
    spec: valueAxis.spec,
  };
  const groupAxes = groups.map((group) => {
    let references = group.axes;
    if (references === undefined) {
      if (
        explicit &&
        !(category.name === "category" && valueAxis.name === "value")
      ) {
        fail(
          "invalid-chart-axes",
          "Chart groups must reference chart-level axes when custom axis names are declared",
        );
      }
      references = { category: category.name, value: valueAxis.name };
    }
    if (
      references.category !== "category" &&
      references.category !== category.name
    ) {
      if (entries.some(({ name }) => name === references.category)) {
        fail(
          "invalid-chart-axis-role",
          `Chart group category axis "${references.category}" is not a category axis`,
        );
      }
      fail(
        "unknown-chart-axis",
        `Chart group references unknown category axis "${references.category}"`,
      );
    }
    if (references.value !== "value" && references.value !== valueAxis.name) {
      if (entries.some(({ name }) => name === references.value)) {
        fail(
          "invalid-chart-axis-role",
          `Chart group value axis "${references.value}" is not a value axis`,
        );
      }
      fail(
        "unknown-chart-axis",
        `Chart group references unknown value axis "${references.value}"`,
      );
    }
    if (
      references.category !== category.name ||
      references.value !== valueAxis.name
    ) {
      fail(
        "unknown-chart-axis",
        "Chart group axis references must match the declared chart-level axes",
      );
    }
    return { category: categoryParsed, value: valueParsed };
  });
  return { axes: [categoryParsed, valueParsed], groupAxes };
}

function displayBlanksAs(
  value: unknown,
  groups: readonly ParsedGroupDraft[],
): ChartDisplayBlanksAs | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value !== "string" ||
    !displayBlanksAsValues.has(value as ChartDisplayBlanksAs)
  ) {
    fail(
      "invalid-chart-missing-values",
      "Chart.displayBlanksAs must be gap, zero, or span",
    );
  }
  const requested = value as ChartDisplayBlanksAs;
  const unsupported = groups.find((group) => {
    const allowed =
      group.kind === "line"
        ? group.stacking === undefined
          ? new Set<ChartDisplayBlanksAs>(["gap", "zero", "span"])
          : new Set<ChartDisplayBlanksAs>(["zero"])
        : group.kind === "area"
          ? group.stacking === undefined
            ? new Set<ChartDisplayBlanksAs>(["gap", "zero"])
            : new Set<ChartDisplayBlanksAs>(["zero"])
          : new Set<ChartDisplayBlanksAs>(["gap", "zero"]);
    return !allowed.has(requested);
  });
  if (unsupported !== undefined) {
    fail(
      "invalid-chart-missing-values",
      `${unsupported.kind}${unsupported.stacking === undefined ? "" : ` ${unsupported.stacking}`} chart does not support displayBlanksAs "${requested}"`,
    );
  }
  return requested;
}

/** Chart module validates input and generates native chart parts with embedded workbook. */
export const chartCapability: ObjectCapabilityModule = {
  declaration: "rp:chart",

  parse(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): CapabilityParseResult<ParsedChart> {
    if (input.children.length > 0) {
      fail(
        "unsupported-children",
        "Chart does not accept children; chart text and data should be passed in groups",
      );
    }
    onlyKeys(
      input.props,
      [
        "id",
        "name",
        "morphKey",
        "link",
        "x",
        "y",
        "width",
        "height",
        "axes",
        "groups",
        "title",
        "displayBlanksAs",
        "legend",
        "plotArea",
        "style",
        "enter",
        "emphasis",
        "motion",
      ],
      "Chart",
    );

    const allowedStyleKeys = new Set<string>([
      ...layoutStyleKeys,
      ...chartTextStyleKeys,
    ]);
    const contract = parseCapabilityLayoutContract({
      nativeKind: "graphicFrame",
      props: input.props,
      allowedStyleKeys,
      allowPadding: false,
      path: context.path,
    });

    const drafts = chartGroupDrafts(input.props["groups"]);
    if (drafts.length > 1) {
      for (const group of drafts.slice(1)) {
        if (
          group.data.categories.length !== drafts[0]!.data.categories.length
        ) {
          fail(
            "chart-category-domain-mismatch",
            "Chart groups sharing a category axis must have the same ordered category domain",
          );
        }
        for (const [index, category] of group.data.categories.entries()) {
          if (category !== drafts[0]!.data.categories[index]) {
            fail(
              "chart-category-domain-mismatch",
              "Chart groups sharing a category axis must have the same ordered category domain",
            );
          }
        }
      }
    }
    const { axes, groupAxes } = parsedAxes(input.props["axes"], drafts);

    let seriesIndex = 0;
    const groups: ParsedChartGroup[] = drafts.map((draft, groupIndex) => {
      const series = draft.series.map((item, groupSeriesIndex) => ({
        ...item,
        groupIndex,
        groupSeriesIndex,
        seriesIndex: seriesIndex++,
      }));
      const parsedGroupAxes = groupAxes[groupIndex];
      return {
        kind: draft.kind,
        data: draft.data,
        series,
        ...(parsedGroupAxes === undefined ? {} : { axes: parsedGroupAxes }),
        ...(draft.dataLabels === undefined
          ? {}
          : { dataLabels: draft.dataLabels }),
        ...(draft.stacking === undefined ? {} : { stacking: draft.stacking }),
        ...(draft.gapWidth === undefined ? {} : { gapWidth: draft.gapWidth }),
        ...(draft.overlap === undefined ? {} : { overlap: draft.overlap }),
        ...(draft.explosion === undefined
          ? {}
          : { explosion: draft.explosion }),
        ...(draft.firstSliceAngle === undefined
          ? {}
          : { firstSliceAngle: draft.firstSliceAngle }),
        ...(draft.holeSize === undefined ? {} : { holeSize: draft.holeSize }),
      };
    });
    const series = groups.flatMap((group) => group.series);
    const categories = groups[0]!.data.categories;

    const name = optionalName(input.props["name"], "name");
    if (name !== undefined) xmlText(name, "Chart.name");

    const rawWidth =
      contract.layoutStyle.width ??
      (input.props["width"] as number | undefined);
    const rawHeight =
      contract.layoutStyle.height ??
      (input.props["height"] as number | undefined);
    const declaredWidth =
      rawWidth !== undefined ? extent(rawWidth, "width") : undefined;
    const declaredHeight =
      rawHeight !== undefined ? extent(rawHeight, "height") : undefined;
    const declaredX =
      input.props["x"] !== undefined
        ? coordinate(input.props["x"], "x")
        : undefined;
    const declaredY =
      input.props["y"] !== undefined
        ? coordinate(input.props["y"], "y")
        : undefined;

    const legend = input.props["legend"];
    if (
      legend !== undefined &&
      legend !== false &&
      legend !== "top" &&
      legend !== "bottom" &&
      legend !== "left" &&
      legend !== "right"
    ) {
      fail(
        "invalid-legend",
        "Chart.legend must be false, top, bottom, left, or right",
      );
    }
    const checkedPlotArea = plotArea(
      input.props["plotArea"],
      declaredWidth !== undefined && declaredHeight !== undefined
        ? { width: declaredWidth, height: declaredHeight }
        : undefined,
    );
    const title = chartTitleValue(input.props["title"], "Chart.title");
    const blanks = displayBlanksAs(input.props["displayBlanksAs"], drafts);

    const parsed: ParsedChart = {
      axes,
      groups,
      series,
      categories,
      contract,
      rawProps: input.props,
      ...(name === undefined ? {} : { name }),
      ...(title === undefined ? {} : { title }),
      ...(blanks === undefined ? {} : { displayBlanksAs: blanks }),
      ...(input.props["morphKey"] !== undefined
        ? { morphKey: input.props["morphKey"] as string }
        : {}),
      ...(input.props["link"] !== undefined
        ? { link: input.props["link"] as LinkTarget }
        : {}),
      ...(declaredX === undefined ? {} : { declaredX }),
      ...(declaredY === undefined ? {} : { declaredY }),
      ...(declaredWidth === undefined ? {} : { declaredWidth }),
      ...(declaredHeight === undefined ? {} : { declaredHeight }),
      ...(legend === undefined ? {} : { legend: legend }),
      ...(checkedPlotArea === undefined ? {} : { plotArea: checkedPlotArea }),
      ...(input.props["style"] === undefined
        ? {}
        : { style: input.props["style"] as TextStyle & LayoutStyle }),
      ...(input.props["enter"] === undefined
        ? {}
        : { enter: input.props["enter"] }),
      ...(input.props["emphasis"] === undefined
        ? {}
        : { emphasis: input.props["emphasis"] }),
      ...(input.props["motion"] === undefined
        ? {}
        : { motion: input.props["motion"] }),
    };

    return { contract, parsed };
  },

  measureLayout(
    parsed: ParsedChart,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement {
    const declaredWidth = parsed.declaredWidth;
    const declaredHeight = parsed.declaredHeight;
    const width =
      request.availableWidth ??
      declaredWidth ??
      (declaredHeight === undefined ? 320 : declaredHeight * (16 / 9));
    const height = declaredHeight ?? width * (9 / 16);
    return {
      width,
      height,
      minContentWidth: declaredWidth ?? 0,
      minContentHeight: declaredHeight ?? 0,
    };
  },

  compile(parsed: ParsedChart, context: CapabilityContext): CapabilityOutput {
    const x = context.layoutBox?.x ?? parsed.declaredX ?? 0;
    const y = context.layoutBox?.y ?? parsed.declaredY ?? 0;
    const width =
      context.layoutBox?.width ??
      parsed.declaredWidth ??
      (parsed.declaredHeight === undefined
        ? 320
        : parsed.declaredHeight * (16 / 9));
    const height =
      context.layoutBox?.height ?? parsed.declaredHeight ?? width * (9 / 16);
    const box = { x, y, width, height };

    if (parsed.plotArea !== undefined) {
      if (
        parsed.plotArea.left < 0 ||
        parsed.plotArea.top < 0 ||
        parsed.plotArea.width <= 0 ||
        parsed.plotArea.height <= 0 ||
        parsed.plotArea.left + parsed.plotArea.width > box.width ||
        parsed.plotArea.top + parsed.plotArea.height > box.height
      ) {
        fail(
          "invalid-plot-area",
          "Chart.plotArea must be inside chart box with positive width and height",
        );
      }
    }

    const name = context.morphName ?? parsed.name ?? `Chart ${context.id - 1}`;
    const firstKind = parsed.groups[0]!.kind;
    const annotations = motionAnnotations(parsed.rawProps, {
      kind: "chart",
      name,
      ...(parsed.rawProps["id"] === undefined
        ? {}
        : { authorId: parsed.rawProps["id"] as string }),
      textBody: false,
      filled: firstKind !== "line",
      box,
      rotationDegrees: 0,
      hasBackground: firstKind !== "line",
      chartSeriesCount: parsed.series.length,
      chartCategoryCount: parsed.categories.length,
    });
    const { textStyles, resourceClaims, fontUses, warnings } =
      chartResourceClaims(parsed, context.theme, context.options.embedFonts);
    const workbook = chartWorkbook(parsed);
    const relationshipId = context.registerPackagePart({
      part: {
        pathTemplate: "ppt/charts/chart{n}.xml",
        contentType: chartContentType,
      },
      relationshipType: chartRelationship,
      content(relatedRelationshipIds) {
        const [workbookRelationshipId] = relatedRelationshipIds;
        if (workbookRelationshipId === undefined) {
          fail(
            "invalid-part",
            "Chart part requires embedded workbook relationship",
          );
        }
        return chartPartXml(
          parsed,
          box,
          workbookRelationshipId,
          textStyles,
          context.theme,
        );
      },
      relatedParts: [
        {
          pathTemplate: "ppt/embeddings/Microsoft_Excel_Worksheet{n}.xlsx",
          contentType: workbookContentType,
          relationshipType: packageRelationship,
          bytes: workbook,
        },
      ],
    });
    return {
      xml: chartFrameXml(
        context.id,
        name,
        box,
        relationshipId,
        context.nativeLink,
      ),
      observation: {
        id: context.id,
        kind: "chart",
        name,
        box,
      },
      extensions: [],
      ...(annotations.length === 0 ? {} : { annotations }),
      resourceClaims,
      presentationContributions: fontPresentationContributions(
        fontUses,
        context.options.embedFonts,
      ),
      warnings,
    };
  },
};

const generalNumberCharacters = "0123456789.,-+eE";
const dateNameCharacters =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
// When simplified Chinese LCID is explicitly specified, date names use localized glyphs instead of English characters.
const dateNameCharactersByLocale: Readonly<Record<string, string>> = {
  "0804": "星期日一二三四五六七八九十年月",
};
const numberFormatLiteralCharacters = new Set("$+():^'{}<=>-/!&~ ");

function dateNameCharactersForLocale(
  locale: string | undefined,
): string | undefined {
  if (locale === undefined) return dateNameCharacters;
  const languageId = /^(?:0x)?([\da-f]{1,4})(?:,.*)?$/iu.exec(locale)?.[1];
  if (languageId === undefined) return undefined;
  const normalizedLanguageId = languageId.padStart(4, "0").toUpperCase();
  if (normalizedLanguageId === "0409") return dateNameCharacters;
  // Locales explicitly specified without registered name glyphs do not masquerade as English.
  return dateNameCharactersByLocale[normalizedLanguageId];
}

function hasDigitPlaceholderAfter(
  format: readonly string[],
  start: number,
): boolean {
  for (let index = start; index < format.length; index += 1) {
    const character = format[index]!;
    if (character === ";") return false;
    if (character === '"') {
      index += 1;
      while (index < format.length && format[index] !== '"') index += 1;
      continue;
    }
    if (character === "\\" || character === "_") {
      index += 1;
      continue;
    }
    if (character === "[") {
      while (index < format.length && format[index] !== "]") index += 1;
      continue;
    }
    if (character === "0" || character === "#" || character === "?") {
      return true;
    }
  }
  return false;
}

/** Claims glyphs that chart number formatting might display, without simulating application-level number formatting at build time. */
function numberLabelCharacters(
  numberFormat: string | undefined,
  percent = false,
): string {
  const characters = new Set<string>();
  const add = (value: string) => {
    for (const character of value) characters.add(character);
  };
  const format = Array.from(numberFormat ?? "");
  const defaultFormat =
    numberFormat === undefined ||
    numberFormat.trim().toLocaleLowerCase("en-US") === "general";
  let hasNumericValue = defaultFormat;
  let hasGeneralValue = defaultFormat;
  let sectionLocale: string | undefined;
  let sectionCount = 1;

  if (defaultFormat) add(generalNumberCharacters);
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index]!;
    if (character === '"') {
      for (index += 1; index < format.length; index += 1) {
        if (format[index] === '"') break;
        characters.add(format[index]!);
      }
      continue;
    }
    if (character === "\\" || character === "*") {
      const following = format[index + 1];
      if (following !== undefined) {
        characters.add(following);
        index += 1;
      }
      continue;
    }
    if (character === "_") {
      index += 1;
      continue;
    }
    if (character === "[") {
      const end = format.indexOf("]", index + 1);
      if (end < 0) continue;
      const directive = format.slice(index + 1, end).join("");
      if (directive.startsWith("$")) {
        const currencyAndLocale = directive.slice(1);
        const separator = currencyAndLocale.indexOf("-");
        const currency =
          separator < 0
            ? currencyAndLocale
            : currencyAndLocale.slice(0, separator);
        add(currency);
        sectionLocale =
          separator < 0 ? undefined : currencyAndLocale.slice(separator + 1);
      } else if (/^[hms]+$/iu.test(directive)) {
        hasNumericValue = true;
      }
      index = end;
      continue;
    }
    if (character === ";") {
      sectionCount += 1;
      if (sectionCount > 3) break;
      sectionLocale = undefined;
      continue;
    }
    if (character === "0" || character === "#" || character === "?") {
      if (character === "?") characters.add(" ");
      hasNumericValue = true;
      continue;
    }
    if (character === ",") {
      if (hasDigitPlaceholderAfter(format, index + 1))
        characters.add(character);
      continue;
    }
    if (character === "." || character === "%") {
      characters.add(character);
      continue;
    }
    if (character === "A" || character === "a") {
      const amPm = format.slice(index, index + 5).join("");
      const amPmHalf = format.slice(index, index + 3).join("");
      if (/^am\/pm$/iu.test(amPm)) {
        add(amPm.replace("/", ""));
        index += 4;
        continue;
      }
      if (/^a\/p$/iu.test(amPmHalf)) {
        add(amPmHalf.replace("/", ""));
        index += 2;
        continue;
      }
    }
    if (/[yYmdDhHsS]/u.test(character)) {
      let end = index + 1;
      while (
        end < format.length &&
        format[end]!.toLocaleLowerCase("en-US") ===
          character.toLocaleLowerCase("en-US")
      ) {
        end += 1;
      }
      const width = end - index;
      hasNumericValue = true;
      if (
        (character.toLocaleLowerCase("en-US") === "m" ||
          character.toLocaleLowerCase("en-US") === "d") &&
        width >= 3
      ) {
        const dateNames = dateNameCharactersForLocale(sectionLocale);
        if (dateNames !== undefined) add(dateNames);
      }
      index = end - 1;
      continue;
    }
    if (character === "G" || character === "g") {
      const general = format.slice(index, index + 7).join("");
      if (
        general.toLocaleLowerCase("en-US") === "general" &&
        (format[index + 7] === undefined || !/[a-z]/iu.test(format[index + 7]!))
      ) {
        add(generalNumberCharacters);
        hasNumericValue = true;
        hasGeneralValue = true;
        index += 6;
        continue;
      }
    }
    if (character === "E" || character === "e") {
      const sign = format[index + 1];
      if (
        (sign === "+" || sign === "-") &&
        hasDigitPlaceholderAfter(format, index + 2)
      ) {
        add(character + sign);
        hasNumericValue = true;
        index += 1;
        continue;
      }
    }
    if (numberFormatLiteralCharacters.has(character)) {
      characters.add(character);
      continue;
    }
    if (/[1-9]/u.test(character) || character.codePointAt(0)! > 0x7f) {
      characters.add(character);
    }
  }

  if (hasNumericValue) {
    add("0123456789");
    if (sectionCount === 1) characters.add("-");
  }
  if (hasGeneralValue) add("+eE");
  if (percent && numberFormat === undefined) characters.add("%");
  return [...characters].join("");
}

function groupDataLabelTexts(
  group: ParsedChartGroup,
  labels: ChartDataLabels | false | undefined,
): string[] {
  if (labels === undefined || labels === false) return [];
  const texts: string[] = [];
  const numericLabels =
    labels.value === true || labels.percent === true
      ? numberLabelCharacters(labels.numberFormat, labels.percent === true)
      : "";
  for (const series of group.series) {
    for (const [index, number] of series.values.entries()) {
      if (number === null) continue;
      if (numericLabels !== "") texts.push(numericLabels);
      if (labels.category === true) texts.push(group.data.categories[index]!);
    }
  }
  return texts;
}

function pointDataLabelTexts(
  group: ParsedChartGroup,
  labels: ChartDataLabels,
  index: number,
): string[] {
  if (group.series.some((series) => series.values[index] !== null)) {
    const texts: string[] = [];
    if (labels.value === true || labels.percent === true) {
      texts.push(
        numberLabelCharacters(labels.numberFormat, labels.percent === true),
      );
    }
    if (labels.category === true) texts.push(group.data.categories[index]!);
    return texts;
  }
  return [];
}

function mergeTextStyle(
  base: TextStyle | undefined,
  override: TextStyle,
): TextStyle {
  return { ...base, ...override };
}

function titleParts(
  value: ChartTitle | undefined,
): { readonly text: string; readonly style?: TextStyle } | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "string") return { text: value };
  return {
    text: value.text,
    ...(value.style === undefined ? {} : { style: value.style }),
  };
}

function chartResourceClaims(
  parsed: ParsedChart,
  theme: ThemeContext,
  embedFonts?: FontEmbeddingMode,
): {
  readonly textStyles: ChartTextStyleXml;
  readonly resourceClaims: readonly TextResourceClaim[];
  readonly fontUses: readonly FontEmbeddingUse[];
  readonly warnings: readonly {
    readonly code: string;
    readonly message: string;
  }[];
} {
  const categoryAxes = parsed.axes.filter(({ role }) => role === "category");
  const valueAxes = parsed.axes.filter(({ role }) => role === "value");
  const dataLabelText = parsed.groups.flatMap((group) =>
    group.dataLabels === undefined ||
    group.dataLabels === false ||
    group.dataLabels.style !== undefined
      ? []
      : groupDataLabelTexts(group, group.dataLabels),
  );
  const chartText = [
    ...parsed.categories,
    ...parsed.series.map(({ name }) => name),
    ...dataLabelText,
  ];
  const valueAxisCharacters =
    valueAxes.length === 0 || valueAxes[0]!.spec.visible === false
      ? []
      : [
          numberLabelCharacters(
            valueAxes[0]!.spec.kind === "value"
              ? valueAxes[0]!.spec.numberFormat
              : undefined,
            parsed.groups.some(({ stacking }) => stacking === "percent"),
          ),
        ];
  valueAxes.forEach(({ spec }) => {
    if (spec.visible !== false && spec.style === undefined) {
      chartText.push(...valueAxisCharacters);
    }
  });
  const chartStyle = textOnlyStyle(parsed.style);
  const chart = writeChartTextProperties(
    chartStyle,
    chartText,
    "Chart.style",
    theme,
    embedFonts,
  );
  const resourceClaims: TextResourceClaim[] = [...chart.resourceClaims];
  const fontUses: FontEmbeddingUse[] = chart.fontUses.map((use) => ({
    ...use,
    path: "Chart.style",
  }));
  const warnings: { readonly code: string; readonly message: string }[] = [];
  const warningKeys = new Set<string>();
  const addWarnings = (
    items: readonly { readonly code: string; readonly message: string }[],
  ) => {
    for (const item of items) {
      const key = `${item.code}\u0000${item.message}`;
      if (warningKeys.has(key)) continue;
      warningKeys.add(key);
      warnings.push(item);
    }
  };
  addWarnings(chart.warnings);
  const textStyles: {
    chart: string;
    title?: string;
    axisTitles?: Map<string, string>;
    categoryAxis?: string;
    valueAxis?: string;
    groupDataLabels?: Map<number, string>;
    pointLabels?: Map<string, string>;
  } = { chart: chart.xml };
  const chartTitle = titleParts(parsed.title);
  if (chartTitle !== undefined) {
    const output = writeChartTitle(
      chartTitle.text,
      mergeTextStyle(chartStyle, chartTitle.style ?? {}),
      "Chart.title",
      theme,
      embedFonts,
    );
    textStyles.title = output.xml;
    resourceClaims.push(...output.resourceClaims);
    fontUses.push(
      ...output.fontUses.map((use) => ({ ...use, path: "Chart.title" })),
    );
    addWarnings(output.warnings);
  }
  for (const { name, spec } of categoryAxes) {
    if (spec.style !== undefined) {
      const output = writeChartTextProperties(
        mergeTextStyle(chartStyle, spec.style),
        parsed.categories,
        `Chart.axes.${name}.style`,
        theme,
        embedFonts,
      );
      textStyles.categoryAxis = output.xml;
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.axes.${name}.style`,
        })),
      );
      addWarnings(output.warnings);
    }
    const axisTitle = titleParts(spec.title);
    if (axisTitle !== undefined) {
      const output = writeChartTitle(
        axisTitle.text,
        mergeTextStyle(
          mergeTextStyle(chartStyle, spec.style ?? {}),
          axisTitle.style ?? {},
        ),
        `Chart.axes.${name}.title`,
        theme,
        embedFonts,
      );
      textStyles.axisTitles ??= new Map();
      textStyles.axisTitles.set(name, output.xml);
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.axes.${name}.title`,
        })),
      );
      addWarnings(output.warnings);
    }
  }
  for (const { name, spec } of valueAxes) {
    if (spec.style !== undefined) {
      const output = writeChartTextProperties(
        mergeTextStyle(chartStyle, spec.style),
        valueAxisCharacters,
        `Chart.axes.${name}.style`,
        theme,
        embedFonts,
      );
      textStyles.valueAxis = output.xml;
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.axes.${name}.style`,
        })),
      );
      addWarnings(output.warnings);
    }
    const axisTitle = titleParts(spec.title);
    if (axisTitle !== undefined) {
      const output = writeChartTitle(
        axisTitle.text,
        mergeTextStyle(
          mergeTextStyle(chartStyle, spec.style ?? {}),
          axisTitle.style ?? {},
        ),
        `Chart.axes.${name}.title`,
        theme,
        embedFonts,
      );
      textStyles.axisTitles ??= new Map();
      textStyles.axisTitles.set(name, output.xml);
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.axes.${name}.title`,
        })),
      );
      addWarnings(output.warnings);
    }
  }
  for (const [groupIndex, group] of parsed.groups.entries()) {
    const groupLabels = group.dataLabels;
    if (groupLabels === undefined || groupLabels === false) continue;
    if (groupLabels.style !== undefined) {
      const output = writeChartTextProperties(
        mergeTextStyle(chartStyle, groupLabels.style),
        groupDataLabelTexts(group, groupLabels),
        `Chart.groups[${groupIndex}].dataLabels.style`,
        theme,
        embedFonts,
      );
      textStyles.groupDataLabels ??= new Map();
      textStyles.groupDataLabels.set(groupIndex, output.xml);
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: `Chart.groups[${groupIndex}].dataLabels.style`,
        })),
      );
      addWarnings(output.warnings);
    }
    for (const series of group.series) {
      for (const [pointIndex, point] of Object.entries(series.points ?? {})) {
        if (point.labelColor === undefined) continue;
        const index = Number(pointIndex);
        const number = series.values[index];
        if (number === null || number === undefined) continue;
        const pointText = pointDataLabelTexts(group, groupLabels, index);
        if (pointText.length === 0) continue;
        const output = writeChartTextProperties(
          mergeTextStyle(chartStyle, {
            ...groupLabels.style,
            color: point.labelColor,
          }),
          pointText,
          "Chart.dataLabels.point.style",
          theme,
          embedFonts,
        );
        textStyles.pointLabels ??= new Map();
        textStyles.pointLabels.set(
          `${series.seriesIndex}:${pointIndex}`,
          output.xml,
        );
        resourceClaims.push(...output.resourceClaims);
        fontUses.push(
          ...output.fontUses.map((use) => ({
            ...use,
            path: "Chart.dataLabels.point.style",
          })),
        );
        addWarnings(output.warnings);
      }
    }
  }
  const uniqueWarnings = new Map(
    warnings.map((warning) => [
      `${warning.code}\u0000${warning.message}`,
      warning,
    ]),
  );
  return {
    textStyles,
    resourceClaims,
    fontUses,
    warnings: [...uniqueWarnings.values()],
  };
}

export type ChartNode = CommonNode<
  "chart",
  ChartProps,
  NonNullable<ChartProps["style"]>
>;

/** Creates the editable native chart shared by the command factory and the Chart JSX element. */
export function createChartNode(props: ChartProps): ChartNode {
  return new CommonNode<"chart", ChartProps, NonNullable<ChartProps["style"]>>(
    "chart",
    "rp:chart",
    props,
  );
}

export type {
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisReferences,
  ChartCategoryAxis,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartGroup,
  ChartGroupData,
  ChartKind,
  ChartLine,
  ChartLineDash,
  ChartPoint,
  ChartProps,
  ChartSeries,
  ChartTextStyle,
  ChartTickLabelPosition,
  ChartTickMark,
  ChartTitle,
  ChartValueAxis,
  ColumnOrBarChartGroup,
  DoughnutChartGroup,
  LineOrAreaChartGroup,
  PieChartGroup,
  SharedCategoryData,
} from "./chart-types.ts";
