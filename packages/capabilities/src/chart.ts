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
import type { MorphProps } from "./morph-api.ts";
import { motionAnnotations, type MotionProps } from "./motion-api.ts";
import {
  writeChartTextProperties,
  type TextResourceClaim,
} from "./text-service.ts";
import type { TextStyle } from "./text-service.ts";

type ChartTextStyle = Omit<TextStyle, "textWrap">;

export type ChartKind = "column" | "bar" | "line" | "area" | "pie" | "doughnut";

export type ChartPoint = {
  readonly color?: ColorValue;
  /** Overrides data label text color for this data point. */
  readonly labelColor?: ColorValue;
};

export type ChartSeries = {
  readonly name: string;
  /** One numeric value per category; null is preserved as empty value in chart and workbook. */
  readonly values: readonly (number | null)[];
  /** Color of column, area, or line. */
  readonly color?: ColorValue;
  /** Overrides color or data label text color for individual points by category index. */
  readonly points?: Readonly<Record<number, ChartPoint>>;
};

export type ChartData = {
  readonly categories: readonly string[];
  readonly series: readonly ChartSeries[];
};

export type ChartLine = {
  readonly color?: ColorValue;
  /** Line width in px. */
  readonly width?: number;
};

export type ChartAxis = {
  /** Whether to show the axis and its labels; defaults to true. */
  readonly visible?: boolean;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  /** Excel number format, e.g. `#,##0`. */
  readonly numberFormat?: string;
  /** Major gridlines; true uses target application default line style. */
  readonly gridlines?: boolean | ChartLine;
  /** Axis line; false hides axis line. */
  readonly line?: boolean | ChartLine;
  /** Text style for axis labels. */
  readonly style?: ChartTextStyle;
};

type ChartDataLabelPosition =
  | "center"
  | "insideEnd"
  | "insideBase"
  | "outsideEnd"
  | "bestFit";

/** Data label settings; ChartProps narrows position by chart group and stacking. */
export type ChartDataLabels<
  Position extends ChartDataLabelPosition = ChartDataLabelPosition,
> = {
  readonly value?: boolean;
  readonly percent?: boolean;
  readonly category?: boolean;
  readonly position?: Position;
  readonly numberFormat?: string;
  readonly style?: ChartTextStyle;
};

type ChartBase = MotionProps &
  MorphProps & {
    readonly testId?: string;
    readonly part?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly data: ChartData;
    /** When omitted or false, hides the legend. */
    readonly legend?: false | "top" | "bottom" | "left" | "right";
    /** Internal plot area position relative to chart box, in px. */
    readonly plotArea?: {
      readonly left: number;
      readonly top: number;
      readonly width: number;
      readonly height: number;
    };
    /** Text style shared by chart legend, axes, and data labels. */
    readonly style?: ChartTextStyle & LayoutStyle;
  };

type CartesianOptions = {
  readonly valueAxis?: ChartAxis;
  readonly categoryAxis?: ChartAxis;
};

type DataLabelsWithoutPosition =
  | false
  | (Omit<ChartDataLabels, "position"> & { readonly position?: never });

type ColumnOrBarUnstackedOptions = {
  readonly stacking?: undefined;
  readonly dataLabels?:
    | false
    | ChartDataLabels<"center" | "insideEnd" | "insideBase" | "outsideEnd">;
};

type ColumnOrBarStackedOptions = {
  readonly stacking: "stacked" | "percent";
  readonly dataLabels?:
    | false
    | ChartDataLabels<"center" | "insideEnd" | "insideBase">;
};

type CartesianStacking = {
  readonly stacking?: "stacked" | "percent";
};

export type ColumnOrBarChartProps = ChartBase &
  CartesianOptions & {
    readonly kind: "column" | "bar";
    /** Gap between category groups as percentage of bar width, range 0–500. */
    readonly gapWidth?: number;
    /** Overlap percentage for bars in the same group, range -100–100. */
    readonly overlap?: number;
  } & (ColumnOrBarUnstackedOptions | ColumnOrBarStackedOptions);

export type LineOrAreaChartProps = ChartBase &
  CartesianOptions &
  CartesianStacking & {
    readonly kind: "line" | "area";
    readonly dataLabels?: DataLabelsWithoutPosition;
  };

type PieOptions = {
  /**
   * Radial explosion percentage pulling the slice outward.
   * Rounded integer must be within the 0–2,147,483,647 range supported by current formal validation.
   */
  readonly explosion?: number;
  /** Clockwise starting angle of the first slice in degrees (0–360), serialized as integer. */
  readonly firstSliceAngle?: number;
};

export type PieChartProps = ChartBase &
  PieOptions & {
    readonly kind: "pie";
    readonly dataLabels?:
      | false
      | ChartDataLabels<"center" | "insideEnd" | "outsideEnd" | "bestFit">;
  };

export type DoughnutChartProps = ChartBase &
  PieOptions & {
    readonly kind: "doughnut";
    /** Percentage size of doughnut chart center hole, range 10–90. */
    readonly holeSize?: number;
    readonly dataLabels?: DataLabelsWithoutPosition;
  };

export type ChartProps =
  | ColumnOrBarChartProps
  | LineOrAreaChartProps
  | PieChartProps
  | DoughnutChartProps;

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
const chartOptions: Readonly<Record<ChartKind, readonly string[]>> = {
  column: ["stacking", "valueAxis", "categoryAxis", "gapWidth", "overlap"],
  bar: ["stacking", "valueAxis", "categoryAxis", "gapWidth", "overlap"],
  line: ["stacking", "valueAxis", "categoryAxis"],
  area: ["stacking", "valueAxis", "categoryAxis"],
  pie: ["explosion", "firstSliceAngle"],
  doughnut: ["explosion", "firstSliceAngle", "holeSize"],
};
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
const allChartOptions = [...new Set(Object.values(chartOptions).flat())];
const positionValues = new Set([
  "center",
  "insideEnd",
  "insideBase",
  "outsideEnd",
  "bestFit",
]);

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

function optionalName(
  value: unknown,
  label: "testId" | "name" | "part",
): string | undefined {
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
  onlyKeys(value, ["color", "width"], label);
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
  return {
    ...(color === undefined ? {} : { color }),
    ...(width === undefined ? {} : { width }),
  };
}

function axis(
  value: unknown,
  label: string,
  kind: "category" | "value",
): ChartAxis | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) fail("invalid-axis", `${label} must be an object`);
  onlyKeys(
    value,
    [
      "visible",
      "min",
      "max",
      "majorUnit",
      "numberFormat",
      "gridlines",
      "line",
      "style",
    ],
    label,
  );
  if (kind === "category") {
    const unsupported = ["min", "max", "majorUnit", "numberFormat"].find(
      (key) => value[key] !== undefined,
    );
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
    value["min"] === undefined
      ? undefined
      : finiteNumber(value["min"], `${label}.min`);
  const max =
    value["max"] === undefined
      ? undefined
      : finiteNumber(value["max"], `${label}.max`);
  const majorUnit =
    value["majorUnit"] === undefined
      ? undefined
      : finiteNumber(value["majorUnit"], `${label}.majorUnit`);
  if (min !== undefined && max !== undefined && min >= max) {
    fail("invalid-axis", `${label}.min must be less than max`);
  }
  if (majorUnit !== undefined && majorUnit <= 0) {
    fail("invalid-axis", `${label}.majorUnit must be greater than 0`);
  }
  const numberFormat = value["numberFormat"];
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
  const line = value["line"];
  if (line !== undefined && typeof line !== "boolean" && !isRecord(line)) {
    fail("invalid-axis", `${label}.line must be a boolean or line style`);
  }
  return {
    ...(visible === undefined ? {} : { visible }),
    ...(min === undefined ? {} : { min }),
    ...(max === undefined ? {} : { max }),
    ...(majorUnit === undefined ? {} : { majorUnit }),
    ...(numberFormat === undefined ? {} : { numberFormat: numberFormat }),
    ...(gridlines === undefined
      ? {}
      : {
          gridlines:
            typeof gridlines === "boolean"
              ? gridlines
              : lineStyle(gridlines, `${label}.gridlines`)!,
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
  };
}

function chartData(value: unknown, kind: ChartKind): ChartData {
  if (!isRecord(value)) fail("invalid-data", "Chart.data must be an object");
  onlyKeys(value, ["categories", "series"], "Chart.data");
  const categories = value["categories"];
  const series = value["series"];
  if (!Array.isArray(categories) || categories.length === 0) {
    fail(
      "invalid-data",
      "Chart.data.categories must contain at least one category",
    );
  }
  for (const [index, category] of categories.entries()) {
    if (typeof category !== "string") {
      fail("invalid-data", `Chart.data.categories[${index}] must be a string`);
    }
  }
  if (!Array.isArray(series) || series.length === 0) {
    fail("invalid-data", "Chart.data.series must contain at least one series");
  }
  const checkedSeries: ChartSeries[] = series.map((item, seriesIndex) => {
    if (!isRecord(item)) {
      fail(
        "invalid-data",
        `Chart.data.series[${seriesIndex}] must be an object`,
      );
    }
    onlyKeys(
      item,
      ["name", "values", "color", "points"],
      `Chart.data.series[${seriesIndex}]`,
    );
    const name = item["name"];
    const values = item["values"];
    if (typeof name !== "string" || name.trim() === "") {
      fail(
        "invalid-data",
        `Chart.data.series[${seriesIndex}].name must be a non-empty string`,
      );
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
        typeof value === "number" &&
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
          `Series "${name}" points must be an object indexed by category index`,
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
          ["color", "labelColor"],
          `Series "${name}" data point ${rawIndex}`,
        );
        const pointColor = optionalColor(
          pointValue["color"],
          `Series "${name}" data point ${rawIndex}.color`,
        );
        const labelColor = optionalColor(
          pointValue["labelColor"],
          `Series "${name}" data point ${rawIndex}.labelColor`,
        );
        points[pointIndex] = {
          ...(pointColor === undefined ? {} : { color: pointColor }),
          ...(labelColor === undefined ? {} : { labelColor }),
        };
      }
    }
    const color = optionalColor(item["color"], `Series "${name}".color`);
    return {
      name,
      values: values as readonly (number | null)[],
      ...(color === undefined ? {} : { color }),
      ...(points === undefined ? {} : { points }),
    };
  });
  return { categories: categories as string[], series: checkedSeries };
}

function labels(
  value: unknown,
  kind: ChartKind,
  stacking?: "stacked" | "percent",
): false | ChartDataLabels | undefined {
  if (value === undefined || value === false) return value;
  if (!isRecord(value))
    fail("invalid-data-labels", "Chart.dataLabels must be an object or false");
  onlyKeys(
    value,
    ["value", "percent", "category", "position", "numberFormat", "style"],
    "Chart.dataLabels",
  );
  for (const key of ["value", "percent", "category"] as const) {
    if (value[key] !== undefined && typeof value[key] !== "boolean") {
      fail("invalid-data-labels", `Chart.dataLabels.${key} must be a boolean`);
    }
  }
  const position = value["position"];
  if (position !== undefined && !positionValues.has(position as string)) {
    fail(
      "invalid-data-labels",
      typeof position === "string"
        ? `Unsupported data label position ${position}`
        : "Chart.dataLabels.position must be a supported string",
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
      "Chart.dataLabels.numberFormat must be a non-empty string",
    );
  }
  if (typeof numberFormat === "string") {
    xmlText(numberFormat, "Chart.dataLabels.numberFormat");
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

function dataLabelTexts(
  data: ChartData,
  value: ChartDataLabels | false | undefined,
): string[] {
  if (value === undefined || value === false) return [];
  const texts: string[] = [];
  const numericLabels =
    value.value === true || value.percent === true
      ? numberLabelCharacters(value.numberFormat, value.percent === true)
      : "";
  for (const series of data.series) {
    for (const [index, number] of series.values.entries()) {
      if (number === null) continue;
      if (numericLabels !== "") texts.push(numericLabels);
      if (value.category === true) texts.push(data.categories[index]!);
    }
  }
  return texts;
}

function pointDataLabelTexts(
  data: ChartData,
  labels: ChartDataLabels,
  index: number,
): string[] {
  if (data.series.some((series) => series.values[index] !== null)) {
    const texts: string[] = [];
    if (labels.value === true || labels.percent === true) {
      texts.push(
        numberLabelCharacters(labels.numberFormat, labels.percent === true),
      );
    }
    if (labels.category === true) texts.push(data.categories[index]!);
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

function chartResourceClaims(
  props: ChartProps,
  data: ChartData,
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
  const cartesian =
    props.kind === "pie" || props.kind === "doughnut" ? undefined : props;
  const dataLabelText = dataLabelTexts(data, props.dataLabels);
  const chartText = [
    ...data.categories,
    ...data.series.map(({ name }) => name),
    ...(props.dataLabels === undefined ||
    props.dataLabels === false ||
    props.dataLabels.style === undefined
      ? dataLabelText
      : []),
  ];
  const valueAxisCharacters =
    cartesian?.valueAxis?.visible === false
      ? []
      : [
          numberLabelCharacters(
            cartesian?.valueAxis?.numberFormat,
            cartesian?.stacking === "percent",
          ),
        ];
  if (
    cartesian !== undefined &&
    cartesian.valueAxis?.visible !== false &&
    cartesian.valueAxis?.style === undefined
  ) {
    chartText.push(...valueAxisCharacters);
  }
  const chartStyle = textOnlyStyle(props.style);
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
    categoryAxis?: string;
    valueAxis?: string;
    dataLabels?: string;
    pointLabels?: Map<string, string>;
  } = { chart: chart.xml };
  if (cartesian?.categoryAxis?.style !== undefined) {
    const output = writeChartTextProperties(
      mergeTextStyle(chartStyle, cartesian.categoryAxis.style),
      data.categories,
      "Chart.categoryAxis.style",
      theme,
      embedFonts,
    );
    textStyles.categoryAxis = output.xml;
    resourceClaims.push(...output.resourceClaims);
    fontUses.push(
      ...output.fontUses.map((use) => ({
        ...use,
        path: "Chart.categoryAxis.style",
      })),
    );
    addWarnings(output.warnings);
  }
  if (cartesian?.valueAxis?.style !== undefined) {
    const output = writeChartTextProperties(
      mergeTextStyle(chartStyle, cartesian.valueAxis.style),
      valueAxisCharacters,
      "Chart.valueAxis.style",
      theme,
      embedFonts,
    );
    textStyles.valueAxis = output.xml;
    resourceClaims.push(...output.resourceClaims);
    fontUses.push(
      ...output.fontUses.map((use) => ({
        ...use,
        path: "Chart.valueAxis.style",
      })),
    );
    addWarnings(output.warnings);
  }
  if (props.dataLabels !== undefined && props.dataLabels !== false) {
    if (props.dataLabels.style !== undefined) {
      const output = writeChartTextProperties(
        mergeTextStyle(chartStyle, props.dataLabels.style),
        dataLabelTexts(data, props.dataLabels),
        "Chart.dataLabels.style",
        theme,
        embedFonts,
      );
      textStyles.dataLabels = output.xml;
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: "Chart.dataLabels.style",
        })),
      );
      addWarnings(output.warnings);
    }
    for (const [seriesIndex, series] of data.series.entries()) {
      for (const [pointIndex, point] of Object.entries(series.points ?? {})) {
        if (point.labelColor === undefined) continue;
        const index = Number(pointIndex);
        const number = series.values[index];
        if (number === null || number === undefined) continue;
        const pointText = pointDataLabelTexts(data, props.dataLabels, index);
        if (pointText.length === 0) continue;
        const output = writeChartTextProperties(
          mergeTextStyle(chartStyle, {
            ...props.dataLabels.style,
            color: point.labelColor,
          }),
          pointText,
          "Chart.dataLabels.point.style",
          theme,
          embedFonts,
        );
        textStyles.pointLabels ??= new Map();
        textStyles.pointLabels.set(`${seriesIndex}:${pointIndex}`, output.xml);
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

export type ParsedChart = {
  readonly testId?: string;
  readonly part?: string;
  readonly name?: string;
  readonly morphKey?: string;
  readonly link?: LinkTarget;
  readonly declaredX?: number;
  readonly declaredY?: number;
  readonly declaredWidth?: number;
  readonly declaredHeight?: number;
  readonly data: ChartData;
  readonly kind: ChartKind;
  readonly legend?: ChartBase["legend"];
  readonly dataLabels?: false | ChartDataLabels;
  readonly plotArea?: ChartProps["plotArea"];
  readonly style?: TextStyle & LayoutStyle;
  readonly stacking?: "stacked" | "percent";
  readonly valueAxis?: ChartAxis;
  readonly categoryAxis?: ChartAxis;
  readonly gapWidth?: number;
  readonly overlap?: number;
  readonly explosion?: number;
  readonly firstSliceAngle?: number;
  readonly holeSize?: number;
  readonly enter?: unknown;
  readonly emphasis?: unknown;
  readonly motion?: unknown;
  readonly contract: CapabilityLayoutContract;
  readonly rawProps: Readonly<Record<string, unknown>>;
};

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
        "Chart does not accept children; chart text and data should be passed in data",
      );
    }

    const kind = input.props["kind"];
    if (typeof kind !== "string" || !Object.hasOwn(chartOptions, kind)) {
      fail("unsupported-chart-kind", `Unsupported chart kind ${String(kind)}`);
    }
    const chartKind = kind as ChartKind;

    for (const option of allChartOptions) {
      if (
        input.props[option] !== undefined &&
        !chartOptions[chartKind].includes(option)
      ) {
        fail(
          "invalid-chart-option",
          `${chartKind} chart does not support ${option}`,
        );
      }
    }

    onlyKeys(
      input.props,
      [
        "testId",
        "part",
        "name",
        "morphKey",
        "link",
        "x",
        "y",
        "width",
        "height",
        "data",
        "legend",
        "dataLabels",
        "plotArea",
        "style",
        "kind",
        ...allChartOptions,
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

    const data = chartData(input.props["data"], chartKind);
    if (
      (chartKind === "pie" || chartKind === "doughnut") &&
      data.series.length !== 1
    ) {
      fail(
        "invalid-data",
        `${chartKind} chart currently only supports a single data series`,
      );
    }

    const testId = optionalName(input.props["testId"], "testId");
    const part = optionalName(input.props["part"], "part");
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

    const stack = input.props["stacking"];
    if (stack !== undefined && stack !== "stacked" && stack !== "percent") {
      fail("invalid-stacking", "Chart.stacking must be stacked or percent");
    }
    const labelsValue = labels(input.props["dataLabels"], chartKind, stack);

    const gapWidth = input.props["gapWidth"];
    if (
      gapWidth !== undefined &&
      (typeof gapWidth !== "number" ||
        !Number.isFinite(gapWidth) ||
        gapWidth < 0 ||
        gapWidth > 500)
    ) {
      fail("invalid-gap-width", "Chart.gapWidth must be between 0 and 500");
    }
    const overlap = input.props["overlap"];
    if (
      overlap !== undefined &&
      (typeof overlap !== "number" ||
        !Number.isFinite(overlap) ||
        overlap < -100 ||
        overlap > 100)
    ) {
      fail("invalid-overlap", "Chart.overlap must be between -100 and 100");
    }

    const checkedPlotArea = plotArea(
      input.props["plotArea"],
      declaredWidth !== undefined && declaredHeight !== undefined
        ? { width: declaredWidth, height: declaredHeight }
        : undefined,
    );

    const explosion =
      chartKind === "pie" || chartKind === "doughnut"
        ? roundedRange(
            input.props["explosion"],
            "Chart.explosion",
            0,
            maximumExplosionForCurrentValidationProfile,
            "rounded-result",
          )
        : undefined;
    const firstSliceAngle =
      chartKind === "pie" || chartKind === "doughnut"
        ? roundedRange(
            input.props["firstSliceAngle"],
            "Chart.firstSliceAngle",
            0,
            360,
          )
        : undefined;
    const holeSize =
      chartKind === "doughnut"
        ? roundedRange(input.props["holeSize"], "Chart.holeSize", 10, 90)
        : undefined;

    const valueAxis =
      input.props["valueAxis"] === undefined
        ? undefined
        : axis(input.props["valueAxis"], "Chart.valueAxis", "value");
    const categoryAxis =
      input.props["categoryAxis"] === undefined
        ? undefined
        : axis(input.props["categoryAxis"], "Chart.categoryAxis", "category");

    const parsed: ParsedChart = {
      kind: chartKind,
      data,
      contract,
      rawProps: input.props,
      ...(testId === undefined ? {} : { testId }),
      ...(part === undefined ? {} : { part }),
      ...(name === undefined ? {} : { name }),
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
      ...(labelsValue === undefined ? {} : { dataLabels: labelsValue }),
      ...(checkedPlotArea === undefined ? {} : { plotArea: checkedPlotArea }),
      ...(input.props["style"] === undefined
        ? {}
        : { style: input.props["style"] as TextStyle & LayoutStyle }),
      ...(stack === undefined ? {} : { stacking: stack }),
      ...(valueAxis === undefined ? {} : { valueAxis }),
      ...(categoryAxis === undefined ? {} : { categoryAxis }),
      ...(gapWidth === undefined ? {} : { gapWidth }),
      ...(overlap === undefined ? {} : { overlap }),
      ...(explosion === undefined ? {} : { explosion }),
      ...(firstSliceAngle === undefined ? {} : { firstSliceAngle }),
      ...(holeSize === undefined ? {} : { holeSize }),
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
    const box = { width, height };

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

    const props: ChartProps = {
      ...parsed.rawProps,
      kind: parsed.kind,
      data: parsed.data,
      x,
      y,
      width,
      height,
      ...(parsed.testId === undefined ? {} : { testId: parsed.testId }),
      ...(parsed.part === undefined ? {} : { part: parsed.part }),
      ...(parsed.name === undefined ? {} : { name: parsed.name }),
      ...(parsed.legend === undefined ? {} : { legend: parsed.legend }),
      ...(parsed.dataLabels === undefined
        ? {}
        : { dataLabels: parsed.dataLabels }),
      ...(parsed.plotArea === undefined ? {} : { plotArea: parsed.plotArea }),
      ...(parsed.style === undefined ? {} : { style: parsed.style }),
      ...(parsed.stacking === undefined ? {} : { stacking: parsed.stacking }),
      ...(parsed.valueAxis === undefined
        ? {}
        : { valueAxis: parsed.valueAxis }),
      ...(parsed.categoryAxis === undefined
        ? {}
        : { categoryAxis: parsed.categoryAxis }),
      ...(parsed.gapWidth === undefined ? {} : { gapWidth: parsed.gapWidth }),
      ...(parsed.overlap === undefined ? {} : { overlap: parsed.overlap }),
      ...(parsed.explosion === undefined
        ? {}
        : { explosion: parsed.explosion }),
      ...(parsed.firstSliceAngle === undefined
        ? {}
        : { firstSliceAngle: parsed.firstSliceAngle }),
      ...(parsed.holeSize === undefined ? {} : { holeSize: parsed.holeSize }),
      ...(parsed.enter === undefined ? {} : { enter: parsed.enter }),
      ...(parsed.emphasis === undefined ? {} : { emphasis: parsed.emphasis }),
      ...(parsed.motion === undefined ? {} : { motion: parsed.motion }),
    } as ChartProps;

    const name =
      context.morphName ??
      props.name ??
      props.testId ??
      `Chart ${context.id - 1}`;
    const annotations = motionAnnotations(parsed.rawProps, {
      kind: "chart",
      name,
      ...(props.testId === undefined ? {} : { testId: props.testId }),
      textBody: false,
      filled: props.kind !== "line",
      box: {
        x: props.x!,
        y: props.y!,
        width: props.width!,
        height: props.height!,
      },
      rotationDegrees: 0,
      hasBackground: props.kind !== "line",
      chartSeriesCount: props.data.series.length,
      chartCategoryCount: props.data.categories.length,
    });
    const { textStyles, resourceClaims, fontUses, warnings } =
      chartResourceClaims(
        props,
        props.data,
        context.theme,
        context.options.embedFonts,
      );
    const workbook = chartWorkbook(props.data);
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
          props,
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
        {
          x: props.x!,
          y: props.y!,
          width: props.width!,
          height: props.height!,
        },
        relationshipId,
        context.nativeLink,
      ),
      observation: {
        id: context.id,
        kind: "chart",
        name,
        ...(props.testId === undefined ? {} : { testId: props.testId }),
        ...(props.part === undefined ? {} : { part: props.part }),
        box: {
          x: props.x!,
          y: props.y!,
          width: props.width!,
          height: props.height!,
        },
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
