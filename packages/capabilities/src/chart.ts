import {
  CapabilityError,
  pxToEmu,
  xmlEscape,
  type CapabilityInput,
  type CapabilityModule,
  type ColorValue,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type GeometryBox,
  type LayoutStyle,
  type LinkTarget,
  type ThemeContext,
} from "@react-pptx/core";
import type { ReactNode } from "react";

import {
  chartFrameXml,
  chartPartXml,
  chartWorkbook,
  type ChartTextStyleXml,
} from "./chart-xml.ts";
import {
  fontPresentationContributions,
  type FontEmbeddingUse,
} from "./font-embedding.ts";
import type { MorphProps } from "./morph-api.ts";
import {
  motionAnnotations,
  motionElement,
  type MotionProps,
} from "./motion-api.ts";
import {
  writeChartTextProperties,
  type TextResourceClaim,
} from "./text-service.ts";
import type { TextStyle } from "./text-service.ts";

type ChartTextStyle = Omit<TextStyle, "textWrap">;

export type ChartKind = "column" | "bar" | "line" | "area" | "pie" | "doughnut";

export type ChartPoint = {
  readonly color?: ColorValue;
  /** 覆盖该数据点的数据标签文字颜色。 */
  readonly labelColor?: ColorValue;
};

export type ChartSeries = {
  readonly name: string;
  /** 每个分类对应一个数值；null 在图表与工作簿中保留为空值。 */
  readonly values: readonly (number | null)[];
  /** 柱、面积或线的颜色。 */
  readonly color?: ColorValue;
  /** 按分类下标覆盖单个点的颜色或数据标签文字颜色。 */
  readonly points?: Readonly<Record<number, ChartPoint>>;
};

export type ChartData = {
  readonly categories: readonly string[];
  readonly series: readonly ChartSeries[];
};

export type ChartLine = {
  readonly color?: ColorValue;
  /** 线宽，单位为 px。 */
  readonly width?: number;
};

export type ChartAxis = {
  /** 是否显示轴及其标签，缺省为 true。 */
  readonly visible?: boolean;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  /** Excel 数字格式，例如 `#,##0`。 */
  readonly numberFormat?: string;
  /** 主网格线；true 使用目标应用默认线型。 */
  readonly gridlines?: boolean | ChartLine;
  /** 坐标轴线；false 隐藏轴线。 */
  readonly line?: boolean | ChartLine;
  /** 轴标签的文字样式。 */
  readonly style?: ChartTextStyle;
};

export type ChartDataLabels = {
  readonly value?: boolean;
  readonly percent?: boolean;
  readonly category?: boolean;
  readonly position?:
    | "center"
    | "insideEnd"
    | "insideBase"
    | "outsideEnd"
    | "bestFit";
  readonly numberFormat?: string;
  readonly style?: ChartTextStyle;
};

type ChartBase = MotionProps &
  MorphProps & {
    readonly testId?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly data: ChartData;
    /** 省略、false 时隐藏图例。 */
    readonly legend?: false | "top" | "bottom" | "left" | "right";
    readonly dataLabels?: false | ChartDataLabels;
    /** 相对图表框的内部绘图区位置，单位为 px。 */
    readonly plotArea?: {
      readonly left: number;
      readonly top: number;
      readonly width: number;
      readonly height: number;
    };
    /** 图表图例、轴和数据标签共用的文字样式。 */
    readonly style?: ChartTextStyle & LayoutStyle;
  };

type CartesianOptions = {
  readonly stacking?: "stacked" | "percent";
  readonly valueAxis?: ChartAxis;
  readonly categoryAxis?: ChartAxis;
};

export type ColumnOrBarChartProps = ChartBase &
  CartesianOptions & {
    readonly kind: "column" | "bar";
    /** 分类组之间的空白占柱宽百分比，取值为 0–500。 */
    readonly gapWidth?: number;
    /** 同组柱重叠百分比，取值为 -100–100。 */
    readonly overlap?: number;
  };

export type LineOrAreaChartProps = ChartBase &
  CartesianOptions & { readonly kind: "line" | "area" };

type PieOptions = {
  /**
   * 将扇区向外拉开的半径百分比。
   * 四舍五入后的整数须在当前正式校验支持的 0–2,147,483,647 范围内。
   */
  readonly explosion?: number;
  /** 第一个扇区的顺时针起始角，范围为 0–360 度，序列化为整数。 */
  readonly firstSliceAngle?: number;
};

export type PieChartProps = ChartBase & PieOptions & { readonly kind: "pie" };

export type DoughnutChartProps = ChartBase &
  PieOptions & {
    readonly kind: "doughnut";
    /** 环形图内孔大小的百分比，范围为 10–90。 */
    readonly holeSize?: number;
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
// DrawingML 的 a:ln/@w 上限为 20,116,800 EMU；先验证换算值再序列化。
const maximumLineWidthEmu = 20_116_800;
// 当前正式校验配置接受的取整值范围为 0–2,147,483,647；此处不是 OOXML 通用格式上限。
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
    fail("unknown-parameter", `${label} 含未实现参数 ${unknown}`);
  }
}

function optionalName(
  value: unknown,
  label: "testId" | "name",
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    fail("invalid-parameter", `${label} 必须是非空字符串`);
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
      `${label} 必须是 -1000000 到 1000000 间的有限像素数值`,
    );
  }
  return value;
}

function extent(value: unknown, label: "width" | "height"): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    fail("invalid-coordinate", `${label} 必须是有限正像素数值`);
  }
  const emu = pxToEmu(value);
  if (emu < 1 || emu > maximumShapeExtentEmu) {
    fail(
      "invalid-coordinate",
      `${label} 换算后必须在 1 到 ${maximumShapeExtentEmu} EMU 的闭区间内`,
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
  ) as TextStyle;
}

function optionalColor(value: unknown, label: string): ColorValue | undefined {
  if (value === undefined) return undefined;
  if (
    typeof value === "string" ||
    (isRecord(value) && ("slot" in value || "color" in value))
  ) {
    return value as ColorValue;
  }
  fail("invalid-color", `${label} 必须是颜色值或主题令牌`);
}

function finiteNumber(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail("invalid-parameter", `${label} 必须是有限数值`);
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
        ? `${label} 四舍五入后必须在 ${minimum} 到 ${maximum} 的范围内`
        : `${label} 必须在 ${minimum} 到 ${maximum} 的格式范围内`,
    );
  }
  return rounded;
}

function xmlText(value: string, label: string): string {
  try {
    xmlEscape(value);
  } catch {
    fail("invalid-xml-character", `${label} 含 XML 1.0 不允许的字符`);
  }
  return value;
}

function lineStyle(value: unknown, label: string): ChartLine | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) fail("invalid-style", `${label} 必须是对象`);
  onlyKeys(value, ["color", "width"], label);
  const color = optionalColor(value["color"], `${label}.color`);
  const width = value["width"];
  if (width !== undefined) {
    if (typeof width !== "number" || !Number.isFinite(width) || width <= 0) {
      fail("invalid-style", `${label}.width 必须是有限正像素数值`);
    }
    if (pxToEmu(width) > maximumLineWidthEmu) {
      fail("invalid-style", `${label}.width 换算后超出可写线宽范围`);
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
  if (!isRecord(value)) fail("invalid-axis", `${label} 必须是对象`);
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
      fail("invalid-axis", `${label}.${unsupported} 仅适用于值轴`);
    }
  }
  const visible = value["visible"];
  if (visible !== undefined && typeof visible !== "boolean") {
    fail("invalid-axis", `${label}.visible 必须是布尔值`);
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
    fail("invalid-axis", `${label}.min 必须小于 max`);
  }
  if (majorUnit !== undefined && majorUnit <= 0) {
    fail("invalid-axis", `${label}.majorUnit 必须大于 0`);
  }
  const numberFormat = value["numberFormat"];
  if (
    numberFormat !== undefined &&
    (typeof numberFormat !== "string" || numberFormat.trim() === "")
  ) {
    fail("invalid-axis", `${label}.numberFormat 必须是非空字符串`);
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
    fail("invalid-axis", `${label}.gridlines 必须是布尔值或线条样式`);
  }
  const line = value["line"];
  if (line !== undefined && typeof line !== "boolean" && !isRecord(line)) {
    fail("invalid-axis", `${label}.line 必须是布尔值或线条样式`);
  }
  return {
    ...(visible === undefined ? {} : { visible }),
    ...(min === undefined ? {} : { min }),
    ...(max === undefined ? {} : { max }),
    ...(majorUnit === undefined ? {} : { majorUnit }),
    ...(numberFormat === undefined
      ? {}
      : { numberFormat: numberFormat as string }),
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
  if (!isRecord(value)) fail("invalid-data", "Chart.data 必须是对象");
  onlyKeys(value, ["categories", "series"], "Chart.data");
  const categories = value["categories"];
  const series = value["series"];
  if (!Array.isArray(categories) || categories.length === 0) {
    fail("invalid-data", "Chart.data.categories 必须包含至少一个分类");
  }
  for (const [index, category] of categories.entries()) {
    if (typeof category !== "string") {
      fail("invalid-data", `Chart.data.categories[${index}] 必须是字符串`);
    }
  }
  if (!Array.isArray(series) || series.length === 0) {
    fail("invalid-data", "Chart.data.series 必须包含至少一个系列");
  }
  const checkedSeries: ChartSeries[] = series.map((item, seriesIndex) => {
    if (!isRecord(item)) {
      fail("invalid-data", `Chart.data.series[${seriesIndex}] 必须是对象`);
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
        `Chart.data.series[${seriesIndex}].name 必须是非空字符串`,
      );
    }
    if (!Array.isArray(values) || values.length !== categories.length) {
      fail(
        "invalid-data",
        `系列「${name}」有 ${Array.isArray(values) ? values.length : 0} 个值，对应 ${categories.length} 个分类`,
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
          `系列「${name}」的第 ${valueIndex + 1} 个值必须是有限数值或 null`,
        );
      }
      if (
        (kind === "pie" || kind === "doughnut") &&
        typeof value === "number" &&
        value < 0
      ) {
        fail(
          "invalid-data",
          `系列「${name}」的第 ${valueIndex + 1} 个饼图值不能为负数`,
        );
      }
    }
    const pointsValue = item["points"];
    let points: Record<number, ChartPoint> | undefined;
    if (pointsValue !== undefined) {
      if (!isRecord(pointsValue)) {
        fail(
          "invalid-data",
          `系列「${name}」的 points 必须是按分类下标索引的对象`,
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
          fail("invalid-data", `系列「${name}」含越界数据点下标 ${rawIndex}`);
        }
        if (!isRecord(pointValue)) {
          fail(
            "invalid-data",
            `系列「${name}」的数据点 ${rawIndex} 必须是对象`,
          );
        }
        onlyKeys(
          pointValue,
          ["color", "labelColor"],
          `系列「${name}」的数据点 ${rawIndex}`,
        );
        const pointColor = optionalColor(
          pointValue["color"],
          `系列「${name}」的数据点 ${rawIndex}.color`,
        );
        const labelColor = optionalColor(
          pointValue["labelColor"],
          `系列「${name}」的数据点 ${rawIndex}.labelColor`,
        );
        points[pointIndex] = {
          ...(pointColor === undefined ? {} : { color: pointColor }),
          ...(labelColor === undefined ? {} : { labelColor }),
        };
      }
    }
    const color = optionalColor(item["color"], `系列「${name}」.color`);
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
): false | ChartDataLabels | undefined {
  if (value === undefined || value === false) return value;
  if (!isRecord(value))
    fail("invalid-data-labels", "Chart.dataLabels 必须是对象或 false");
  onlyKeys(
    value,
    ["value", "percent", "category", "position", "numberFormat", "style"],
    "Chart.dataLabels",
  );
  for (const key of ["value", "percent", "category"] as const) {
    if (value[key] !== undefined && typeof value[key] !== "boolean") {
      fail("invalid-data-labels", `Chart.dataLabels.${key} 必须是布尔值`);
    }
  }
  const position = value["position"];
  if (position !== undefined && !positionValues.has(position as string)) {
    fail(
      "invalid-data-labels",
      typeof position === "string"
        ? `不支持数据标签位置 ${position}`
        : "Chart.dataLabels.position 必须是支持的字符串",
    );
  }
  if (
    position !== undefined &&
    kind !== "column" &&
    kind !== "bar" &&
    kind !== "pie" &&
    kind !== "doughnut"
  ) {
    fail("invalid-data-labels", `${kind} 图表不支持指定数据标签位置`);
  }
  const numberFormat = value["numberFormat"];
  if (
    numberFormat !== undefined &&
    (typeof numberFormat !== "string" || numberFormat.trim() === "")
  ) {
    fail(
      "invalid-data-labels",
      "Chart.dataLabels.numberFormat 必须是非空字符串",
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
    ...(numberFormat === undefined
      ? {}
      : { numberFormat: numberFormat as string }),
    ...(value["style"] === undefined
      ? {}
      : { style: value["style"] as TextStyle }),
  };
}

function plotArea(
  value: unknown,
  box: { readonly width: number; readonly height: number },
): ChartProps["plotArea"] {
  if (value === undefined) return undefined;
  if (!isRecord(value)) fail("invalid-plot-area", "Chart.plotArea 必须是对象");
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
    result.left + result.width > box.width ||
    result.top + result.height > box.height
  ) {
    fail("invalid-plot-area", "Chart.plotArea 必须位于图表框内且宽高为正数");
  }
  return result;
}

const generalNumberCharacters = "0123456789.,-+eE";
const dateNameCharacters =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
// 明确给出简体中文 LCID 时，日期名称使用本地化字形，不按英文字符处理。
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
  // 明确指定但未登记名称字形的 locale 不伪装成英文。
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

/** 申领图表数字格式可能显示的字形，不在构建时模拟应用的数值格式化。 */
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
  );
  const resourceClaims: TextResourceClaim[] = [...chart.resourceClaims];
  const fontUses: FontEmbeddingUse[] = chart.fontUses.map((use) => ({
    ...use,
    path: "Chart.style",
  }));
  const warnings = [...chart.warnings];
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
    );
    textStyles.categoryAxis = output.xml;
    resourceClaims.push(...output.resourceClaims);
    fontUses.push(
      ...output.fontUses.map((use) => ({
        ...use,
        path: "Chart.categoryAxis.style",
      })),
    );
    warnings.push(...output.warnings);
  }
  if (cartesian?.valueAxis?.style !== undefined) {
    const output = writeChartTextProperties(
      mergeTextStyle(chartStyle, cartesian.valueAxis.style),
      valueAxisCharacters,
      "Chart.valueAxis.style",
      theme,
    );
    textStyles.valueAxis = output.xml;
    resourceClaims.push(...output.resourceClaims);
    fontUses.push(
      ...output.fontUses.map((use) => ({
        ...use,
        path: "Chart.valueAxis.style",
      })),
    );
    warnings.push(...output.warnings);
  }
  if (props.dataLabels !== undefined && props.dataLabels !== false) {
    if (props.dataLabels.style !== undefined) {
      const output = writeChartTextProperties(
        mergeTextStyle(chartStyle, props.dataLabels.style),
        dataLabelTexts(data, props.dataLabels),
        "Chart.dataLabels.style",
        theme,
      );
      textStyles.dataLabels = output.xml;
      resourceClaims.push(...output.resourceClaims);
      fontUses.push(
        ...output.fontUses.map((use) => ({
          ...use,
          path: "Chart.dataLabels.style",
        })),
      );
      warnings.push(...output.warnings);
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
        warnings.push(...output.warnings);
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

function validateProps(
  props: Record<string, unknown>,
  layoutBox?: GeometryBox,
): ChartProps {
  const kind = props["kind"];
  if (typeof kind !== "string" || !Object.hasOwn(chartOptions, kind)) {
    fail("unsupported-chart-kind", `不支持图表类型 ${String(kind)}`);
  }
  const chartKind = kind as ChartKind;
  for (const option of allChartOptions) {
    if (
      props[option] !== undefined &&
      !chartOptions[chartKind].includes(option)
    ) {
      fail("invalid-chart-option", `${chartKind} 图表不支持 ${option}`);
    }
  }
  onlyKeys(
    props,
    [
      "testId",
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
      "__motionContext",
    ],
    "Chart",
  );
  const data = chartData(props["data"], chartKind);
  if (
    (chartKind === "pie" || chartKind === "doughnut") &&
    data.series.length !== 1
  ) {
    fail("invalid-data", `${chartKind} 图表当前仅支持一个数据系列`);
  }
  const testId = optionalName(props["testId"], "testId");
  const name = optionalName(props["name"], "name");
  if (name !== undefined) xmlText(name, "Chart.name");
  const style = isRecord(props["style"]) ? props["style"] : {};
  const x = layoutBox?.x ?? coordinate(props["x"], "x");
  const y = layoutBox?.y ?? coordinate(props["y"], "y");
  const width =
    layoutBox?.width ?? extent(props["width"] ?? style["width"], "width");
  const height =
    layoutBox?.height ?? extent(props["height"] ?? style["height"], "height");
  const box = { width, height };
  const legend = props["legend"];
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
      "Chart.legend 必须是 false、top、bottom、left 或 right",
    );
  }
  const labelsValue = labels(props["dataLabels"], chartKind);
  const stack = props["stacking"];
  if (stack !== undefined && stack !== "stacked" && stack !== "percent") {
    fail("invalid-stacking", "Chart.stacking 必须是 stacked 或 percent");
  }
  const gapWidth = props["gapWidth"];
  if (
    gapWidth !== undefined &&
    (typeof gapWidth !== "number" ||
      !Number.isFinite(gapWidth) ||
      gapWidth < 0 ||
      gapWidth > 500)
  ) {
    fail("invalid-gap-width", "Chart.gapWidth 必须在 0 到 500 间");
  }
  const overlap = props["overlap"];
  if (
    overlap !== undefined &&
    (typeof overlap !== "number" ||
      !Number.isFinite(overlap) ||
      overlap < -100 ||
      overlap > 100)
  ) {
    fail("invalid-overlap", "Chart.overlap 必须在 -100 到 100 间");
  }
  const checkedPlotArea = plotArea(props["plotArea"], box);
  const explosion =
    chartKind === "pie" || chartKind === "doughnut"
      ? roundedRange(
          props["explosion"],
          "Chart.explosion",
          0,
          maximumExplosionForCurrentValidationProfile,
          "rounded-result",
        )
      : undefined;
  const firstSliceAngle =
    chartKind === "pie" || chartKind === "doughnut"
      ? roundedRange(props["firstSliceAngle"], "Chart.firstSliceAngle", 0, 360)
      : undefined;
  const holeSize =
    chartKind === "doughnut"
      ? roundedRange(props["holeSize"], "Chart.holeSize", 10, 90)
      : undefined;
  return {
    ...(testId === undefined ? {} : { testId }),
    ...(name === undefined ? {} : { name }),
    ...(props["enter"] === undefined ? {} : { enter: props["enter"] }),
    ...(props["emphasis"] === undefined ? {} : { emphasis: props["emphasis"] }),
    ...(props["motion"] === undefined ? {} : { motion: props["motion"] }),
    ...(props["__motionContext"] === undefined
      ? {}
      : { __motionContext: props["__motionContext"] }),
    x,
    y,
    width,
    height,
    data,
    kind: chartKind,
    ...(legend === undefined ? {} : { legend: legend as ChartBase["legend"] }),
    ...(labelsValue === undefined ? {} : { dataLabels: labelsValue }),
    ...(checkedPlotArea === undefined ? {} : { plotArea: checkedPlotArea }),
    ...(props["style"] === undefined
      ? {}
      : { style: props["style"] as TextStyle & LayoutStyle }),
    ...(stack === undefined ? {} : { stacking: stack }),
    ...(props["valueAxis"] === undefined
      ? {}
      : { valueAxis: axis(props["valueAxis"], "Chart.valueAxis", "value") }),
    ...(props["categoryAxis"] === undefined
      ? {}
      : {
          categoryAxis: axis(
            props["categoryAxis"],
            "Chart.categoryAxis",
            "category",
          ),
        }),
    ...(gapWidth === undefined ? {} : { gapWidth: gapWidth as number }),
    ...(overlap === undefined ? {} : { overlap: overlap as number }),
    ...(chartKind === "pie" || chartKind === "doughnut"
      ? {
          ...(explosion === undefined ? {} : { explosion }),
          ...(firstSliceAngle === undefined ? {} : { firstSliceAngle }),
        }
      : {}),
    ...(chartKind === "doughnut" && holeSize !== undefined ? { holeSize } : {}),
  } as ChartProps;
}

/** 图表模块校验输入，并生成原生图表部件及其嵌入工作簿。 */
export const chartCapability: CapabilityModule = {
  declaration: "rp:chart",
  measureLayout(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement {
    if (input.children.length > 0) {
      fail(
        "unsupported-children",
        "Chart 不接受子内容；图表文字和数据应放入 data",
      );
    }
    const style = isRecord(input.props["style"]) ? input.props["style"] : {};
    const rawWidth = style["width"] ?? input.props["width"];
    const rawHeight = style["height"] ?? input.props["height"];
    const declaredWidth = typeof rawWidth === "number" ? rawWidth : undefined;
    const declaredHeight =
      typeof rawHeight === "number" ? rawHeight : undefined;
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
  compile(input: CapabilityInput, context) {
    if (input.children.length > 0) {
      fail(
        "unsupported-children",
        "Chart 不接受子内容；图表文字和数据应放入 data",
      );
    }
    const props = validateProps({ ...input.props }, context.layoutBox);
    const name =
      context.morphName ??
      props.name ??
      props.testId ??
      `Chart ${context.id - 1}`;
    const annotations = motionAnnotations(props, {
      kind: "chart",
      name,
      ...(props.testId === undefined ? {} : { testId: props.testId }),
      textBody: false,
      filled: props.kind !== "line",
      chartSeriesCount: props.data.series.length,
      chartCategoryCount: props.data.categories.length,
    });
    const { textStyles, resourceClaims, fontUses, warnings } =
      chartResourceClaims(props, props.data, context.theme);
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
          fail("invalid-part", "图表部件需要嵌入工作簿关系");
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

/** 声明一个由官方能力模块编译为原生图表的对象。 */
export function Chart(props: ChartProps): ReactNode {
  return motionElement("rp:chart", props);
}
