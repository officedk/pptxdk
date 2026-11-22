import {
  CapabilityError,
  authorIdSchema,
  colorValueSchema,
  layoutStyleSchema,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  textStyleSchema,
} from "@pptxdk/core/capability";
import { pxToEmu, xmlEscape } from "@pptxdk/emit";
import { z } from "zod";

import { MotionPropsSchema } from "../motion/motion-schemas.ts";
import { MorphKeySchema } from "../shared/morph-api.ts";
import { chartDataLabelPositions } from "./chart-labels.ts";

const maximumLineWidthEmu = 20_116_800;
const maximumShapeExtentEmu = 2_147_483_647;
const maximumExplosionForCurrentValidationProfile = 2_147_483_647;

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isValidXml10String(value: string): boolean {
  try {
    xmlEscape(value);
    return true;
  } catch {
    return false;
  }
}

function addChartIssue(
  ctx: z.RefinementCtx,
  errorCode: string,
  message: string,
  path: readonly (string | number)[] = [],
  fieldPath?: string,
): void {
  ctx.addIssue({
    code: "custom",
    message,
    path: [...path],
    params: {
      errorCode,
      ...(fieldPath === undefined ? {} : { fieldPath }),
    },
  });
}

function nonEmptyStringSchema(errorCode: string, message: string) {
  return z.string().superRefine((value, ctx) => {
    if (value.trim() === "") {
      addChartIssue(ctx, errorCode, message);
    }
  });
}

function nonEmptyXmlStringSchema(errorCode: string, label: string) {
  return z.string().superRefine((value, ctx) => {
    if (value.trim() === "") {
      addChartIssue(ctx, errorCode, `${label} must be a non-empty string`);
      return;
    }
    if (!isValidXml10String(value)) {
      addChartIssue(
        ctx,
        "invalid-xml-character",
        `${label} contains characters not allowed in XML 1.0`,
      );
    }
  });
}

function xmlStringSchema(label: string) {
  return z.string().superRefine((value, ctx) => {
    if (!isValidXml10String(value)) {
      addChartIssue(
        ctx,
        "invalid-xml-character",
        `${label} contains characters not allowed in XML 1.0`,
      );
    }
  });
}

function roundedIntInRangeSchema(options: {
  readonly minimum: number;
  readonly maximum: number;
  readonly boundary: "input-and-result" | "rounded-result";
  readonly errorCode: string;
  readonly label: string;
}) {
  return z
    .number()
    .finite()
    .superRefine((value, ctx) => {
      const rounded = Math.round(value);
      if (
        (options.boundary === "input-and-result" &&
          (value < options.minimum || value > options.maximum)) ||
        rounded < options.minimum ||
        rounded > options.maximum
      ) {
        addChartIssue(
          ctx,
          options.errorCode,
          options.boundary === "rounded-result"
            ? `${options.label} rounded value must be within range of ${options.minimum} to ${options.maximum}`
            : `${options.label} must be within format range of ${options.minimum} to ${options.maximum}`,
        );
      }
    })
    .transform((value) => Math.round(value));
}

export const ChartKindSchema = z.enum([
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "doughnut",
  "scatter",
]);
export type ChartKind = z.infer<typeof ChartKindSchema>;

export const ColorValueSchema = colorValueSchema;

const textStyleBaseSchema = textStyleSchema.unwrap();
export const ChartTextStyleSchema = textStyleBaseSchema
  .omit({ textWrap: true })
  .readonly();
export type ChartTextStyle = z.infer<typeof ChartTextStyleSchema>;

export const ChartStyleSchema = z
  .strictObject({
    ...ChartTextStyleSchema.unwrap().shape,
    ...layoutStyleSchema.unwrap().shape,
  })
  .readonly();

export const ChartTitleSchema = z.union([
  nonEmptyXmlStringSchema("invalid-parameter", "title"),
  z
    .strictObject({
      text: nonEmptyXmlStringSchema("invalid-parameter", "title.text"),
      style: ChartTextStyleSchema.optional(),
    })
    .readonly(),
]);
export type ChartTitle = z.infer<typeof ChartTitleSchema>;

export const ChartLineDashSchema = z.enum([
  "solid",
  "dash",
  "dot",
  "dashDot",
  "longDash",
  "longDashDot",
  "longDashDotDot",
]);
export type ChartLineDash = z.infer<typeof ChartLineDashSchema>;

export const ChartLineSchema = z
  .strictObject({
    color: ColorValueSchema.optional(),
    width: z
      .number()
      .finite()
      .positive()
      .refine((value) => pxToEmu(value) <= maximumLineWidthEmu, {
        message: "width converted value exceeds writable line width range",
      })
      .optional(),
    dash: ChartLineDashSchema.optional(),
  })
  .readonly();
export type ChartLine = z.infer<typeof ChartLineSchema>;

export const ChartMarkerSymbolSchema = z.enum([
  "auto",
  "circle",
  "dash",
  "diamond",
  "dot",
  "none",
  "plus",
  "square",
  "star",
  "triangle",
  "x",
]);
export type ChartMarkerSymbol = z.infer<typeof ChartMarkerSymbolSchema>;

export const ChartMarkerSchema = z
  .strictObject({
    symbol: ChartMarkerSymbolSchema.optional(),
    size: z.number().int().min(2).max(72).optional(),
    fill: ColorValueSchema.optional(),
    line: ChartLineSchema.optional(),
  })
  .readonly();
export type ChartMarker = z.infer<typeof ChartMarkerSchema>;

export const ChartMarkerOptionSchema = z.union([
  z.boolean(),
  ChartMarkerSchema,
]);

export const ChartPointValueSchema = z.number().finite().nullable();
export const ChartSeriesValuesSchema = z
  .array(ChartPointValueSchema)
  .readonly();
export const ChartCategoriesSchema = z
  .array(xmlStringSchema("category"))
  .readonly();

export const ChartDataLabelPositionSchema = z.enum([
  "center",
  "insideEnd",
  "insideBase",
  "outsideEnd",
  "bestFit",
  "left",
  "right",
  "bottom",
  "top",
]);
export type ChartDataLabelPosition = z.infer<
  typeof ChartDataLabelPositionSchema
>;

const separatorStringSchema = z
  .string()
  .min(1)
  .superRefine((value, ctx) => {
    if (!isValidXml10String(value)) {
      addChartIssue(
        ctx,
        "invalid-xml-character",
        "separator contains characters not allowed in XML 1.0",
      );
    }
  });

const dataLabelCommonFields = {
  value: z.boolean().optional(),
  percent: z.boolean().optional(),
  category: z.boolean().optional(),
  seriesName: z.boolean().optional(),
  xValue: z.boolean().optional(),
  separator: separatorStringSchema.optional(),
  numberFormat: nonEmptyXmlStringSchema(
    "invalid-data-labels",
    "numberFormat",
  ).optional(),
  leaderLines: z.union([z.boolean(), ChartLineSchema]).optional(),
  style: ChartTextStyleSchema.optional(),
};

const chartDataLabelsObjectSchema = z.strictObject({
  ...dataLabelCommonFields,
  position: ChartDataLabelPositionSchema.optional(),
});

export const ChartDataLabelsBaseSchema = chartDataLabelsObjectSchema.readonly();

export type ChartDataLabels<
  Position extends ChartDataLabelPosition = ChartDataLabelPosition,
> = Omit<z.infer<typeof ChartDataLabelsBaseSchema>, "position" | "xValue"> & {
  readonly position?: Position | undefined;
};

export const ChartDataLabelOverrideSchema = chartDataLabelsObjectSchema
  .omit({
    style: true,
    leaderLines: true,
    xValue: true,
  })
  .readonly();
export type ChartDataLabelOverride = z.infer<
  typeof ChartDataLabelOverrideSchema
>;

export const ChartSeriesDataLabelOverrideSchema = chartDataLabelsObjectSchema
  .omit({
    style: true,
    xValue: true,
  })
  .readonly();
export type ChartSeriesDataLabelOverride = z.infer<
  typeof ChartSeriesDataLabelOverrideSchema
>;

export const ChartDataLabelSettingsSchema =
  chartDataLabelsObjectSchema.readonly();
export type ChartDataLabelSettings = z.infer<
  typeof ChartDataLabelSettingsSchema
>;

const scatterDataLabelsObjectSchema = z.strictObject({
  ...dataLabelCommonFields,
  position: z.enum(["left", "right", "bottom", "top", "center"]).optional(),
});

export const ScatterDataLabelsBaseSchema = scatterDataLabelsObjectSchema
  .omit({ percent: true, category: true, leaderLines: true })
  .readonly();
export type ScatterDataLabels = z.infer<typeof ScatterDataLabelsBaseSchema>;

export const ScatterDataLabelOverrideSchema = scatterDataLabelsObjectSchema
  .omit({
    style: true,
    percent: true,
    category: true,
    leaderLines: true,
  })
  .readonly();
export type ScatterDataLabelOverride = z.infer<
  typeof ScatterDataLabelOverrideSchema
>;

const ChartPointRawSchema = z
  .strictObject({
    color: ColorValueSchema.optional(),
    line: ChartLineSchema.optional(),
    marker: z.unknown().optional(),
    explosion: roundedIntInRangeSchema({
      minimum: 0,
      maximum: maximumExplosionForCurrentValidationProfile,
      boundary: "rounded-result",
      errorCode: "invalid-chart-option",
      label: "explosion",
    }).optional(),
    labelColor: ColorValueSchema.optional(),
    dataLabels: z
      .union([z.literal(false), ChartDataLabelOverrideSchema])
      .optional(),
  })
  .readonly();

export type ChartPoint = Omit<
  z.infer<typeof ChartPointRawSchema>,
  "dataLabels" | "marker"
> & {
  readonly marker?: boolean | ChartMarker | undefined;
  readonly dataLabels?: false | ChartDataLabelOverride | undefined;
};
export const ChartPointSchema =
  ChartPointRawSchema as unknown as z.ZodType<ChartPoint>;

export const ScatterChartPointSchema = z
  .strictObject({
    color: ColorValueSchema.optional(),
    line: ChartLineSchema.optional(),
    marker: ChartMarkerOptionSchema.optional(),
    labelColor: ColorValueSchema.optional(),
    dataLabels: z
      .union([z.literal(false), ScatterDataLabelOverrideSchema])
      .optional(),
  })
  .readonly();
export type ScatterChartPoint = z.infer<typeof ScatterChartPointSchema>;

export const ChartTrendlineTypeSchema = z.enum([
  "linear",
  "exponential",
  "logarithmic",
  "polynomial",
  "power",
  "movingAverage",
]);
export type ChartTrendlineType = z.infer<typeof ChartTrendlineTypeSchema>;

const trendlineBaseFields = {
  line: ChartLineSchema.optional(),
  forward: z.number().finite().positive().optional(),
  backward: z.number().finite().positive().optional(),
};

const trendlineDisplayFields = {
  displayEquation: z.boolean().optional(),
  displayRSquared: z.boolean().optional(),
};

export const ChartTrendlineSchema = z.discriminatedUnion("type", [
  z
    .strictObject({
      ...trendlineBaseFields,
      ...trendlineDisplayFields,
      type: z.literal("linear"),
      intercept: z.number().finite().optional(),
      order: z.never().optional(),
      period: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...trendlineBaseFields,
      ...trendlineDisplayFields,
      type: z.literal("exponential"),
      intercept: z.number().finite().optional(),
      order: z.never().optional(),
      period: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...trendlineBaseFields,
      ...trendlineDisplayFields,
      type: z.literal("logarithmic"),
      intercept: z.never().optional(),
      order: z.never().optional(),
      period: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...trendlineBaseFields,
      ...trendlineDisplayFields,
      type: z.literal("polynomial"),
      order: z.number().int().min(2).max(6),
      intercept: z.number().finite().optional(),
      period: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...trendlineBaseFields,
      ...trendlineDisplayFields,
      type: z.literal("power"),
      intercept: z.never().optional(),
      order: z.never().optional(),
      period: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...trendlineBaseFields,
      type: z.literal("movingAverage"),
      period: z.number().int().min(2).max(255),
      intercept: z.never().optional(),
      order: z.never().optional(),
      displayEquation: z.never().optional(),
      displayRSquared: z.never().optional(),
    })
    .readonly(),
]);
export type ChartTrendline = z.infer<typeof ChartTrendlineSchema>;

export const ChartErrorBarValuesSchema = z
  .array(z.number().finite().nonnegative().nullable())
  .readonly();
export type ChartErrorBarValues = z.infer<typeof ChartErrorBarValuesSchema>;

export const ChartErrorBarIncludeSchema = z.enum(["both", "plus", "minus"]);
export type ChartErrorBarInclude = z.infer<typeof ChartErrorBarIncludeSchema>;

const errorBarStyleFields = {
  endCap: z.boolean().optional(),
  line: ChartLineSchema.optional(),
};

export const ChartCustomErrorBarsSchema = z.union([
  z
    .strictObject({
      ...errorBarStyleFields,
      type: z.literal("custom"),
      include: z.literal("both").optional(),
      plus: ChartErrorBarValuesSchema,
      minus: ChartErrorBarValuesSchema,
      value: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...errorBarStyleFields,
      type: z.literal("custom"),
      include: z.literal("plus"),
      plus: ChartErrorBarValuesSchema,
      minus: z.never().optional(),
      value: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...errorBarStyleFields,
      type: z.literal("custom"),
      include: z.literal("minus"),
      minus: ChartErrorBarValuesSchema,
      plus: z.never().optional(),
      value: z.never().optional(),
    })
    .readonly(),
]);
export type ChartCustomErrorBars = z.infer<typeof ChartCustomErrorBarsSchema>;

export const ChartErrorBarsSchema = z.union([
  z
    .strictObject({
      ...errorBarStyleFields,
      include: ChartErrorBarIncludeSchema.optional(),
      type: z.enum(["fixed", "percentage", "standardDeviation"]),
      value: z.number().finite().nonnegative(),
      plus: z.never().optional(),
      minus: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...errorBarStyleFields,
      include: ChartErrorBarIncludeSchema.optional(),
      type: z.literal("standardError"),
      value: z.never().optional(),
      plus: z.never().optional(),
      minus: z.never().optional(),
    })
    .readonly(),
  ChartCustomErrorBarsSchema,
]);
export type ChartErrorBars = z.infer<typeof ChartErrorBarsSchema>;

export const ScatterErrorBarsSchema = z
  .strictObject({
    x: ChartErrorBarsSchema.optional(),
    y: ChartErrorBarsSchema.optional(),
  })
  .readonly()
  .superRefine((value, ctx) => {
    if (value.x === undefined && value.y === undefined) {
      addChartIssue(
        ctx,
        "invalid-error-bars",
        "errorBars must declare x, y, or both",
      );
    }
  });
export type ScatterErrorBars = z.infer<typeof ScatterErrorBarsSchema>;

const chartSeriesFormatFields = {
  name: nonEmptyStringSchema(
    "invalid-data",
    "Series.name must be a non-empty string",
  ),
  color: ColorValueSchema.optional(),
  line: ChartLineSchema.optional(),
  explosion: roundedIntInRangeSchema({
    minimum: 0,
    maximum: maximumExplosionForCurrentValidationProfile,
    boundary: "rounded-result",
    errorCode: "invalid-chart-option",
    label: "explosion",
  }).optional(),
  dataLabels: z
    .union([z.literal(false), chartDataLabelsObjectSchema.readonly()])
    .optional(),
  marker: z.unknown().optional(),
  points: z.record(z.string(), ChartPointSchema).readonly().optional(),
  trendlines: z.array(ChartTrendlineSchema).readonly().optional(),
};

const ChartSeriesRawSchema = z
  .strictObject({
    ...chartSeriesFormatFields,
    values: ChartSeriesValuesSchema,
    smooth: z.boolean().optional(),
    errorBars: ChartErrorBarsSchema.optional(),
  })
  .readonly();

export type ChartSeries = Omit<
  z.infer<typeof ChartSeriesRawSchema>,
  "points" | "dataLabels" | "marker"
> & {
  readonly marker?: boolean | ChartMarker | undefined;
  readonly dataLabels?: false | ChartSeriesDataLabelOverride | undefined;
  readonly points?: Readonly<Record<number, ChartPoint>> | undefined;
};

export const ChartSeriesSchema =
  ChartSeriesRawSchema as unknown as z.ZodType<ChartSeries>;

const IndependentCategorySeriesRawSchema = z
  .strictObject({
    ...chartSeriesFormatFields,
    categories: ChartCategoriesSchema,
    values: ChartSeriesValuesSchema,
    smooth: z.boolean().optional(),
    errorBars: ChartErrorBarsSchema.optional(),
  })
  .readonly();

export type IndependentCategorySeries = Omit<
  z.infer<typeof IndependentCategorySeriesRawSchema>,
  "points" | "dataLabels" | "marker"
> & {
  readonly marker?: boolean | ChartMarker | undefined;
  readonly dataLabels?: false | ChartSeriesDataLabelOverride | undefined;
  readonly points?: Readonly<Record<number, ChartPoint>> | undefined;
};

export const IndependentCategorySeriesSchema =
  IndependentCategorySeriesRawSchema as unknown as z.ZodType<IndependentCategorySeries>;

const ScatterChartSeriesRawSchema = z
  .strictObject({
    name: nonEmptyStringSchema(
      "invalid-data",
      "Series.name must be a non-empty string",
    ),
    x: ChartSeriesValuesSchema,
    y: ChartSeriesValuesSchema,
    color: ColorValueSchema.optional(),
    line: ChartLineSchema.optional(),
    dataLabels: z
      .union([z.literal(false), ScatterDataLabelOverrideSchema])
      .optional(),
    marker: ChartMarkerOptionSchema.optional(),
    points: z.record(z.string(), ScatterChartPointSchema).readonly().optional(),
    trendlines: z.array(ChartTrendlineSchema).readonly().optional(),
    errorBars: ScatterErrorBarsSchema.optional(),
  })
  .readonly();

export type ScatterChartSeries = Omit<
  z.infer<typeof ScatterChartSeriesRawSchema>,
  "points"
> & {
  readonly points?: Readonly<Record<number, ScatterChartPoint>> | undefined;
};

export const ScatterChartSeriesSchema =
  ScatterChartSeriesRawSchema as unknown as z.ZodType<ScatterChartSeries>;

export const SharedCategoryDataSchema = z
  .strictObject({
    kind: z.literal("shared-categories"),
    categories: ChartCategoriesSchema,
  })
  .readonly();
export type SharedCategoryData = z.infer<typeof SharedCategoryDataSchema>;
export type ChartGroupData = SharedCategoryData;

export const IndependentCategoryDataSchema = z
  .strictObject({
    kind: z.literal("independent-categories"),
  })
  .readonly();
export type IndependentCategoryData = z.infer<
  typeof IndependentCategoryDataSchema
>;

export const ScatterDataSchema = z
  .strictObject({
    kind: z.literal("xy"),
  })
  .readonly();
export type ScatterData = z.infer<typeof ScatterDataSchema>;

export const ChartTickMarkSchema = z.enum(["none", "in", "out", "cross"]);
export type ChartTickMark = z.infer<typeof ChartTickMarkSchema>;

export const ChartTickLabelPositionSchema = z.enum([
  "nextTo",
  "high",
  "low",
  "none",
]);
export type ChartTickLabelPosition = z.infer<
  typeof ChartTickLabelPositionSchema
>;

export const ChartAxisCrossingSchema = z.enum(["autoZero", "max", "min"]);
export type ChartAxisCrossing = z.infer<typeof ChartAxisCrossingSchema>;

export const ChartAxisPositionSchema = z.enum([
  "bottom",
  "top",
  "left",
  "right",
]);
export type ChartAxisPosition = z.infer<typeof ChartAxisPositionSchema>;

export const ChartDisplayBlanksAsSchema = z.enum(["gap", "zero", "span"]);
export type ChartDisplayBlanksAs = z.infer<typeof ChartDisplayBlanksAsSchema>;

const axisBaseFields = {
  visible: z.boolean().optional(),
  gridlines: z.union([z.boolean(), ChartLineSchema]).optional(),
  minorGridlines: z.union([z.boolean(), ChartLineSchema]).optional(),
  line: z.union([z.boolean(), ChartLineSchema]).optional(),
  style: ChartTextStyleSchema.optional(),
  title: ChartTitleSchema.optional(),
  majorTickMark: ChartTickMarkSchema.optional(),
  minorTickMark: ChartTickMarkSchema.optional(),
  tickLabelPosition: ChartTickLabelPositionSchema.optional(),
  reversed: z.boolean().optional(),
  position: ChartAxisPositionSchema.optional(),
  crosses: ChartAxisCrossingSchema.optional(),
  crossesAt: z.number().finite().optional(),
};

export const ChartCategoryAxisSchema = z
  .strictObject({
    kind: z.literal("category"),
    ...axisBaseFields,
  })
  .readonly()
  .superRefine((value, ctx) => {
    if (value.crosses !== undefined && value.crossesAt !== undefined) {
      addChartIssue(
        ctx,
        "invalid-axis",
        "axis must declare crosses or crossesAt, not both",
      );
    }
    if (
      value.crossesAt !== undefined &&
      (!Number.isSafeInteger(value.crossesAt) || value.crossesAt < 1)
    ) {
      addChartIssue(
        ctx,
        "invalid-axis",
        "crossesAt must be a positive integer category position",
        ["crossesAt"],
      );
    }
  });
export type ChartCategoryAxis = z.infer<typeof ChartCategoryAxisSchema>;

export const ChartValueAxisSchema = z
  .strictObject({
    kind: z.literal("value"),
    ...axisBaseFields,
    min: z.number().finite().optional(),
    max: z.number().finite().optional(),
    majorUnit: z.number().finite().positive().optional(),
    minorUnit: z.number().finite().positive().optional(),
    numberFormat: nonEmptyXmlStringSchema(
      "invalid-axis",
      "numberFormat",
    ).optional(),
    logBase: z.number().finite().min(2).max(1000).optional(),
  })
  .readonly()
  .superRefine((value, ctx) => {
    if (
      value.min !== undefined &&
      value.max !== undefined &&
      value.min >= value.max
    ) {
      addChartIssue(ctx, "invalid-axis", "min must be less than max", ["min"]);
    }
    if (value.logBase !== undefined) {
      if (value.min !== undefined && value.min <= 0) {
        addChartIssue(
          ctx,
          "invalid-axis",
          "min must be greater than 0 for a logarithmic axis",
          ["min"],
        );
      }
      if (value.max !== undefined && value.max <= 0) {
        addChartIssue(
          ctx,
          "invalid-axis",
          "max must be greater than 0 for a logarithmic axis",
          ["max"],
        );
      }
      if (value.crossesAt !== undefined && value.crossesAt <= 0) {
        addChartIssue(
          ctx,
          "invalid-axis",
          "crossesAt must be greater than 0 for a logarithmic axis",
          ["crossesAt"],
        );
      }
    }
    if (value.crosses !== undefined && value.crossesAt !== undefined) {
      addChartIssue(
        ctx,
        "invalid-axis",
        "axis must declare crosses or crossesAt, not both",
      );
    }
  });
export type ChartValueAxis = z.infer<typeof ChartValueAxisSchema>;

export const ChartAxisSchema = z.discriminatedUnion("kind", [
  ChartCategoryAxisSchema,
  ChartValueAxisSchema,
]);
export type ChartAxis = z.infer<typeof ChartAxisSchema>;

export const ChartAxesSchema = z
  .record(z.string(), ChartAxisSchema)
  .superRefine((axes, ctx) => {
    const names = Object.keys(axes);
    if (names.length === 0) {
      addChartIssue(
        ctx,
        "invalid-axis",
        "Chart.axes must define at least one axis",
      );
    }
    for (const name of names) {
      if (name.trim() === "") {
        addChartIssue(
          ctx,
          "invalid-axis",
          "Chart.axes names must be non-empty strings",
          [name],
        );
      }
    }
  });
export type ChartAxes = Readonly<Record<string, ChartAxis>>;

export const ChartCategoryAxisReferencesSchema = z
  .strictObject({
    category: nonEmptyStringSchema(
      "invalid-chart-axes",
      "category axis reference must be non-empty",
    ),
    value: nonEmptyStringSchema(
      "invalid-chart-axes",
      "value axis reference must be non-empty",
    ),
  })
  .readonly();
export type ChartCategoryAxisReferences = z.infer<
  typeof ChartCategoryAxisReferencesSchema
>;

export const ChartXYAxisReferencesSchema = z
  .strictObject({
    x: nonEmptyStringSchema(
      "invalid-chart-axes",
      "X axis reference must be non-empty",
    ),
    y: nonEmptyStringSchema(
      "invalid-chart-axes",
      "Y axis reference must be non-empty",
    ),
  })
  .readonly()
  .superRefine((value, ctx) => {
    if (value.x === value.y) {
      addChartIssue(
        ctx,
        "invalid-chart-axes",
        "axes must reference two distinct X and Y value axes",
      );
    }
  });
export type ChartXYAxisReferences = z.infer<typeof ChartXYAxisReferencesSchema>;

export const ChartAxisReferencesSchema = z.union([
  ChartCategoryAxisReferencesSchema,
  ChartXYAxisReferencesSchema,
]);
export type ChartAxisReferences = z.infer<typeof ChartAxisReferencesSchema>;

export const ScatterLineSchema = z.union([
  z.literal(false),
  z.literal("straight"),
  z.literal("smooth"),
]);
export type ScatterLine = z.infer<typeof ScatterLineSchema>;

function validatePointIndexKeys(
  points: Readonly<Record<string, unknown>> | undefined,
  pointCount: number,
  seriesName: string,
  ctx: z.RefinementCtx,
  pathPrefix: readonly (string | number)[],
): void {
  if (points === undefined) return;
  for (const rawIndex of Object.keys(points)) {
    const pointIndex = Number(rawIndex);
    if (
      !/^\d+$/u.test(rawIndex) ||
      !Number.isSafeInteger(pointIndex) ||
      pointIndex < 0 ||
      pointIndex >= pointCount
    ) {
      addChartIssue(
        ctx,
        "invalid-data",
        `Series "${seriesName}" contains out-of-bounds data point index ${rawIndex}`,
        [...pathPrefix, "points", rawIndex],
      );
    }
  }
}

function validateDataLabelSettingsForKind(
  labels: false | ChartDataLabelSettings | undefined,
  kind: ChartKind,
  stacking: "stacked" | "percent" | undefined,
  ctx: z.RefinementCtx,
  path: readonly (string | number)[],
  allowStyle: boolean,
  allowLeaderLines: boolean,
): void {
  if (labels === undefined || labels === false) return;
  if (!allowStyle && labels.style !== undefined) {
    addChartIssue(
      ctx,
      "unknown-parameter",
      "dataLabels contains unimplemented parameter style",
      [...path, "style"],
    );
  }
  if (!allowLeaderLines && labels.leaderLines !== undefined) {
    addChartIssue(
      ctx,
      "unknown-parameter",
      "dataLabels contains unimplemented parameter leaderLines",
      [...path, "leaderLines"],
    );
  }
  if (labels.xValue !== undefined && kind !== "scatter") {
    addChartIssue(
      ctx,
      "invalid-data-labels",
      "xValue is only supported for scatter charts",
      [...path, "xValue"],
    );
  }
  if (labels.percent === true && kind !== "pie" && kind !== "doughnut") {
    addChartIssue(
      ctx,
      "invalid-data-labels",
      "percent is only supported for pie and doughnut charts",
      [...path, "percent"],
    );
  }
  if (kind === "scatter" && labels.category === true) {
    addChartIssue(
      ctx,
      "invalid-data-labels",
      "category is not applicable to scatter charts; X values are not category labels",
      [...path, "category"],
    );
  }
  if (
    labels.leaderLines !== undefined &&
    kind !== "pie" &&
    kind !== "doughnut"
  ) {
    addChartIssue(
      ctx,
      "invalid-data-labels",
      "leaderLines is only supported for pie and doughnut charts",
      [...path, "leaderLines"],
    );
  }
  const position = labels.position;
  if (position === undefined) return;
  const mapped = chartDataLabelPositions[position];
  if (kind === "area") {
    addChartIssue(
      ctx,
      "invalid-data-labels",
      "Office does not allow dLblPos for areaChart series data labels",
      [...path, "position"],
    );
  } else if (kind === "doughnut") {
    addChartIssue(
      ctx,
      "invalid-data-labels",
      "Office does not allow dLblPos for doughnutChart series data labels",
      [...path, "position"],
    );
  } else if (kind === "scatter" || kind === "line") {
    if (!["left", "right", "bottom", "top", "center"].includes(position)) {
      addChartIssue(
        ctx,
        "invalid-data-labels",
        `Office does not allow dLblPos="${mapped}" for ${kind}Chart series data labels`,
        [...path, "position"],
      );
    }
  } else if (kind === "pie") {
    if (position === "insideBase") {
      addChartIssue(
        ctx,
        "invalid-data-labels",
        'Office does not allow dLblPos="inBase" for pieChart series data labels',
        [...path, "position"],
      );
    } else if (
      !["center", "insideEnd", "outsideEnd", "bestFit"].includes(position)
    ) {
      addChartIssue(
        ctx,
        "invalid-data-labels",
        `Office does not allow dLblPos="${mapped}" for pieChart series data labels`,
        [...path, "position"],
      );
    }
  } else {
    if (position === "bestFit") {
      addChartIssue(
        ctx,
        "invalid-data-labels",
        'Office does not allow dLblPos="bestFit" for barChart series data labels',
        [...path, "position"],
      );
    } else if (stacking !== undefined && position === "outsideEnd") {
      addChartIssue(
        ctx,
        "invalid-data-labels",
        'Office does not allow dLblPos="outEnd" for stacked barChart series data labels',
        [...path, "position"],
      );
    } else {
      const allowed =
        stacking === undefined
          ? ["center", "insideEnd", "insideBase", "outsideEnd"]
          : ["center", "insideEnd", "insideBase"];
      if (!allowed.includes(position)) {
        addChartIssue(
          ctx,
          "invalid-data-labels",
          `Office does not allow dLblPos="${mapped}" for barChart series data labels`,
          [...path, "position"],
        );
      }
    }
  }
}

function validateErrorBarsForKind(
  bars: ChartErrorBars | undefined,
  kind: ChartKind,
  pointCount: number,
  ctx: z.RefinementCtx,
  path: readonly (string | number)[],
): void {
  if (bars === undefined) return;
  if (kind === "pie" || kind === "doughnut") {
    addChartIssue(
      ctx,
      "invalid-error-bars",
      `${kind} charts do not support native error bars`,
      path,
    );
    return;
  }
  if (bars.type === "custom") {
    if (bars.plus !== undefined && bars.plus.length !== pointCount) {
      addChartIssue(
        ctx,
        "invalid-error-bars",
        `plus must contain one magnitude per original series slot (${pointCount})`,
        [...path, "plus"],
      );
    }
    if (bars.minus !== undefined && bars.minus.length !== pointCount) {
      addChartIssue(
        ctx,
        "invalid-error-bars",
        `minus must contain one magnitude per original series slot (${pointCount})`,
        [...path, "minus"],
      );
    }
  }
}

function validateTrendlinesForGroup(
  trendlines: readonly ChartTrendline[] | undefined,
  kind: ChartKind,
  stacking: "stacked" | "percent" | undefined,
  pointCount: number,
  fitPoints: readonly { readonly x: number; readonly y: number }[],
  ctx: z.RefinementCtx,
  path: readonly (string | number)[],
): void {
  if (trendlines === undefined || trendlines.length === 0) return;
  if (kind === "pie" || kind === "doughnut") {
    addChartIssue(
      ctx,
      "invalid-trendline",
      `${kind} charts do not support native trendlines`,
      path,
    );
    return;
  }
  if (stacking !== undefined) {
    addChartIssue(
      ctx,
      "invalid-trendline",
      `${kind} trendlines are not supported on stacked or percent-stacked groups`,
      path,
    );
    return;
  }
  for (const [index, item] of trendlines.entries()) {
    const itemPath = [...path, index];
    if (kind !== "scatter") {
      for (const side of ["forward", "backward"] as const) {
        const forecast = item[side];
        if (
          forecast !== undefined &&
          Math.abs(forecast * 2 - Math.round(forecast * 2)) > 1e-9
        ) {
          addChartIssue(
            ctx,
            "invalid-trendline",
            `${side} must be a multiple of 0.5 for category charts`,
            [...itemPath, side],
          );
        }
      }
    }
    if (item.type === "movingAverage" && item.period >= pointCount) {
      addChartIssue(
        ctx,
        "invalid-trendline",
        `period must be less than the series point count ${pointCount}`,
        [...itemPath, "period"],
      );
    }
    if (item.type === "logarithmic" && fitPoints.some(({ x }) => x <= 0)) {
      addChartIssue(
        ctx,
        "invalid-trendline-domain",
        "type logarithmic requires positive X values (category positions or scatter X)",
        [...itemPath, "type"],
      );
    }
    if (item.type === "exponential" && fitPoints.some(({ y }) => y <= 0)) {
      addChartIssue(
        ctx,
        "invalid-trendline-domain",
        "type exponential requires positive Y values",
        [...itemPath, "type"],
      );
    }
    if (
      item.type === "power" &&
      fitPoints.some(({ x, y }) => x <= 0 || y <= 0)
    ) {
      addChartIssue(
        ctx,
        "invalid-trendline-domain",
        "type power requires positive X and Y values",
        [...itemPath, "type"],
      );
    }
  }
}

function validateCategoryGroupSeries(
  group: {
    readonly kind: Exclude<ChartKind, "scatter">;
    readonly data: SharedCategoryData | IndependentCategoryData;
    readonly series: readonly (ChartSeries | IndependentCategorySeries)[];
    readonly stacking?: "stacked" | "percent" | undefined;
    readonly dataLabels?: false | ChartDataLabelSettings | undefined;
  },
  ctx: z.RefinementCtx,
): void {
  const { kind, data, series, stacking } = group;
  validateDataLabelSettingsForKind(
    group.dataLabels,
    kind,
    stacking,
    ctx,
    ["dataLabels"],
    true,
    true,
  );
  if (kind === "pie" && series.length !== 1) {
    addChartIssue(
      ctx,
      "invalid-data",
      "pie chart currently only supports a single data series",
      ["series"],
    );
  }
  for (const [seriesIndex, item] of series.entries()) {
    const seriesPath = ["series", seriesIndex] as const;
    const categories =
      data.kind === "shared-categories"
        ? data.categories
        : (item as IndependentCategorySeries).categories;
    if (item.values.length !== categories.length) {
      addChartIssue(
        ctx,
        "invalid-data",
        `Series "${item.name}" has ${item.values.length} values for ${categories.length} categories`,
        [...seriesPath, "values"],
      );
    }
    if (kind === "pie" || kind === "doughnut") {
      for (const [valueIndex, value] of item.values.entries()) {
        if (value !== null && value < 0) {
          addChartIssue(
            ctx,
            "invalid-data",
            `Series "${item.name}" pie value ${valueIndex + 1} cannot be negative`,
            [...seriesPath, "values", valueIndex],
          );
        }
      }
    }
    if (item.explosion !== undefined && kind !== "pie" && kind !== "doughnut") {
      addChartIssue(
        ctx,
        "invalid-chart-option",
        `Series "${item.name}" explosion is only supported for pie and doughnut charts`,
        [...seriesPath, "explosion"],
      );
    }
    if (item.marker !== undefined) {
      if (kind !== "line") {
        addChartIssue(
          ctx,
          "invalid-marker",
          `Series "${item.name}".marker is only applicable to line and scatter chart series and points`,
          [...seriesPath, "marker"],
        );
      } else {
        const parsedMarker = ChartMarkerOptionSchema.safeParse(item.marker);
        if (!parsedMarker.success) {
          for (const issue of parsedMarker.error.issues) {
            ctx.addIssue({
              ...issue,
              path: [...seriesPath, "marker", ...issue.path],
            });
          }
        }
      }
    }
    if (item.smooth !== undefined && kind !== "line") {
      addChartIssue(
        ctx,
        "invalid-smooth",
        `Series "${item.name}".smooth is only applicable to line chart series`,
        [...seriesPath, "smooth"],
      );
    }
    validateDataLabelSettingsForKind(
      item.dataLabels,
      kind,
      stacking,
      ctx,
      [...seriesPath, "dataLabels"],
      false,
      true,
    );
    validatePointIndexKeys(
      item.points,
      categories.length,
      item.name,
      ctx,
      seriesPath,
    );
    if (item.points !== undefined) {
      for (const [rawIndex, point] of Object.entries(item.points)) {
        if (kind === "area" && point.color !== undefined) {
          addChartIssue(
            ctx,
            "invalid-data",
            "Area chart data points do not support individual fill colors; set the series color instead.",
            [...seriesPath, "points", rawIndex, "color"],
            `series[${seriesIndex}].points[${rawIndex}].color`,
          );
        }
        if (point.marker !== undefined) {
          if (kind !== "line") {
            addChartIssue(
              ctx,
              "invalid-marker",
              `Series "${item.name}" data point ${rawIndex}.marker is only applicable to line and scatter chart series and points`,
              [...seriesPath, "points", rawIndex, "marker"],
            );
          } else {
            const parsedMarker = ChartMarkerOptionSchema.safeParse(
              point.marker,
            );
            if (!parsedMarker.success) {
              for (const issue of parsedMarker.error.issues) {
                ctx.addIssue({
                  ...issue,
                  path: [
                    ...seriesPath,
                    "points",
                    rawIndex,
                    "marker",
                    ...issue.path,
                  ],
                });
              }
            }
          }
        }
        if (
          point.explosion !== undefined &&
          kind !== "pie" &&
          kind !== "doughnut"
        ) {
          addChartIssue(
            ctx,
            "invalid-chart-option",
            `Series "${item.name}" data point ${rawIndex} explosion is only supported for pie and doughnut charts`,
            [...seriesPath, "points", rawIndex, "explosion"],
          );
        }
        validateDataLabelSettingsForKind(
          point.dataLabels,
          kind,
          stacking,
          ctx,
          [...seriesPath, "points", rawIndex, "dataLabels"],
          false,
          false,
        );
      }
    }
    validateErrorBarsForKind(item.errorBars, kind, categories.length, ctx, [
      ...seriesPath,
      "errorBars",
    ]);
    validateTrendlinesForGroup(
      item.trendlines,
      kind,
      stacking,
      categories.length,
      item.values.flatMap((value, pointIndex) =>
        value === null ? [] : [{ x: pointIndex + 1, y: value }],
      ),
      ctx,
      [...seriesPath, "trendlines"],
    );
  }
}

const columnOrBarCommonFields = {
  kind: z.enum(["column", "bar"]),
  data: SharedCategoryDataSchema,
  series: z.array(ChartSeriesSchema).min(1).readonly(),
  axes: ChartCategoryAxisReferencesSchema.optional(),
  gapWidth: z.number().finite().min(0).max(500).optional(),
  overlap: z.number().finite().min(-100).max(100).optional(),
};

const ColumnOrBarUnstackedGroupSchema = z
  .strictObject({
    ...columnOrBarCommonFields,
    stacking: z.undefined().optional(),
    dataLabels: z
      .union([z.literal(false), ChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly();

const ColumnOrBarStackedGroupSchema = z
  .strictObject({
    ...columnOrBarCommonFields,
    stacking: z.enum(["stacked", "percent"]),
    dataLabels: z
      .union([z.literal(false), ChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly();

export type ColumnOrBarChartGroup =
  | (Omit<z.infer<typeof ColumnOrBarUnstackedGroupSchema>, "dataLabels"> & {
      readonly dataLabels?:
        | false
        | ChartDataLabels<"center" | "insideEnd" | "insideBase" | "outsideEnd">
        | undefined;
    })
  | (Omit<z.infer<typeof ColumnOrBarStackedGroupSchema>, "dataLabels"> & {
      readonly dataLabels?:
        | false
        | ChartDataLabels<"center" | "insideEnd" | "insideBase">
        | undefined;
    });

export const ColumnOrBarChartGroupSchema = z
  .union([ColumnOrBarUnstackedGroupSchema, ColumnOrBarStackedGroupSchema])
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  }) as unknown as z.ZodType<ColumnOrBarChartGroup>;

const LineChartGroupRawSchema = z
  .strictObject({
    kind: z.literal("line"),
    data: SharedCategoryDataSchema,
    series: z.array(ChartSeriesSchema).min(1).readonly(),
    axes: ChartCategoryAxisReferencesSchema.optional(),
    stacking: z.enum(["stacked", "percent"]).optional(),
    dataLabels: z
      .union([z.literal(false), ChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly()
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  });

export type LineChartGroup = Omit<
  z.infer<typeof LineChartGroupRawSchema>,
  "dataLabels"
> & {
  readonly dataLabels?:
    | false
    | ChartDataLabels<"left" | "right" | "bottom" | "top" | "center">
    | undefined;
};
export const LineChartGroupSchema =
  LineChartGroupRawSchema as unknown as z.ZodType<LineChartGroup>;

const AreaChartGroupRawSchema = z
  .strictObject({
    kind: z.literal("area"),
    data: SharedCategoryDataSchema,
    series: z.array(ChartSeriesSchema).min(1).readonly(),
    axes: ChartCategoryAxisReferencesSchema.optional(),
    stacking: z.enum(["stacked", "percent"]).optional(),
    dataLabels: z
      .union([z.literal(false), ChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly()
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  });

export type AreaChartGroup = Omit<
  z.infer<typeof AreaChartGroupRawSchema>,
  "dataLabels"
> & {
  readonly dataLabels?:
    | false
    | (Omit<ChartDataLabels, "position"> & { readonly position?: never })
    | undefined;
};
export const AreaChartGroupSchema =
  AreaChartGroupRawSchema as unknown as z.ZodType<AreaChartGroup>;

export type LineOrAreaChartGroup = LineChartGroup | AreaChartGroup;

const pieOptionsFields = {
  explosion: roundedIntInRangeSchema({
    minimum: 0,
    maximum: maximumExplosionForCurrentValidationProfile,
    boundary: "rounded-result",
    errorCode: "invalid-chart-option",
    label: "explosion",
  }).optional(),
  firstSliceAngle: roundedIntInRangeSchema({
    minimum: 0,
    maximum: 360,
    boundary: "input-and-result",
    errorCode: "invalid-chart-option",
    label: "firstSliceAngle",
  }).optional(),
};

const PieChartGroupRawSchema = z
  .strictObject({
    kind: z.literal("pie"),
    data: SharedCategoryDataSchema,
    series: z.array(ChartSeriesSchema).min(1).readonly(),
    ...pieOptionsFields,
    dataLabels: z
      .union([z.literal(false), ChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly()
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  });

export type PieChartGroup = Omit<
  z.infer<typeof PieChartGroupRawSchema>,
  "dataLabels"
> & {
  readonly dataLabels?:
    | false
    | ChartDataLabels<"center" | "insideEnd" | "outsideEnd" | "bestFit">
    | undefined;
};
export const PieChartGroupSchema =
  PieChartGroupRawSchema as unknown as z.ZodType<PieChartGroup>;

const doughnutOptionsFields = {
  kind: z.literal("doughnut"),
  ...pieOptionsFields,
  holeSize: roundedIntInRangeSchema({
    minimum: 10,
    maximum: 90,
    boundary: "input-and-result",
    errorCode: "invalid-chart-option",
    label: "holeSize",
  }).optional(),
  dataLabels: z.union([z.literal(false), ChartDataLabelsBaseSchema]).optional(),
};

const DoughnutSharedGroupRawSchema = z
  .strictObject({
    ...doughnutOptionsFields,
    data: SharedCategoryDataSchema,
    series: z.array(ChartSeriesSchema).min(1).readonly(),
  })
  .readonly();

const DoughnutIndependentGroupRawSchema = z
  .strictObject({
    ...doughnutOptionsFields,
    data: IndependentCategoryDataSchema,
    series: z.array(IndependentCategorySeriesSchema).min(1).readonly(),
  })
  .readonly();

export type DoughnutChartGroup =
  | (Omit<z.infer<typeof DoughnutSharedGroupRawSchema>, "dataLabels"> & {
      readonly dataLabels?:
        | false
        | (Omit<ChartDataLabels, "position"> & { readonly position?: never })
        | undefined;
    })
  | (Omit<z.infer<typeof DoughnutIndependentGroupRawSchema>, "dataLabels"> & {
      readonly dataLabels?:
        | false
        | (Omit<ChartDataLabels, "position"> & { readonly position?: never })
        | undefined;
    });

export const DoughnutChartGroupSchema = z
  .union([DoughnutSharedGroupRawSchema, DoughnutIndependentGroupRawSchema])
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  }) as unknown as z.ZodType<DoughnutChartGroup>;

export const ScatterChartGroupSchema = z
  .strictObject({
    kind: z.literal("scatter"),
    data: ScatterDataSchema,
    series: z.array(ScatterChartSeriesSchema).min(1).readonly(),
    axes: ChartXYAxisReferencesSchema.optional(),
    line: ScatterLineSchema.optional(),
    markers: z.boolean().optional(),
    dataLabels: z
      .union([z.literal(false), scatterDataLabelsObjectSchema.readonly()])
      .optional(),
  })
  .readonly()
  .superRefine((group, ctx) => {
    const line = group.line;
    const markers = group.markers ?? true;
    const showLine = line !== undefined && line !== false;
    if (!showLine && !markers) {
      addChartIssue(
        ctx,
        "invalid-scatter-presentation",
        "scatter group must draw markers or a connecting line",
      );
    }
    validateDataLabelSettingsForKind(
      group.dataLabels,
      "scatter",
      undefined,
      ctx,
      ["dataLabels"],
      true,
      false,
    );
    for (const [seriesIndex, item] of group.series.entries()) {
      const seriesPath = ["series", seriesIndex] as const;
      if (item.x.length !== item.y.length) {
        addChartIssue(
          ctx,
          "invalid-data",
          `Series "${item.name}" has ${item.x.length} X and ${item.y.length} Y observations`,
          [...seriesPath],
        );
      }
      if (!showLine && item.line !== undefined) {
        addChartIssue(
          ctx,
          "invalid-line",
          `Series "${item.name}".line requires a connecting line on its scatter group`,
          [...seriesPath, "line"],
        );
      }
      if (!markers && item.marker !== undefined && item.marker !== false) {
        addChartIssue(
          ctx,
          "invalid-marker",
          `Series "${item.name}".marker requires the scatter group to enable markers`,
          [...seriesPath, "marker"],
        );
      }
      validateDataLabelSettingsForKind(
        item.dataLabels,
        "scatter",
        undefined,
        ctx,
        [...seriesPath, "dataLabels"],
        false,
        false,
      );
      validatePointIndexKeys(
        item.points,
        item.x.length,
        item.name,
        ctx,
        seriesPath,
      );
      if (item.points !== undefined) {
        for (const [rawIndex, point] of Object.entries(item.points)) {
          if (
            !markers &&
            point.marker !== undefined &&
            point.marker !== false
          ) {
            addChartIssue(
              ctx,
              "invalid-marker",
              `Series "${item.name}" data point ${rawIndex}.marker requires the scatter group to enable markers`,
              [...seriesPath, "points", rawIndex, "marker"],
            );
          }
          validateDataLabelSettingsForKind(
            point.dataLabels,
            "scatter",
            undefined,
            ctx,
            [...seriesPath, "points", rawIndex, "dataLabels"],
            false,
            false,
          );
        }
      }
      if (item.errorBars !== undefined) {
        validateErrorBarsForKind(
          item.errorBars.x,
          "scatter",
          item.x.length,
          ctx,
          [...seriesPath, "errorBars", "x"],
        );
        validateErrorBarsForKind(
          item.errorBars.y,
          "scatter",
          item.y.length,
          ctx,
          [...seriesPath, "errorBars", "y"],
        );
      }
      validateTrendlinesForGroup(
        item.trendlines,
        "scatter",
        undefined,
        item.x.length,
        item.x.flatMap((xValue, pointIndex) => {
          const yValue = item.y[pointIndex];
          return xValue === null || yValue === null || yValue === undefined
            ? []
            : [{ x: xValue, y: yValue }];
        }),
        ctx,
        [...seriesPath, "trendlines"],
      );
    }
  });
export type ScatterChartGroup = z.infer<typeof ScatterChartGroupSchema>;

export type ChartGroup =
  | ColumnOrBarChartGroup
  | LineOrAreaChartGroup
  | PieChartGroup
  | DoughnutChartGroup
  | ScatterChartGroup;

export const ChartGroupSchema: z.ZodType<ChartGroup> = z.union([
  ColumnOrBarChartGroupSchema,
  LineChartGroupSchema,
  AreaChartGroupSchema,
  PieChartGroupSchema,
  DoughnutChartGroupSchema,
  ScatterChartGroupSchema,
]);

export const ChartLegendSchema = z.union([
  z.literal(false),
  z.enum(["top", "bottom", "left", "right", "topRight"]),
]);
export type ChartLegend = z.infer<typeof ChartLegendSchema>;

export const ChartPlotAreaSchema = z
  .strictObject({
    left: z.number().finite().nonnegative(),
    top: z.number().finite().nonnegative(),
    width: z.number().finite().positive(),
    height: z.number().finite().positive(),
  })
  .readonly();
export type ChartPlotArea = z.infer<typeof ChartPlotAreaSchema>;

export const ChartCoordinateSchema = z
  .number()
  .finite()
  .min(-1_000_000)
  .max(1_000_000);

export const ChartExtentSchema = z
  .number()
  .finite()
  .positive()
  .superRefine((value, ctx) => {
    const emu = pxToEmu(value);
    if (emu < 1 || emu > maximumShapeExtentEmu) {
      addChartIssue(
        ctx,
        "invalid-coordinate",
        `converted value must be within closed interval of 1 to ${maximumShapeExtentEmu} EMU`,
      );
    }
  });

const motionPropsShape = MotionPropsSchema.unwrap().shape;

const chartPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: xmlStringSchema("Chart.name").optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: MorphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  x: ChartCoordinateSchema.optional(),
  y: ChartCoordinateSchema.optional(),
  width: ChartExtentSchema.optional(),
  height: ChartExtentSchema.optional(),
  axes: ChartAxesSchema.optional(),
  groups: z.array(ChartGroupSchema).min(1).readonly(),
  title: ChartTitleSchema.optional(),
  displayBlanksAs: ChartDisplayBlanksAsSchema.optional(),
  legend: ChartLegendSchema.optional(),
  plotArea: ChartPlotAreaSchema.optional(),
  style: ChartStyleSchema.optional(),
  enter: motionPropsShape.enter,
  emphasis: motionPropsShape.emphasis,
  motion: motionPropsShape.motion,
});

export const ChartPropsSchema = chartPropsObjectSchema.readonly();

export const chartAllowedPropKeys: readonly string[] = Object.keys(
  chartPropsObjectSchema.shape,
);

export type ChartProps = z.infer<typeof ChartPropsSchema>;
export type ChartBase = ChartProps;

/** Maps a Zod validation issue inside chart authoring data to a stable `CapabilityError` code. */
export function resolveChartZodError(
  issue: z.core.$ZodIssue,
  fallbackCode: string,
  groupLabel?: string,
): {
  readonly code: string;
  readonly message: string;
  readonly fieldPath?: string;
} {
  if (
    issue.code === "custom" &&
    "params" in issue &&
    isPlainRecord(issue.params) &&
    typeof issue.params["errorCode"] === "string"
  ) {
    const rawFieldPath = issue.params["fieldPath"];
    const fieldPath =
      typeof rawFieldPath === "string"
        ? groupLabel !== undefined
          ? `${groupLabel}.${rawFieldPath}`
          : rawFieldPath
        : undefined;
    return {
      code: issue.params["errorCode"],
      message: issue.message,
      ...(fieldPath === undefined ? {} : { fieldPath }),
    };
  }
  const path = issue.path;
  if (issue.code === "unrecognized_keys") {
    return {
      code: "unknown-parameter",
      message: `${path.join(".") || "Chart"} contains unimplemented parameter ${issue.keys[0] ?? ""}`,
    };
  }
  const pathStr = path.map(String).join(".");
  if (pathStr === "fontSize" || pathStr.endsWith(".fontSize")) {
    return {
      code: "invalid-font-size",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (
    pathStr.endsWith(".color") ||
    pathStr.endsWith(".fill") ||
    pathStr.endsWith(".labelColor")
  ) {
    return { code: "invalid-color", message: `${pathStr}: ${issue.message}` };
  }
  if (
    pathStr.startsWith("line.") ||
    pathStr.startsWith("gridlines.") ||
    pathStr.startsWith("minorGridlines.") ||
    pathStr.includes(".line.") ||
    pathStr.includes(".leaderLines.") ||
    pathStr.includes(".style.") ||
    pathStr.startsWith("style.")
  ) {
    return { code: "invalid-style", message: `${pathStr}: ${issue.message}` };
  }
  if (pathStr.includes("trendlines")) {
    return {
      code: "invalid-trendline",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr.includes("errorBars")) {
    return {
      code: "invalid-error-bars",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr.includes("dataLabels")) {
    return {
      code: "invalid-data-labels",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr === "markers" || pathStr.endsWith(".markers")) {
    return {
      code: "invalid-scatter-markers",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr.includes("marker")) {
    return { code: "invalid-marker", message: `${pathStr}: ${issue.message}` };
  }
  if (pathStr.endsWith(".smooth") || pathStr === "smooth") {
    return { code: "invalid-smooth", message: `${pathStr}: ${issue.message}` };
  }
  if (pathStr.endsWith(".gapWidth")) {
    return {
      code: "invalid-gap-width",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr.endsWith(".overlap")) {
    return { code: "invalid-overlap", message: `${pathStr}: ${issue.message}` };
  }
  if (pathStr.endsWith(".stacking")) {
    return {
      code: "invalid-stacking",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (
    pathStr === "line" ||
    (pathStr.endsWith(".line") && !pathStr.includes("series"))
  ) {
    return {
      code: "invalid-scatter-line",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr.includes(".line")) {
    return { code: "invalid-style", message: `${pathStr}: ${issue.message}` };
  }
  if (pathStr.includes(".axes")) {
    return {
      code: "invalid-chart-axes",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr.startsWith("axes")) {
    return { code: "invalid-axis", message: `${pathStr}: ${issue.message}` };
  }
  if (pathStr.includes(".series") || pathStr.includes(".data")) {
    return { code: "invalid-data", message: `${pathStr}: ${issue.message}` };
  }
  if (pathStr.startsWith("plotArea")) {
    return {
      code: "invalid-plot-area",
      message: `${pathStr}: ${issue.message}`,
    };
  }
  if (pathStr === "legend") {
    return { code: "invalid-legend", message: issue.message };
  }
  if (pathStr === "displayBlanksAs") {
    return { code: "invalid-chart-missing-values", message: issue.message };
  }
  if (pathStr === "groups") {
    return { code: "invalid-chart-groups", message: issue.message };
  }
  return { code: fallbackCode, message: `${pathStr}: ${issue.message}` };
}

function throwChartZodError(
  error: z.ZodError,
  fallbackCode: string,
  label: string,
): never {
  const issue = error.issues[0];
  if (issue === undefined) {
    throw new CapabilityError(fallbackCode, `${label} is invalid`);
  }
  const resolved = resolveChartZodError(issue, fallbackCode, label);
  throw new CapabilityError(
    resolved.code,
    resolved.message,
    resolved.fieldPath,
  );
}

/** Validates and parses a chart axis entry using the local `ChartAxis` Zod schemas. */
export function parseChartAxisBySchema(
  value: unknown,
  label: string,
  expected: "category" | "value",
): ChartAxis {
  if (!isPlainRecord(value)) {
    throw new CapabilityError("invalid-axis", `${label} must be an object`);
  }
  if (value["kind"] !== expected) {
    throw new CapabilityError(
      "invalid-chart-axis-role",
      `${label} must declare kind "${expected}" as referenced by its chart group`,
    );
  }
  const schema =
    expected === "category" ? ChartCategoryAxisSchema : ChartValueAxisSchema;
  const parsed = schema.safeParse(value);
  if (!parsed.success) {
    const first = parsed.error.issues[0];
    if (
      expected === "category" &&
      first?.code === "unrecognized_keys" &&
      first.keys.some((k) =>
        [
          "min",
          "max",
          "majorUnit",
          "minorUnit",
          "numberFormat",
          "logBase",
        ].includes(k),
      )
    ) {
      throw new CapabilityError(
        "invalid-axis",
        `${label}.${first.keys[0]} is only applicable to value axis`,
      );
    }
    throwChartZodError(parsed.error, "invalid-axis", label);
  }
  return parsed.data;
}

const groupOptionsByKind: Readonly<Record<ChartKind, readonly string[]>> = {
  column: ["stacking", "gapWidth", "overlap"],
  bar: ["stacking", "gapWidth", "overlap"],
  line: ["stacking"],
  area: ["stacking"],
  pie: ["explosion", "firstSliceAngle"],
  doughnut: ["explosion", "firstSliceAngle", "holeSize"],
  scatter: ["line", "markers"],
};

/** Validates and parses one chart group entry via its kind-specific local Zod schema. */
export function parseChartGroupBySchema(
  value: unknown,
  label: string,
): ChartGroup {
  if (!isPlainRecord(value)) {
    throw new CapabilityError(
      "invalid-chart-groups",
      `${label} must be an object`,
    );
  }
  const kindParsed = ChartKindSchema.safeParse(value["kind"]);
  if (!kindParsed.success) {
    throw new CapabilityError(
      "unsupported-chart-kind",
      `Unsupported chart kind ${String(value["kind"])}`,
    );
  }
  const kind = kindParsed.data;
  for (const key of Object.keys(value)) {
    if (
      key !== "kind" &&
      key !== "data" &&
      key !== "series" &&
      key !== "axes" &&
      key !== "dataLabels" &&
      !groupOptionsByKind[kind].includes(key)
    ) {
      const isKnownOtherGroupOption = Object.values(groupOptionsByKind)
        .flat()
        .includes(key);
      throw new CapabilityError(
        isKnownOtherGroupOption ? "invalid-chart-option" : "unknown-parameter",
        isKnownOtherGroupOption
          ? `${kind} chart group does not support ${key}`
          : `${label} contains unimplemented parameter ${key}`,
      );
    }
  }
  if (kind === "scatter") {
    const parsed = ScatterChartGroupSchema.safeParse(value);
    if (!parsed.success) {
      throwChartZodError(parsed.error, "invalid-data", label);
    }
    return parsed.data;
  }
  if (kind === "column" || kind === "bar") {
    const schema =
      value["stacking"] === undefined
        ? ColumnOrBarUnstackedGroupSchema
        : ColumnOrBarStackedGroupSchema;
    const parsed = schema
      .superRefine((group, ctx) => validateCategoryGroupSeries(group, ctx))
      .safeParse(value);
    if (!parsed.success) {
      throwChartZodError(parsed.error, "invalid-data", label);
    }
    return parsed.data as unknown as ColumnOrBarChartGroup;
  }
  if (kind === "line") {
    const parsed = LineChartGroupSchema.safeParse(value);
    if (!parsed.success) {
      throwChartZodError(parsed.error, "invalid-data", label);
    }
    return parsed.data;
  }
  if (kind === "area") {
    const parsed = AreaChartGroupSchema.safeParse(value);
    if (!parsed.success) {
      throwChartZodError(parsed.error, "invalid-data", label);
    }
    return parsed.data;
  }
  if (kind === "pie") {
    const parsed = PieChartGroupSchema.safeParse(value);
    if (!parsed.success) {
      throwChartZodError(parsed.error, "invalid-data", label);
    }
    return parsed.data;
  }
  // doughnut
  const isIndependent =
    isPlainRecord(value["data"]) &&
    value["data"]["kind"] === "independent-categories";
  const doughnutSchema = (
    isIndependent
      ? DoughnutIndependentGroupRawSchema
      : DoughnutSharedGroupRawSchema
  ).superRefine((group, ctx) => validateCategoryGroupSeries(group, ctx));
  const parsed = doughnutSchema.safeParse(value);
  if (!parsed.success) {
    throwChartZodError(parsed.error, "invalid-data", label);
  }
  return parsed.data as unknown as DoughnutChartGroup;
}

/** Validates and parses `Chart.title` using `ChartTitleSchema`. */
export function parseChartTitleBySchema(
  value: unknown,
  label: string,
): ChartTitle | undefined {
  if (value === undefined) return undefined;
  const parsed = ChartTitleSchema.safeParse(value);
  if (!parsed.success) {
    throwChartZodError(parsed.error, "invalid-parameter", label);
  }
  return parsed.data;
}

/** Validates and parses `Chart.plotArea` using `ChartPlotAreaSchema` + chart box bounds. */
export function parseChartPlotAreaBySchema(
  value: unknown,
  box?: { readonly width: number; readonly height: number },
): ChartPlotArea | undefined {
  if (value === undefined) return undefined;
  const parsed = ChartPlotAreaSchema.safeParse(value);
  if (!parsed.success) {
    throwChartZodError(parsed.error, "invalid-plot-area", "Chart.plotArea");
  }
  const result = parsed.data;
  if (
    box !== undefined &&
    (result.left + result.width > box.width ||
      result.top + result.height > box.height)
  ) {
    throw new CapabilityError(
      "invalid-plot-area",
      "Chart.plotArea must be inside chart box with positive width and height",
    );
  }
  return result;
}

/** Validates and parses `Chart.legend` using `ChartLegendSchema`. */
export function parseChartLegendBySchema(
  value: unknown,
): ChartLegend | undefined {
  if (value === undefined) return undefined;
  const parsed = ChartLegendSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-legend",
      "Chart.legend must be false, top, bottom, left, right, or topRight",
    );
  }
  return parsed.data;
}

/** Validates and parses a chart coordinate (`x` or `y`) via `ChartCoordinateSchema`. */
export function parseChartCoordinateBySchema(
  value: unknown,
  label: "x" | "y",
): number {
  if (value === undefined) return 0;
  const parsed = ChartCoordinateSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${label} must be a finite pixel value between -1000000 and 1000000`,
    );
  }
  return parsed.data;
}

/** Validates and parses a chart extent (`width` or `height`) via `ChartExtentSchema`. */
export function parseChartExtentBySchema(
  value: unknown,
  label: "width" | "height",
): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${label} must be a finite positive pixel value`,
    );
  }
  const parsed = ChartExtentSchema.safeParse(value);
  if (!parsed.success) {
    throwChartZodError(parsed.error, "invalid-coordinate", `Chart.${label}`);
  }
  return parsed.data;
}

/** Validates and parses `Chart.displayBlanksAs` via `ChartDisplayBlanksAsSchema`. */
export function parseChartDisplayBlanksAsBySchema(
  value: unknown,
): ChartDisplayBlanksAs | undefined {
  if (value === undefined) return undefined;
  const parsed = ChartDisplayBlanksAsSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-chart-missing-values",
      "Chart.displayBlanksAs must be gap, zero, or span",
    );
  }
  return parsed.data;
}

/** Validates an XML 1.0 string via `xmlStringSchema`. */
export function parseChartXmlTextBySchema(
  value: string,
  label: string,
): string {
  const parsed = xmlStringSchema(label).safeParse(value);
  if (!parsed.success) {
    throwChartZodError(parsed.error, "invalid-xml-character", label);
  }
  return parsed.data;
}
