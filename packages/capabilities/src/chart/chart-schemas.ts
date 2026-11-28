import {
  CapabilityError,
  authorIdSchema,
  colorValueSchema,
  layoutStyleShape,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  textStyleShape,
  type AuthorId,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type NativePlaceholderRef,
  type TextStyle,
} from "@pptxdk/core/capability";
import { pxToEmu, xmlEscape } from "@pptxdk/emit";
import { z } from "zod";

import {
  motionPropsShape,
  type MotionProps,
} from "../motion/motion-schemas.ts";
import { morphKeySchema, type MorphKey } from "../shared/morph-api.ts";
import { chartDataLabelPositions } from "./chart-labels.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

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

export type ChartKind =
  | "column"
  | "bar"
  | "line"
  | "area"
  | "pie"
  | "doughnut"
  | "scatter";
const rawChartKindSchema = z.enum([
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "doughnut",
  "scatter",
]);
export const chartKindSchema: z.ZodType<ChartKind> =
  rawChartKindSchema satisfies z.ZodType<
    AssertExact<ChartKind, z.infer<typeof rawChartKindSchema>> extends true
      ? ChartKind
      : never
  >;

export type ChartTextStyle = Omit<TextStyle, "textWrap">;
const rawChartTextStyleObjectSchema = z
  .strictObject(textStyleShape)
  .omit({ textWrap: true });
const rawChartTextStyleSchema = rawChartTextStyleObjectSchema.readonly();
export const chartTextStyleKeys: readonly string[] = Object.keys(
  rawChartTextStyleObjectSchema.shape,
);
export const chartTextStyleSchema: z.ZodType<ChartTextStyle> =
  rawChartTextStyleSchema satisfies z.ZodType<
    AssertExact<
      ChartTextStyle,
      z.infer<typeof rawChartTextStyleSchema>
    > extends true
      ? ChartTextStyle
      : never
  >;

export type ChartStyle = LayoutStyle & ChartTextStyle;
const rawChartStyleObjectSchema = z.strictObject({
  ...rawChartTextStyleObjectSchema.shape,
  ...layoutStyleShape,
});
const rawChartStyleSchema = rawChartStyleObjectSchema.readonly();
export const chartStyleKeys: readonly string[] = Object.keys(
  rawChartStyleObjectSchema.shape,
);
export const chartStyleSchema: z.ZodType<ChartStyle> =
  rawChartStyleSchema satisfies z.ZodType<
    AssertExact<ChartStyle, z.infer<typeof rawChartStyleSchema>> extends true
      ? ChartStyle
      : never
  >;

export type ChartTitle =
  | string
  | {
      readonly text: string;
      readonly style?: ChartTextStyle | undefined;
    };
const rawChartTitleSchema = z.union([
  nonEmptyXmlStringSchema("invalid-parameter", "title"),
  z
    .strictObject({
      text: nonEmptyXmlStringSchema("invalid-parameter", "title.text"),
      style: rawChartTextStyleSchema.optional(),
    })
    .readonly(),
]);
export const chartTitleSchema: z.ZodType<ChartTitle> =
  rawChartTitleSchema satisfies z.ZodType<
    AssertExact<ChartTitle, z.infer<typeof rawChartTitleSchema>> extends true
      ? ChartTitle
      : never
  >;

export type ChartLineDash =
  | "solid"
  | "dash"
  | "dot"
  | "dashDot"
  | "longDash"
  | "longDashDot"
  | "longDashDotDot";
const rawChartLineDashSchema = z.enum([
  "solid",
  "dash",
  "dot",
  "dashDot",
  "longDash",
  "longDashDot",
  "longDashDotDot",
]);
export const chartLineDashSchema: z.ZodType<ChartLineDash> =
  rawChartLineDashSchema satisfies z.ZodType<
    AssertExact<
      ChartLineDash,
      z.infer<typeof rawChartLineDashSchema>
    > extends true
      ? ChartLineDash
      : never
  >;

export type ChartLine = {
  readonly color?: ColorValue | undefined;
  readonly width?: number | undefined;
  readonly dash?: ChartLineDash | undefined;
};
const rawChartLineSchema = z
  .strictObject({
    color: colorValueSchema.optional(),
    width: z
      .number()
      .finite()
      .positive()
      .refine((value) => pxToEmu(value) <= maximumLineWidthEmu, {
        message: "width converted value exceeds writable line width range",
      })
      .optional(),
    dash: rawChartLineDashSchema.optional(),
  })
  .readonly();
export const chartLineSchema: z.ZodType<ChartLine> =
  rawChartLineSchema satisfies z.ZodType<
    AssertExact<ChartLine, z.infer<typeof rawChartLineSchema>> extends true
      ? ChartLine
      : never
  >;

export type ChartMarkerSymbol =
  | "auto"
  | "circle"
  | "dash"
  | "diamond"
  | "dot"
  | "none"
  | "plus"
  | "square"
  | "star"
  | "triangle"
  | "x";
const rawChartMarkerSymbolSchema = z.enum([
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
export const chartMarkerSymbolSchema: z.ZodType<ChartMarkerSymbol> =
  rawChartMarkerSymbolSchema satisfies z.ZodType<
    AssertExact<
      ChartMarkerSymbol,
      z.infer<typeof rawChartMarkerSymbolSchema>
    > extends true
      ? ChartMarkerSymbol
      : never
  >;

export type ChartMarker = {
  readonly symbol?: ChartMarkerSymbol | undefined;
  readonly size?: number | undefined;
  readonly fill?: ColorValue | undefined;
  readonly line?: ChartLine | undefined;
};
const rawChartMarkerSchema = z
  .strictObject({
    symbol: rawChartMarkerSymbolSchema.optional(),
    size: z.number().int().min(2).max(72).optional(),
    fill: colorValueSchema.optional(),
    line: rawChartLineSchema.optional(),
  })
  .readonly();
export const chartMarkerSchema: z.ZodType<ChartMarker> =
  rawChartMarkerSchema satisfies z.ZodType<
    AssertExact<ChartMarker, z.infer<typeof rawChartMarkerSchema>> extends true
      ? ChartMarker
      : never
  >;

const rawChartMarkerOptionSchema = z.union([z.boolean(), rawChartMarkerSchema]);
export const chartMarkerOptionSchema: z.ZodType<boolean | ChartMarker> =
  rawChartMarkerOptionSchema;

export const chartPointValueSchema: z.ZodType<number | null> = z
  .number()
  .finite()
  .nullable();
const rawChartSeriesValuesSchema = z.array(chartPointValueSchema).readonly();
export const chartSeriesValuesSchema: z.ZodType<readonly (number | null)[]> =
  rawChartSeriesValuesSchema;
const rawChartCategoriesSchema = z
  .array(xmlStringSchema("category"))
  .readonly();
export const chartCategoriesSchema: z.ZodType<readonly string[]> =
  rawChartCategoriesSchema;

export type ChartDataLabelPosition =
  | "center"
  | "insideEnd"
  | "insideBase"
  | "outsideEnd"
  | "bestFit"
  | "left"
  | "right"
  | "bottom"
  | "top";
const rawChartDataLabelPositionSchema = z.enum([
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
export const chartDataLabelPositionSchema: z.ZodType<ChartDataLabelPosition> =
  rawChartDataLabelPositionSchema satisfies z.ZodType<
    AssertExact<
      ChartDataLabelPosition,
      z.infer<typeof rawChartDataLabelPositionSchema>
    > extends true
      ? ChartDataLabelPosition
      : never
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
  leaderLines: z.union([z.boolean(), rawChartLineSchema]).optional(),
  style: rawChartTextStyleSchema.optional(),
};

const chartDataLabelsObjectSchema = z.strictObject({
  ...dataLabelCommonFields,
  position: rawChartDataLabelPositionSchema.optional(),
});

const rawChartDataLabelsBaseSchema = chartDataLabelsObjectSchema.readonly();
export const chartDataLabelsBaseSchema: z.ZodType<ChartDataLabelSettings> =
  rawChartDataLabelsBaseSchema;

export type ChartDataLabels<
  Position extends ChartDataLabelPosition = ChartDataLabelPosition,
> = {
  readonly value?: boolean | undefined;
  readonly percent?: boolean | undefined;
  readonly category?: boolean | undefined;
  readonly seriesName?: boolean | undefined;
  readonly separator?: string | undefined;
  readonly numberFormat?: string | undefined;
  readonly leaderLines?: boolean | ChartLine | undefined;
  readonly style?: ChartTextStyle | undefined;
  readonly position?: Position | undefined;
};

export type ChartDataLabelOverride = {
  readonly value?: boolean | undefined;
  readonly percent?: boolean | undefined;
  readonly category?: boolean | undefined;
  readonly seriesName?: boolean | undefined;
  readonly separator?: string | undefined;
  readonly numberFormat?: string | undefined;
  readonly position?: ChartDataLabelPosition | undefined;
};
const rawChartDataLabelOverrideSchema = chartDataLabelsObjectSchema
  .omit({
    style: true,
    leaderLines: true,
    xValue: true,
  })
  .readonly();
export const chartDataLabelOverrideSchema: z.ZodType<ChartDataLabelOverride> =
  rawChartDataLabelOverrideSchema satisfies z.ZodType<
    AssertExact<
      ChartDataLabelOverride,
      z.infer<typeof rawChartDataLabelOverrideSchema>
    > extends true
      ? ChartDataLabelOverride
      : never
  >;

export type ChartSeriesDataLabelOverride = {
  readonly value?: boolean | undefined;
  readonly percent?: boolean | undefined;
  readonly category?: boolean | undefined;
  readonly seriesName?: boolean | undefined;
  readonly separator?: string | undefined;
  readonly numberFormat?: string | undefined;
  readonly leaderLines?: boolean | ChartLine | undefined;
  readonly position?: ChartDataLabelPosition | undefined;
};
const rawChartSeriesDataLabelOverrideSchema = chartDataLabelsObjectSchema
  .omit({
    style: true,
    xValue: true,
  })
  .readonly();
export const chartSeriesDataLabelOverrideSchema: z.ZodType<ChartSeriesDataLabelOverride> =
  rawChartSeriesDataLabelOverrideSchema satisfies z.ZodType<
    AssertExact<
      ChartSeriesDataLabelOverride,
      z.infer<typeof rawChartSeriesDataLabelOverrideSchema>
    > extends true
      ? ChartSeriesDataLabelOverride
      : never
  >;

export type ChartDataLabelSettings = {
  readonly value?: boolean | undefined;
  readonly percent?: boolean | undefined;
  readonly category?: boolean | undefined;
  readonly seriesName?: boolean | undefined;
  readonly xValue?: boolean | undefined;
  readonly separator?: string | undefined;
  readonly numberFormat?: string | undefined;
  readonly leaderLines?: boolean | ChartLine | undefined;
  readonly style?: ChartTextStyle | undefined;
  readonly position?: ChartDataLabelPosition | undefined;
};
const rawChartDataLabelSettingsSchema = chartDataLabelsObjectSchema.readonly();
export const chartDataLabelSettingsSchema: z.ZodType<ChartDataLabelSettings> =
  rawChartDataLabelSettingsSchema satisfies z.ZodType<
    AssertExact<
      ChartDataLabelSettings,
      z.infer<typeof rawChartDataLabelSettingsSchema>
    > extends true
      ? ChartDataLabelSettings
      : never
  >;

const scatterDataLabelsObjectSchema = z.strictObject({
  ...dataLabelCommonFields,
  position: z.enum(["left", "right", "bottom", "top", "center"]).optional(),
});

export type ScatterDataLabels = {
  readonly value?: boolean | undefined;
  readonly seriesName?: boolean | undefined;
  readonly xValue?: boolean | undefined;
  readonly separator?: string | undefined;
  readonly numberFormat?: string | undefined;
  readonly style?: ChartTextStyle | undefined;
  readonly position?:
    | "left"
    | "right"
    | "bottom"
    | "top"
    | "center"
    | undefined;
};
const rawScatterDataLabelsBaseSchema = scatterDataLabelsObjectSchema
  .omit({ percent: true, category: true, leaderLines: true })
  .readonly();
export const scatterDataLabelsBaseSchema: z.ZodType<ScatterDataLabels> =
  rawScatterDataLabelsBaseSchema satisfies z.ZodType<
    AssertExact<
      ScatterDataLabels,
      z.infer<typeof rawScatterDataLabelsBaseSchema>
    > extends true
      ? ScatterDataLabels
      : never
  >;

export type ScatterDataLabelOverride = {
  readonly value?: boolean | undefined;
  readonly seriesName?: boolean | undefined;
  readonly xValue?: boolean | undefined;
  readonly separator?: string | undefined;
  readonly numberFormat?: string | undefined;
  readonly position?:
    | "left"
    | "right"
    | "bottom"
    | "top"
    | "center"
    | undefined;
};
const rawScatterDataLabelOverrideSchema = scatterDataLabelsObjectSchema
  .omit({
    style: true,
    percent: true,
    category: true,
    leaderLines: true,
  })
  .readonly();
export const scatterDataLabelOverrideSchema: z.ZodType<ScatterDataLabelOverride> =
  rawScatterDataLabelOverrideSchema satisfies z.ZodType<
    AssertExact<
      ScatterDataLabelOverride,
      z.infer<typeof rawScatterDataLabelOverrideSchema>
    > extends true
      ? ScatterDataLabelOverride
      : never
  >;

const ChartPointRawSchema = z
  .strictObject({
    color: colorValueSchema.optional(),
    line: rawChartLineSchema.optional(),
    marker: z.unknown().optional(),
    explosion: roundedIntInRangeSchema({
      minimum: 0,
      maximum: maximumExplosionForCurrentValidationProfile,
      boundary: "rounded-result",
      errorCode: "invalid-chart-option",
      label: "explosion",
    }).optional(),
    labelColor: colorValueSchema.optional(),
    dataLabels: z
      .union([z.literal(false), rawChartDataLabelOverrideSchema])
      .optional(),
  })
  .readonly();

export type ChartPoint = {
  readonly color?: ColorValue | undefined;
  readonly line?: ChartLine | undefined;
  readonly explosion?: number | undefined;
  readonly labelColor?: ColorValue | undefined;
  readonly marker?: boolean | ChartMarker | undefined;
  readonly dataLabels?: false | ChartDataLabelOverride | undefined;
};
export const chartPointSchema: z.ZodType<ChartPoint> =
  ChartPointRawSchema as unknown as z.ZodType<
    AssertExact<
      Omit<ChartPoint, "marker">,
      Omit<z.infer<typeof ChartPointRawSchema>, "marker">
    > extends true
      ? ChartPoint
      : never
  >;

export type ScatterChartPoint = {
  readonly color?: ColorValue | undefined;
  readonly line?: ChartLine | undefined;
  readonly marker?: boolean | ChartMarker | undefined;
  readonly labelColor?: ColorValue | undefined;
  readonly dataLabels?: false | ScatterDataLabelOverride | undefined;
};
const rawScatterChartPointSchema = z
  .strictObject({
    color: colorValueSchema.optional(),
    line: rawChartLineSchema.optional(),
    marker: rawChartMarkerOptionSchema.optional(),
    labelColor: colorValueSchema.optional(),
    dataLabels: z
      .union([z.literal(false), rawScatterDataLabelOverrideSchema])
      .optional(),
  })
  .readonly();
export const scatterChartPointSchema: z.ZodType<ScatterChartPoint> =
  rawScatterChartPointSchema satisfies z.ZodType<
    AssertExact<
      ScatterChartPoint,
      z.infer<typeof rawScatterChartPointSchema>
    > extends true
      ? ScatterChartPoint
      : never
  >;

export type ChartTrendlineType =
  | "linear"
  | "exponential"
  | "logarithmic"
  | "polynomial"
  | "power"
  | "movingAverage";
const rawChartTrendlineTypeSchema = z.enum([
  "linear",
  "exponential",
  "logarithmic",
  "polynomial",
  "power",
  "movingAverage",
]);
export const chartTrendlineTypeSchema: z.ZodType<ChartTrendlineType> =
  rawChartTrendlineTypeSchema satisfies z.ZodType<
    AssertExact<
      ChartTrendlineType,
      z.infer<typeof rawChartTrendlineTypeSchema>
    > extends true
      ? ChartTrendlineType
      : never
  >;

const trendlineBaseFields = {
  line: rawChartLineSchema.optional(),
  forward: z.number().finite().positive().optional(),
  backward: z.number().finite().positive().optional(),
};

const trendlineDisplayFields = {
  displayEquation: z.boolean().optional(),
  displayRSquared: z.boolean().optional(),
};

export type ChartTrendline =
  | {
      readonly line?: ChartLine | undefined;
      readonly forward?: number | undefined;
      readonly backward?: number | undefined;
      readonly displayEquation?: boolean | undefined;
      readonly displayRSquared?: boolean | undefined;
      readonly type: "linear" | "exponential";
      readonly intercept?: number | undefined;
      readonly order?: undefined;
      readonly period?: undefined;
    }
  | {
      readonly line?: ChartLine | undefined;
      readonly forward?: number | undefined;
      readonly backward?: number | undefined;
      readonly displayEquation?: boolean | undefined;
      readonly displayRSquared?: boolean | undefined;
      readonly type: "logarithmic" | "power";
      readonly intercept?: undefined;
      readonly order?: undefined;
      readonly period?: undefined;
    }
  | {
      readonly line?: ChartLine | undefined;
      readonly forward?: number | undefined;
      readonly backward?: number | undefined;
      readonly displayEquation?: boolean | undefined;
      readonly displayRSquared?: boolean | undefined;
      readonly type: "polynomial";
      readonly order: number;
      readonly intercept?: number | undefined;
      readonly period?: undefined;
    }
  | {
      readonly line?: ChartLine | undefined;
      readonly forward?: number | undefined;
      readonly backward?: number | undefined;
      readonly type: "movingAverage";
      readonly period: number;
      readonly intercept?: undefined;
      readonly order?: undefined;
      readonly displayEquation?: undefined;
      readonly displayRSquared?: undefined;
    };
const rawChartTrendlineSchema = z.discriminatedUnion("type", [
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
export const chartTrendlineSchema: z.ZodType<ChartTrendline> =
  rawChartTrendlineSchema satisfies z.ZodType<
    AssertExact<
      ChartTrendline,
      z.infer<typeof rawChartTrendlineSchema>
    > extends true
      ? ChartTrendline
      : never
  >;

export type ChartErrorBarValues = readonly (number | null)[];
const rawChartErrorBarValuesSchema = z
  .array(z.number().finite().nonnegative().nullable())
  .readonly();
export const chartErrorBarValuesSchema: z.ZodType<ChartErrorBarValues> =
  rawChartErrorBarValuesSchema satisfies z.ZodType<
    AssertExact<
      ChartErrorBarValues,
      z.infer<typeof rawChartErrorBarValuesSchema>
    > extends true
      ? ChartErrorBarValues
      : never
  >;

export type ChartErrorBarInclude = "both" | "plus" | "minus";
const rawChartErrorBarIncludeSchema = z.enum(["both", "plus", "minus"]);
export const chartErrorBarIncludeSchema: z.ZodType<ChartErrorBarInclude> =
  rawChartErrorBarIncludeSchema satisfies z.ZodType<
    AssertExact<
      ChartErrorBarInclude,
      z.infer<typeof rawChartErrorBarIncludeSchema>
    > extends true
      ? ChartErrorBarInclude
      : never
  >;

const errorBarStyleFields = {
  endCap: z.boolean().optional(),
  line: rawChartLineSchema.optional(),
};

export type ChartCustomErrorBars =
  | {
      readonly endCap?: boolean | undefined;
      readonly line?: ChartLine | undefined;
      readonly type: "custom";
      readonly include?: "both" | undefined;
      readonly plus: ChartErrorBarValues;
      readonly minus: ChartErrorBarValues;
      readonly value?: undefined;
    }
  | {
      readonly endCap?: boolean | undefined;
      readonly line?: ChartLine | undefined;
      readonly type: "custom";
      readonly include: "plus";
      readonly plus: ChartErrorBarValues;
      readonly minus?: undefined;
      readonly value?: undefined;
    }
  | {
      readonly endCap?: boolean | undefined;
      readonly line?: ChartLine | undefined;
      readonly type: "custom";
      readonly include: "minus";
      readonly minus: ChartErrorBarValues;
      readonly plus?: undefined;
      readonly value?: undefined;
    };
const rawChartCustomErrorBarsSchema = z.union([
  z
    .strictObject({
      ...errorBarStyleFields,
      type: z.literal("custom"),
      include: z.literal("both").optional(),
      plus: rawChartErrorBarValuesSchema,
      minus: rawChartErrorBarValuesSchema,
      value: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...errorBarStyleFields,
      type: z.literal("custom"),
      include: z.literal("plus"),
      plus: rawChartErrorBarValuesSchema,
      minus: z.never().optional(),
      value: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...errorBarStyleFields,
      type: z.literal("custom"),
      include: z.literal("minus"),
      minus: rawChartErrorBarValuesSchema,
      plus: z.never().optional(),
      value: z.never().optional(),
    })
    .readonly(),
]);
export const chartCustomErrorBarsSchema: z.ZodType<ChartCustomErrorBars> =
  rawChartCustomErrorBarsSchema satisfies z.ZodType<
    AssertExact<
      ChartCustomErrorBars,
      z.infer<typeof rawChartCustomErrorBarsSchema>
    > extends true
      ? ChartCustomErrorBars
      : never
  >;

export type ChartErrorBars =
  | {
      readonly endCap?: boolean | undefined;
      readonly line?: ChartLine | undefined;
      readonly include?: ChartErrorBarInclude | undefined;
      readonly type: "fixed" | "percentage" | "standardDeviation";
      readonly value: number;
      readonly plus?: undefined;
      readonly minus?: undefined;
    }
  | {
      readonly endCap?: boolean | undefined;
      readonly line?: ChartLine | undefined;
      readonly include?: ChartErrorBarInclude | undefined;
      readonly type: "standardError";
      readonly value?: undefined;
      readonly plus?: undefined;
      readonly minus?: undefined;
    }
  | ChartCustomErrorBars;
const rawChartErrorBarsSchema = z.union([
  z
    .strictObject({
      ...errorBarStyleFields,
      include: rawChartErrorBarIncludeSchema.optional(),
      type: z.enum(["fixed", "percentage", "standardDeviation"]),
      value: z.number().finite().nonnegative(),
      plus: z.never().optional(),
      minus: z.never().optional(),
    })
    .readonly(),
  z
    .strictObject({
      ...errorBarStyleFields,
      include: rawChartErrorBarIncludeSchema.optional(),
      type: z.literal("standardError"),
      value: z.never().optional(),
      plus: z.never().optional(),
      minus: z.never().optional(),
    })
    .readonly(),
  rawChartCustomErrorBarsSchema,
]);
export const chartErrorBarsSchema: z.ZodType<ChartErrorBars> =
  rawChartErrorBarsSchema satisfies z.ZodType<
    AssertExact<
      ChartErrorBars,
      z.infer<typeof rawChartErrorBarsSchema>
    > extends true
      ? ChartErrorBars
      : never
  >;

export type ScatterErrorBars = {
  readonly x?: ChartErrorBars | undefined;
  readonly y?: ChartErrorBars | undefined;
};
const rawScatterErrorBarsSchema = z
  .strictObject({
    x: rawChartErrorBarsSchema.optional(),
    y: rawChartErrorBarsSchema.optional(),
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
export const scatterErrorBarsSchema: z.ZodType<ScatterErrorBars> =
  rawScatterErrorBarsSchema satisfies z.ZodType<
    AssertExact<
      ScatterErrorBars,
      z.infer<typeof rawScatterErrorBarsSchema>
    > extends true
      ? ScatterErrorBars
      : never
  >;

const chartSeriesFormatFields = {
  name: nonEmptyStringSchema(
    "invalid-data",
    "Series.name must be a non-empty string",
  ),
  color: colorValueSchema.optional(),
  line: rawChartLineSchema.optional(),
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
  points: z.record(z.string(), chartPointSchema).readonly().optional(),
  trendlines: z.array(rawChartTrendlineSchema).readonly().optional(),
};

const ChartSeriesRawSchema = z
  .strictObject({
    ...chartSeriesFormatFields,
    values: rawChartSeriesValuesSchema,
    smooth: z.boolean().optional(),
    errorBars: rawChartErrorBarsSchema.optional(),
  })
  .readonly();

export type ChartSeries = {
  readonly name: string;
  readonly color?: ColorValue | undefined;
  readonly line?: ChartLine | undefined;
  readonly explosion?: number | undefined;
  readonly dataLabels?: false | ChartSeriesDataLabelOverride | undefined;
  readonly marker?: boolean | ChartMarker | undefined;
  readonly points?: Readonly<Record<number, ChartPoint>> | undefined;
  readonly trendlines?: readonly ChartTrendline[] | undefined;
  readonly values: readonly (number | null)[];
  readonly smooth?: boolean | undefined;
  readonly errorBars?: ChartErrorBars | undefined;
};

export const chartSeriesSchema: z.ZodType<ChartSeries> =
  ChartSeriesRawSchema as unknown as z.ZodType<
    AssertExact<
      Omit<ChartSeries, "points" | "dataLabels" | "marker">,
      Omit<
        z.infer<typeof ChartSeriesRawSchema>,
        "points" | "dataLabels" | "marker"
      >
    > extends true
      ? ChartSeries
      : never
  >;

const IndependentCategorySeriesRawSchema = z
  .strictObject({
    ...chartSeriesFormatFields,
    categories: rawChartCategoriesSchema,
    values: rawChartSeriesValuesSchema,
    smooth: z.boolean().optional(),
    errorBars: rawChartErrorBarsSchema.optional(),
  })
  .readonly();

export type IndependentCategorySeries = {
  readonly name: string;
  readonly color?: ColorValue | undefined;
  readonly line?: ChartLine | undefined;
  readonly explosion?: number | undefined;
  readonly dataLabels?: false | ChartSeriesDataLabelOverride | undefined;
  readonly marker?: boolean | ChartMarker | undefined;
  readonly points?: Readonly<Record<number, ChartPoint>> | undefined;
  readonly trendlines?: readonly ChartTrendline[] | undefined;
  readonly categories: readonly string[];
  readonly values: readonly (number | null)[];
  readonly smooth?: boolean | undefined;
  readonly errorBars?: ChartErrorBars | undefined;
};

export const independentCategorySeriesSchema: z.ZodType<IndependentCategorySeries> =
  IndependentCategorySeriesRawSchema as unknown as z.ZodType<
    AssertExact<
      Omit<IndependentCategorySeries, "points" | "dataLabels" | "marker">,
      Omit<
        z.infer<typeof IndependentCategorySeriesRawSchema>,
        "points" | "dataLabels" | "marker"
      >
    > extends true
      ? IndependentCategorySeries
      : never
  >;

const ScatterChartSeriesRawSchema = z
  .strictObject({
    name: nonEmptyStringSchema(
      "invalid-data",
      "Series.name must be a non-empty string",
    ),
    x: rawChartSeriesValuesSchema,
    y: rawChartSeriesValuesSchema,
    color: colorValueSchema.optional(),
    line: rawChartLineSchema.optional(),
    dataLabels: z
      .union([z.literal(false), rawScatterDataLabelOverrideSchema])
      .optional(),
    marker: rawChartMarkerOptionSchema.optional(),
    points: z
      .record(z.string(), rawScatterChartPointSchema)
      .readonly()
      .optional(),
    trendlines: z.array(rawChartTrendlineSchema).readonly().optional(),
    errorBars: rawScatterErrorBarsSchema.optional(),
  })
  .readonly();

export type ScatterChartSeries = {
  readonly name: string;
  readonly x: readonly (number | null)[];
  readonly y: readonly (number | null)[];
  readonly color?: ColorValue | undefined;
  readonly line?: ChartLine | undefined;
  readonly dataLabels?: false | ScatterDataLabelOverride | undefined;
  readonly marker?: boolean | ChartMarker | undefined;
  readonly points?: Readonly<Record<number, ScatterChartPoint>> | undefined;
  readonly trendlines?: readonly ChartTrendline[] | undefined;
  readonly errorBars?: ScatterErrorBars | undefined;
};

export const scatterChartSeriesSchema: z.ZodType<ScatterChartSeries> =
  ScatterChartSeriesRawSchema as unknown as z.ZodType<
    AssertExact<
      Omit<ScatterChartSeries, "points">,
      Omit<z.infer<typeof ScatterChartSeriesRawSchema>, "points">
    > extends true
      ? ScatterChartSeries
      : never
  >;

export type SharedCategoryData = {
  readonly kind: "shared-categories";
  readonly categories: readonly string[];
};
const rawSharedCategoryDataSchema = z
  .strictObject({
    kind: z.literal("shared-categories"),
    categories: rawChartCategoriesSchema,
  })
  .readonly();
export const sharedCategoryDataSchema: z.ZodType<SharedCategoryData> =
  rawSharedCategoryDataSchema satisfies z.ZodType<
    AssertExact<
      SharedCategoryData,
      z.infer<typeof rawSharedCategoryDataSchema>
    > extends true
      ? SharedCategoryData
      : never
  >;
export type ChartGroupData = SharedCategoryData;

export type IndependentCategoryData = {
  readonly kind: "independent-categories";
};
const rawIndependentCategoryDataSchema = z
  .strictObject({
    kind: z.literal("independent-categories"),
  })
  .readonly();
export const independentCategoryDataSchema: z.ZodType<IndependentCategoryData> =
  rawIndependentCategoryDataSchema satisfies z.ZodType<
    AssertExact<
      IndependentCategoryData,
      z.infer<typeof rawIndependentCategoryDataSchema>
    > extends true
      ? IndependentCategoryData
      : never
  >;

export type ScatterData = {
  readonly kind: "xy";
};
const rawScatterDataSchema = z
  .strictObject({
    kind: z.literal("xy"),
  })
  .readonly();
export const scatterDataSchema: z.ZodType<ScatterData> =
  rawScatterDataSchema satisfies z.ZodType<
    AssertExact<ScatterData, z.infer<typeof rawScatterDataSchema>> extends true
      ? ScatterData
      : never
  >;

export type ChartTickMark = "none" | "in" | "out" | "cross";
const rawChartTickMarkSchema = z.enum(["none", "in", "out", "cross"]);
export const chartTickMarkSchema: z.ZodType<ChartTickMark> =
  rawChartTickMarkSchema satisfies z.ZodType<
    AssertExact<
      ChartTickMark,
      z.infer<typeof rawChartTickMarkSchema>
    > extends true
      ? ChartTickMark
      : never
  >;

export type ChartTickLabelPosition = "nextTo" | "high" | "low" | "none";
const rawChartTickLabelPositionSchema = z.enum([
  "nextTo",
  "high",
  "low",
  "none",
]);
export const chartTickLabelPositionSchema: z.ZodType<ChartTickLabelPosition> =
  rawChartTickLabelPositionSchema satisfies z.ZodType<
    AssertExact<
      ChartTickLabelPosition,
      z.infer<typeof rawChartTickLabelPositionSchema>
    > extends true
      ? ChartTickLabelPosition
      : never
  >;

export type ChartAxisCrossing = "autoZero" | "max" | "min";
const rawChartAxisCrossingSchema = z.enum(["autoZero", "max", "min"]);
export const chartAxisCrossingSchema: z.ZodType<ChartAxisCrossing> =
  rawChartAxisCrossingSchema satisfies z.ZodType<
    AssertExact<
      ChartAxisCrossing,
      z.infer<typeof rawChartAxisCrossingSchema>
    > extends true
      ? ChartAxisCrossing
      : never
  >;

export type ChartAxisPosition = "bottom" | "top" | "left" | "right";
const rawChartAxisPositionSchema = z.enum(["bottom", "top", "left", "right"]);
export const chartAxisPositionSchema: z.ZodType<ChartAxisPosition> =
  rawChartAxisPositionSchema satisfies z.ZodType<
    AssertExact<
      ChartAxisPosition,
      z.infer<typeof rawChartAxisPositionSchema>
    > extends true
      ? ChartAxisPosition
      : never
  >;

export type ChartDisplayBlanksAs = "gap" | "zero" | "span";
const rawChartDisplayBlanksAsSchema = z.enum(["gap", "zero", "span"]);
export const chartDisplayBlanksAsSchema: z.ZodType<ChartDisplayBlanksAs> =
  rawChartDisplayBlanksAsSchema satisfies z.ZodType<
    AssertExact<
      ChartDisplayBlanksAs,
      z.infer<typeof rawChartDisplayBlanksAsSchema>
    > extends true
      ? ChartDisplayBlanksAs
      : never
  >;

const axisBaseFields = {
  visible: z.boolean().optional(),
  gridlines: z.union([z.boolean(), rawChartLineSchema]).optional(),
  minorGridlines: z.union([z.boolean(), rawChartLineSchema]).optional(),
  line: z.union([z.boolean(), rawChartLineSchema]).optional(),
  style: rawChartTextStyleSchema.optional(),
  title: rawChartTitleSchema.optional(),
  majorTickMark: rawChartTickMarkSchema.optional(),
  minorTickMark: rawChartTickMarkSchema.optional(),
  tickLabelPosition: rawChartTickLabelPositionSchema.optional(),
  reversed: z.boolean().optional(),
  position: rawChartAxisPositionSchema.optional(),
  crosses: rawChartAxisCrossingSchema.optional(),
  crossesAt: z.number().finite().optional(),
};

export type ChartCategoryAxis = {
  readonly kind: "category";
  readonly visible?: boolean | undefined;
  readonly gridlines?: boolean | ChartLine | undefined;
  readonly minorGridlines?: boolean | ChartLine | undefined;
  readonly line?: boolean | ChartLine | undefined;
  readonly style?: ChartTextStyle | undefined;
  readonly title?: ChartTitle | undefined;
  readonly majorTickMark?: ChartTickMark | undefined;
  readonly minorTickMark?: ChartTickMark | undefined;
  readonly tickLabelPosition?: ChartTickLabelPosition | undefined;
  readonly reversed?: boolean | undefined;
  readonly position?: ChartAxisPosition | undefined;
  readonly crosses?: ChartAxisCrossing | undefined;
  readonly crossesAt?: number | undefined;
};
const rawChartCategoryAxisSchema = z
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
export const chartCategoryAxisSchema: z.ZodType<ChartCategoryAxis> =
  rawChartCategoryAxisSchema satisfies z.ZodType<
    AssertExact<
      ChartCategoryAxis,
      z.infer<typeof rawChartCategoryAxisSchema>
    > extends true
      ? ChartCategoryAxis
      : never
  >;

export type ChartValueAxis = {
  readonly kind: "value";
  readonly visible?: boolean | undefined;
  readonly gridlines?: boolean | ChartLine | undefined;
  readonly minorGridlines?: boolean | ChartLine | undefined;
  readonly line?: boolean | ChartLine | undefined;
  readonly style?: ChartTextStyle | undefined;
  readonly title?: ChartTitle | undefined;
  readonly majorTickMark?: ChartTickMark | undefined;
  readonly minorTickMark?: ChartTickMark | undefined;
  readonly tickLabelPosition?: ChartTickLabelPosition | undefined;
  readonly reversed?: boolean | undefined;
  readonly position?: ChartAxisPosition | undefined;
  readonly crosses?: ChartAxisCrossing | undefined;
  readonly crossesAt?: number | undefined;
  readonly min?: number | undefined;
  readonly max?: number | undefined;
  readonly majorUnit?: number | undefined;
  readonly minorUnit?: number | undefined;
  readonly numberFormat?: string | undefined;
  readonly logBase?: number | undefined;
};
const rawChartValueAxisSchema = z
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
export const chartValueAxisSchema: z.ZodType<ChartValueAxis> =
  rawChartValueAxisSchema satisfies z.ZodType<
    AssertExact<
      ChartValueAxis,
      z.infer<typeof rawChartValueAxisSchema>
    > extends true
      ? ChartValueAxis
      : never
  >;

export type ChartAxis = ChartCategoryAxis | ChartValueAxis;
const rawChartAxisSchema = z.discriminatedUnion("kind", [
  rawChartCategoryAxisSchema,
  rawChartValueAxisSchema,
]);
export const chartAxisSchema: z.ZodType<ChartAxis> =
  rawChartAxisSchema satisfies z.ZodType<
    AssertExact<ChartAxis, z.infer<typeof rawChartAxisSchema>> extends true
      ? ChartAxis
      : never
  >;

export type ChartAxes = Readonly<Record<string, ChartAxis>>;
const rawChartAxesSchema = z
  .record(z.string(), rawChartAxisSchema)
  .readonly()
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
export const chartAxesSchema: z.ZodType<ChartAxes> =
  rawChartAxesSchema satisfies z.ZodType<
    AssertExact<ChartAxes, z.infer<typeof rawChartAxesSchema>> extends true
      ? ChartAxes
      : never
  >;

export type ChartCategoryAxisReferences = {
  readonly category: string;
  readonly value: string;
};
const rawChartCategoryAxisReferencesSchema = z
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
export const chartCategoryAxisReferencesSchema: z.ZodType<ChartCategoryAxisReferences> =
  rawChartCategoryAxisReferencesSchema satisfies z.ZodType<
    AssertExact<
      ChartCategoryAxisReferences,
      z.infer<typeof rawChartCategoryAxisReferencesSchema>
    > extends true
      ? ChartCategoryAxisReferences
      : never
  >;

export type ChartXYAxisReferences = {
  readonly x: string;
  readonly y: string;
};
const rawChartXYAxisReferencesSchema = z
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
export const chartXYAxisReferencesSchema: z.ZodType<ChartXYAxisReferences> =
  rawChartXYAxisReferencesSchema satisfies z.ZodType<
    AssertExact<
      ChartXYAxisReferences,
      z.infer<typeof rawChartXYAxisReferencesSchema>
    > extends true
      ? ChartXYAxisReferences
      : never
  >;

export type ChartAxisReferences =
  | ChartCategoryAxisReferences
  | ChartXYAxisReferences;
const rawChartAxisReferencesSchema = z.union([
  rawChartCategoryAxisReferencesSchema,
  rawChartXYAxisReferencesSchema,
]);
export const chartAxisReferencesSchema: z.ZodType<ChartAxisReferences> =
  rawChartAxisReferencesSchema satisfies z.ZodType<
    AssertExact<
      ChartAxisReferences,
      z.infer<typeof rawChartAxisReferencesSchema>
    > extends true
      ? ChartAxisReferences
      : never
  >;

export type ScatterLine = false | "straight" | "smooth";
const rawScatterLineSchema = z.union([
  z.literal(false),
  z.literal("straight"),
  z.literal("smooth"),
]);
export const scatterLineSchema: z.ZodType<ScatterLine> =
  rawScatterLineSchema satisfies z.ZodType<
    AssertExact<ScatterLine, z.infer<typeof rawScatterLineSchema>> extends true
      ? ScatterLine
      : never
  >;

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
        const parsedMarker = chartMarkerOptionSchema.safeParse(item.marker);
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
            const parsedMarker = chartMarkerOptionSchema.safeParse(
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
  data: rawSharedCategoryDataSchema,
  series: z.array(chartSeriesSchema).min(1).readonly(),
  axes: rawChartCategoryAxisReferencesSchema.optional(),
  gapWidth: z.number().finite().min(0).max(500).optional(),
  overlap: z.number().finite().min(-100).max(100).optional(),
};

const ColumnOrBarUnstackedGroupSchema = z
  .strictObject({
    ...columnOrBarCommonFields,
    stacking: z.undefined().optional(),
    dataLabels: z
      .union([z.literal(false), rawChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly();

const ColumnOrBarStackedGroupSchema = z
  .strictObject({
    ...columnOrBarCommonFields,
    stacking: z.enum(["stacked", "percent"]),
    dataLabels: z
      .union([z.literal(false), rawChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly();

export type ColumnOrBarChartGroup =
  | {
      readonly kind: "column" | "bar";
      readonly data: SharedCategoryData;
      readonly series: readonly ChartSeries[];
      readonly axes?: ChartCategoryAxisReferences | undefined;
      readonly gapWidth?: number | undefined;
      readonly overlap?: number | undefined;
      readonly stacking?: undefined;
      readonly dataLabels?:
        | false
        | ChartDataLabels<"center" | "insideEnd" | "insideBase" | "outsideEnd">
        | undefined;
    }
  | {
      readonly kind: "column" | "bar";
      readonly data: SharedCategoryData;
      readonly series: readonly ChartSeries[];
      readonly axes?: ChartCategoryAxisReferences | undefined;
      readonly gapWidth?: number | undefined;
      readonly overlap?: number | undefined;
      readonly stacking: "stacked" | "percent";
      readonly dataLabels?:
        | false
        | ChartDataLabels<"center" | "insideEnd" | "insideBase">
        | undefined;
    };

export const columnOrBarChartGroupSchema = z
  .union([ColumnOrBarUnstackedGroupSchema, ColumnOrBarStackedGroupSchema])
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  }) as unknown as z.ZodType<ColumnOrBarChartGroup>;

const LineChartGroupRawSchema = z
  .strictObject({
    kind: z.literal("line"),
    data: rawSharedCategoryDataSchema,
    series: z.array(chartSeriesSchema).min(1).readonly(),
    axes: rawChartCategoryAxisReferencesSchema.optional(),
    stacking: z.enum(["stacked", "percent"]).optional(),
    dataLabels: z
      .union([z.literal(false), rawChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly()
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  });

export type LineChartGroup = {
  readonly kind: "line";
  readonly data: SharedCategoryData;
  readonly series: readonly ChartSeries[];
  readonly axes?: ChartCategoryAxisReferences | undefined;
  readonly stacking?: "stacked" | "percent" | undefined;
  readonly dataLabels?:
    | false
    | ChartDataLabels<"left" | "right" | "bottom" | "top" | "center">
    | undefined;
};
export const lineChartGroupSchema =
  LineChartGroupRawSchema as unknown as z.ZodType<LineChartGroup>;

const AreaChartGroupRawSchema = z
  .strictObject({
    kind: z.literal("area"),
    data: rawSharedCategoryDataSchema,
    series: z.array(chartSeriesSchema).min(1).readonly(),
    axes: rawChartCategoryAxisReferencesSchema.optional(),
    stacking: z.enum(["stacked", "percent"]).optional(),
    dataLabels: z
      .union([z.literal(false), rawChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly()
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  });

export type AreaChartGroup = {
  readonly kind: "area";
  readonly data: SharedCategoryData;
  readonly series: readonly ChartSeries[];
  readonly axes?: ChartCategoryAxisReferences | undefined;
  readonly stacking?: "stacked" | "percent" | undefined;
  readonly dataLabels?:
    | false
    | (Omit<ChartDataLabels, "position"> & { readonly position?: never })
    | undefined;
};
export const areaChartGroupSchema =
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
    data: rawSharedCategoryDataSchema,
    series: z.array(chartSeriesSchema).min(1).readonly(),
    ...pieOptionsFields,
    dataLabels: z
      .union([z.literal(false), rawChartDataLabelsBaseSchema])
      .optional(),
  })
  .readonly()
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  });

export type PieChartGroup = {
  readonly kind: "pie";
  readonly data: SharedCategoryData;
  readonly series: readonly ChartSeries[];
  readonly explosion?: number | undefined;
  readonly firstSliceAngle?: number | undefined;
  readonly dataLabels?:
    | false
    | ChartDataLabels<"center" | "insideEnd" | "outsideEnd" | "bestFit">
    | undefined;
};
export const pieChartGroupSchema =
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
  dataLabels: z
    .union([z.literal(false), rawChartDataLabelsBaseSchema])
    .optional(),
};

const DoughnutSharedGroupRawSchema = z
  .strictObject({
    ...doughnutOptionsFields,
    data: rawSharedCategoryDataSchema,
    series: z.array(chartSeriesSchema).min(1).readonly(),
  })
  .readonly();

const DoughnutIndependentGroupRawSchema = z
  .strictObject({
    ...doughnutOptionsFields,
    data: rawIndependentCategoryDataSchema,
    series: z.array(independentCategorySeriesSchema).min(1).readonly(),
  })
  .readonly();

export type DoughnutChartGroup =
  | {
      readonly kind: "doughnut";
      readonly explosion?: number | undefined;
      readonly firstSliceAngle?: number | undefined;
      readonly holeSize?: number | undefined;
      readonly data: SharedCategoryData;
      readonly series: readonly ChartSeries[];
      readonly dataLabels?:
        | false
        | (Omit<ChartDataLabels, "position"> & { readonly position?: never })
        | undefined;
    }
  | {
      readonly kind: "doughnut";
      readonly explosion?: number | undefined;
      readonly firstSliceAngle?: number | undefined;
      readonly holeSize?: number | undefined;
      readonly data: IndependentCategoryData;
      readonly series: readonly IndependentCategorySeries[];
      readonly dataLabels?:
        | false
        | (Omit<ChartDataLabels, "position"> & { readonly position?: never })
        | undefined;
    };

export const doughnutChartGroupSchema = z
  .union([DoughnutSharedGroupRawSchema, DoughnutIndependentGroupRawSchema])
  .superRefine((group, ctx) => {
    validateCategoryGroupSeries(group, ctx);
  }) as unknown as z.ZodType<DoughnutChartGroup>;

const rawScatterChartGroupSchema = z
  .strictObject({
    kind: z.literal("scatter"),
    data: rawScatterDataSchema,
    series: z.array(scatterChartSeriesSchema).min(1).readonly(),
    axes: rawChartXYAxisReferencesSchema.optional(),
    line: rawScatterLineSchema.optional(),
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

export type ScatterChartGroup = {
  readonly kind: "scatter";
  readonly data: ScatterData;
  readonly series: readonly ScatterChartSeries[];
  readonly axes?: ChartXYAxisReferences | undefined;
  readonly line?: ScatterLine | undefined;
  readonly markers?: boolean | undefined;
  readonly dataLabels?: false | ScatterDataLabels | undefined;
};

export const scatterChartGroupSchema: z.ZodType<ScatterChartGroup> =
  rawScatterChartGroupSchema satisfies z.ZodType<
    AssertExact<
      ScatterChartGroup,
      z.infer<typeof rawScatterChartGroupSchema>
    > extends true
      ? ScatterChartGroup
      : never
  >;

export type ChartGroup =
  | ColumnOrBarChartGroup
  | LineOrAreaChartGroup
  | PieChartGroup
  | DoughnutChartGroup
  | ScatterChartGroup;

export const chartGroupSchema: z.ZodType<ChartGroup> = z.union([
  columnOrBarChartGroupSchema,
  lineChartGroupSchema,
  areaChartGroupSchema,
  pieChartGroupSchema,
  doughnutChartGroupSchema,
  rawScatterChartGroupSchema,
]);

export type ChartLegend =
  | false
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "topRight";
const rawChartLegendSchema = z.union([
  z.literal(false),
  z.enum(["top", "bottom", "left", "right", "topRight"]),
]);
export const chartLegendSchema: z.ZodType<ChartLegend> =
  rawChartLegendSchema satisfies z.ZodType<
    AssertExact<ChartLegend, z.infer<typeof rawChartLegendSchema>> extends true
      ? ChartLegend
      : never
  >;

/**
 * Explicit plot-area rectangle inside the chart box, in pixels (`px`) relative to the chart box's
 * top-left corner (`{ left, top, width, height }`, not `0..1` fractions).
 */
export type ChartPlotArea = {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
};

const rawChartPlotAreaSchema = z
  .strictObject({
    left: z.number().finite().nonnegative(),
    top: z.number().finite().nonnegative(),
    width: z.number().finite().positive(),
    height: z.number().finite().positive(),
  })
  .readonly();
export const chartPlotAreaSchema: z.ZodType<ChartPlotArea> =
  rawChartPlotAreaSchema satisfies z.ZodType<
    AssertExact<
      ChartPlotArea,
      z.infer<typeof rawChartPlotAreaSchema>
    > extends true
      ? ChartPlotArea
      : never
  >;

export const chartCoordinateSchema = z
  .number()
  .finite()
  .min(-1_000_000)
  .max(1_000_000);

export const chartExtentSchema = z
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

export type ChartProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly name?: string | undefined;
  readonly placeholder?: string | undefined;
  readonly nativePlaceholder?: NativePlaceholderRef | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  readonly axes?: ChartAxes | undefined;
  /**
   * Non-empty list of chart series groups; each group shares one `kind`, group options, `data` mode, and `series`.
   *
   * @example
   * ```ts
   * groups: [
   *   {
   *     kind: "column",
   *     data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
   *     series: [{ name: "Revenue", values: [120, 150] }],
   *   },
   * ]
   * ```
   */
  readonly groups: readonly ChartGroup[];
  readonly title?: ChartTitle | undefined;
  readonly displayBlanksAs?: ChartDisplayBlanksAs | undefined;
  readonly legend?: ChartLegend | undefined;
  /**
   * Explicit plot-area rectangle `{ left, top, width, height }` in pixels (`px`) relative to the
   * chart box's top-left corner (not `0..1` fractions). When omitted, automatic margins are computed.
   */
  readonly plotArea?: ChartPlotArea | undefined;
  readonly style?: ChartStyle | undefined;
};
export type ChartBase = ChartProps;

const chartPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: xmlStringSchema("Chart.name").optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  axes: chartAxesSchema.optional(),
  groups: z.array(chartGroupSchema).min(1).readonly(),
  title: rawChartTitleSchema.optional(),
  displayBlanksAs: rawChartDisplayBlanksAsSchema.optional(),
  legend: rawChartLegendSchema.optional(),
  plotArea: rawChartPlotAreaSchema.optional(),
  style: rawChartStyleSchema.optional(),
  ...motionPropsShape,
});

const rawChartPropsSchema = chartPropsObjectSchema.readonly();
export const chartPropsSchema: z.ZodType<ChartProps> =
  rawChartPropsSchema satisfies z.ZodType<
    AssertExact<ChartProps, z.infer<typeof rawChartPropsSchema>> extends true
      ? ChartProps
      : never
  >;

export const chartAllowedPropKeys: readonly string[] = Object.keys(
  chartPropsObjectSchema.shape,
);

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
    expected === "category" ? chartCategoryAxisSchema : chartValueAxisSchema;
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

const commonGroupKeys = new Set([
  "kind",
  "data",
  "series",
  "axes",
  "dataLabels",
]);

function groupOptionKeysFromShape(
  shape: Readonly<Record<string, unknown>>,
): readonly string[] {
  return Object.keys(shape).filter((key) => !commonGroupKeys.has(key));
}

const groupOptionsByKind: Readonly<Record<ChartKind, readonly string[]>> = {
  column: groupOptionKeysFromShape(
    ColumnOrBarStackedGroupSchema.unwrap().shape,
  ),
  bar: groupOptionKeysFromShape(ColumnOrBarStackedGroupSchema.unwrap().shape),
  line: groupOptionKeysFromShape(LineChartGroupRawSchema.unwrap().shape),
  area: groupOptionKeysFromShape(AreaChartGroupRawSchema.unwrap().shape),
  pie: groupOptionKeysFromShape(PieChartGroupRawSchema.unwrap().shape),
  doughnut: groupOptionKeysFromShape(
    DoughnutSharedGroupRawSchema.unwrap().shape,
  ),
  scatter: groupOptionKeysFromShape(rawScatterChartGroupSchema.unwrap().shape),
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
  const kindParsed = chartKindSchema.safeParse(value["kind"]);
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
    const parsed = scatterChartGroupSchema.safeParse(value);
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
    const parsed = lineChartGroupSchema.safeParse(value);
    if (!parsed.success) {
      throwChartZodError(parsed.error, "invalid-data", label);
    }
    return parsed.data;
  }
  if (kind === "area") {
    const parsed = areaChartGroupSchema.safeParse(value);
    if (!parsed.success) {
      throwChartZodError(parsed.error, "invalid-data", label);
    }
    return parsed.data;
  }
  if (kind === "pie") {
    const parsed = pieChartGroupSchema.safeParse(value);
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

/** Validates and parses `Chart.title` using `chartTitleSchema`. */
export function parseChartTitleBySchema(
  value: unknown,
  label: string,
): ChartTitle | undefined {
  if (value === undefined) return undefined;
  const parsed = chartTitleSchema.safeParse(value);
  if (!parsed.success) {
    throwChartZodError(parsed.error, "invalid-parameter", label);
  }
  return parsed.data;
}

/** Validates and parses `Chart.plotArea` using `chartPlotAreaSchema` + chart box bounds. */
export function parseChartPlotAreaBySchema(
  value: unknown,
  box?: { readonly width: number; readonly height: number },
): ChartPlotArea | undefined {
  if (value === undefined) return undefined;
  const parsed = chartPlotAreaSchema.safeParse(value);
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

/** Validates and parses `Chart.legend` using `chartLegendSchema`. */
export function parseChartLegendBySchema(
  value: unknown,
): ChartLegend | undefined {
  if (value === undefined) return undefined;
  const parsed = chartLegendSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-legend",
      "Chart.legend must be false, top, bottom, left, right, or topRight",
    );
  }
  return parsed.data;
}

/** Validates and parses a chart coordinate (`x` or `y`) via `chartCoordinateSchema`. */
export function parseChartCoordinateBySchema(
  value: unknown,
  label: "x" | "y",
): number {
  if (value === undefined) return 0;
  const parsed = chartCoordinateSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${label} must be a finite pixel value between -1000000 and 1000000`,
    );
  }
  return parsed.data;
}

/** Validates and parses a chart extent (`width` or `height`) via `chartExtentSchema`. */
export function parseChartExtentBySchema(
  value: unknown,
  label: "width" | "height",
): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    if (value === 0 || pxToEmu(value) === 0) {
      throw new CapabilityError(
        "invalid-coordinate",
        `Computed ${label} is 0px; ${label} must be greater than 0px`,
      );
    }
    if (value < 0 || pxToEmu(value) < 1) {
      throw new CapabilityError(
        "invalid-coordinate",
        `Computed ${label} is ${value}px; ${label} must be greater than 0px`,
      );
    }
  } else {
    throw new CapabilityError(
      "invalid-coordinate",
      `${label} must be a finite positive pixel value`,
    );
  }
  const parsed = chartExtentSchema.safeParse(value);
  if (!parsed.success) {
    throwChartZodError(parsed.error, "invalid-coordinate", `Chart.${label}`);
  }
  return parsed.data;
}

/** Validates and parses `Chart.displayBlanksAs` via `chartDisplayBlanksAsSchema`. */
export function parseChartDisplayBlanksAsBySchema(
  value: unknown,
): ChartDisplayBlanksAs | undefined {
  if (value === undefined) return undefined;
  const parsed = chartDisplayBlanksAsSchema.safeParse(value);
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
