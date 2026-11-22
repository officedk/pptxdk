import type {
  ChartAxisPosition,
  ChartDataLabelPosition,
  ChartDataLabelSettings,
  ChartErrorBars,
  ChartLegend,
  ChartLineDash,
  ChartTrendlineType,
} from "./chart-types.ts";

function invertRecord<K extends string, V extends string>(
  record: Readonly<Record<K, V>>,
): Readonly<Record<V, K>> {
  return Object.fromEntries(
    Object.entries(record).map(([key, value]) => [value, key]),
  ) as Readonly<Record<V, K>>;
}

/** Maps semantic label positions to native c:dLblPos values. */
export const chartDataLabelPositions: Readonly<
  Record<ChartDataLabelPosition, string>
> = {
  center: "ctr",
  insideEnd: "inEnd",
  insideBase: "inBase",
  outsideEnd: "outEnd",
  bestFit: "bestFit",
  left: "l",
  right: "r",
  bottom: "b",
  top: "t",
};

/** Maps native c:dLblPos values back to semantic label positions. */
export const chartDataLabelPositionFromNative: Readonly<
  Record<string, ChartDataLabelPosition>
> = invertRecord(chartDataLabelPositions);

/** Maps semantic line dash presets to DrawingML a:prstDash values. */
export const chartLineDashValues: Readonly<Record<ChartLineDash, string>> = {
  solid: "solid",
  dash: "dash",
  dot: "dot",
  dashDot: "dashDot",
  longDash: "lgDash",
  longDashDot: "lgDashDot",
  longDashDotDot: "lgDashDotDot",
};

/** Maps DrawingML a:prstDash values back to semantic line dash presets. */
export const chartLineDashFromNative: Readonly<Record<string, ChartLineDash>> =
  invertRecord(chartLineDashValues);

/** Maps semantic axis positions to native c:axPos values. */
export const chartAxisPositionCodes: Readonly<
  Record<ChartAxisPosition, "b" | "t" | "l" | "r">
> = {
  bottom: "b",
  top: "t",
  left: "l",
  right: "r",
};

/** Maps native c:axPos values back to semantic axis positions. */
export const chartAxisPositionFromNative: Readonly<
  Record<string, ChartAxisPosition>
> = invertRecord(chartAxisPositionCodes);

/** Maps semantic legend positions to native c:legendPos values. */
export const chartLegendPositions: Readonly<
  Record<Exclude<ChartLegend, false>, "t" | "b" | "l" | "r" | "tr">
> = {
  top: "t",
  bottom: "b",
  left: "l",
  right: "r",
  topRight: "tr",
};

/** Maps native c:legendPos values back to semantic legend positions. */
export const chartLegendPositionFromNative: Readonly<
  Record<string, Exclude<ChartLegend, false>>
> = invertRecord(chartLegendPositions);

/** Maps semantic trendline types to native c:trendlineType values. */
export const chartTrendlineTypeValues: Readonly<
  Record<ChartTrendlineType, string>
> = {
  linear: "linear",
  exponential: "exp",
  logarithmic: "log",
  polynomial: "poly",
  power: "power",
  movingAverage: "movingAvg",
};

/** Maps native c:trendlineType values back to semantic trendline types. */
export const chartTrendlineTypeFromNative: Readonly<
  Record<string, ChartTrendlineType>
> = invertRecord(chartTrendlineTypeValues);

/** Maps semantic error-bar modes to native c:errValType values. */
export const chartErrorBarValueTypes: Readonly<
  Record<ChartErrorBars["type"], string>
> = {
  custom: "cust",
  fixed: "fixedVal",
  percentage: "percentage",
  standardDeviation: "stdDev",
  standardError: "stdErr",
};

/** Maps native c:errValType values back to semantic error-bar modes. */
export const chartErrorBarValueTypeFromNative: Readonly<
  Record<string, ChartErrorBars["type"]>
> = invertRecord(chartErrorBarValueTypes);

/**
 * Merges a more local data-label declaration over its inherited settings.
 *
 * An explicit `false` disables labels at that level. Omitted fields keep the inherited value;
 * this is a field-level merge rather than a whole-object replacement.
 */
export function mergeDataLabels(
  base: false | ChartDataLabelSettings | undefined,
  override: false | ChartDataLabelSettings | undefined,
): false | ChartDataLabelSettings | undefined {
  if (override === undefined) return base;
  if (override === false) return false;
  if (base === false || base === undefined) return override;
  return { ...base, ...override };
}
