import type { ChartDataLabelPosition, ChartDataLabels } from "./chart-types.ts";

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

/**
 * Merges a more local data-label declaration over its inherited settings.
 *
 * An explicit `false` disables labels at that level. Omitted fields keep the inherited value;
 * this is a field-level merge rather than a whole-object replacement.
 */
export function mergeDataLabels(
  base: false | ChartDataLabels | undefined,
  override: false | ChartDataLabels | undefined,
): false | ChartDataLabels | undefined {
  if (override === undefined) return base;
  if (override === false) return false;
  if (base === false || base === undefined) return override;
  return { ...base, ...override };
}
