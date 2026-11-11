import type { ChartDataLabels } from "./chart-types.ts";

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
