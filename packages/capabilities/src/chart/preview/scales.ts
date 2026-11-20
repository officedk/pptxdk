import type { GeometryBox } from "@react-pptx/core/capability";

import type {
  ParsedCategoryAxis,
  ParsedChart,
  ParsedChartAxis,
  ParsedChartGroup,
  ParsedValueAxis,
} from "../chart-model.ts";
import {
  assertNever,
  type AxisScale,
  type CategoryAxisScale,
  type ValueAxisScale,
} from "./types.ts";

function cleanNumber(value: number): number {
  if (Object.is(value, -0) || value === 0) return 0;
  return Number(value.toPrecision(12));
}

function niceStep(span: number): number {
  const safeSpan = span > 0 && Number.isFinite(span) ? span : 1;
  const rough = safeSpan / 5;
  const mag = 10 ** Math.floor(Math.log10(rough));
  const candidate =
    [1, 2, 2.5, 5, 10].find((c) => c * mag >= rough - 1e-12) ?? 10;
  return cleanNumber(candidate * mag);
}

function generateLinearTicks(
  min: number,
  max: number,
  step: number,
  maxCount = 50,
): number[] {
  if (!Number.isFinite(step) || step <= 0 || max <= min) {
    return [cleanNumber(min), cleanNumber(max)];
  }
  const count = Math.floor((max - min) / step + 1e-9) + 1;
  const effectiveStep =
    count > maxCount ? cleanNumber((max - min) / (maxCount - 1)) : step;
  const ticks: number[] = [];
  const limit = max + effectiveStep * 1e-6;
  for (let i = 0; i <= maxCount + 1; i += 1) {
    const val = cleanNumber(min + i * effectiveStep);
    if (val > limit) break;
    ticks.push(val);
  }
  if (ticks.length === 0) {
    ticks.push(cleanNumber(min), cleanNumber(max));
  }
  return ticks;
}

function collectAxisValues(
  axis: ParsedValueAxis,
  groups: readonly ParsedChartGroup[],
): {
  readonly values: number[];
  readonly hasPercent: boolean;
  readonly hasBaselineZero: boolean;
} {
  const values: number[] = [];
  let hasPercent = false;
  let hasBaselineZero = false;

  for (const group of groups) {
    switch (group.kind) {
      case "column":
      case "bar":
      case "line":
      case "area": {
        if (group.axes.value !== axis) continue;
        if (
          group.kind === "column" ||
          group.kind === "bar" ||
          group.kind === "area"
        ) {
          hasBaselineZero = true;
        }
        const catCount = group.data.categories.length;
        if (group.stacking === "percent") {
          hasPercent = true;
          for (let i = 0; i < catCount; i += 1) {
            let posSum = 0;
            let negAbsSum = 0;
            for (const series of group.series) {
              const v = series.values[i];
              if (v === null || v === undefined || !Number.isFinite(v)) {
                continue;
              }
              if (v >= 0) posSum += v;
              else negAbsSum += -v;
            }
            const totalAbs = posSum + negAbsSum;
            if (totalAbs > 0) {
              values.push(posSum / totalAbs, -negAbsSum / totalAbs);
            }
          }
        } else if (group.stacking === "stacked") {
          for (let i = 0; i < catCount; i += 1) {
            let posSum = 0;
            let negSum = 0;
            let hasAny = false;
            for (const series of group.series) {
              const v = series.values[i];
              if (v === null || v === undefined || !Number.isFinite(v)) {
                continue;
              }
              hasAny = true;
              if (v >= 0) posSum += v;
              else negSum += v;
            }
            if (hasAny) values.push(posSum, negSum);
          }
        } else {
          for (const series of group.series) {
            for (const v of series.values) {
              if (v !== null && v !== undefined && Number.isFinite(v)) {
                values.push(v);
              }
            }
          }
        }
        break;
      }
      case "scatter": {
        const isX = group.axes.x === axis;
        const isY = group.axes.y === axis;
        if (!isX && !isY) continue;
        for (const series of group.series) {
          const arr = isX ? series.x : series.y;
          for (const v of arr) {
            if (v !== null && v !== undefined && Number.isFinite(v)) {
              values.push(v);
            }
          }
        }
        break;
      }
      case "pie":
      case "doughnut":
        break;
      default:
        assertNever(group);
    }
  }

  return { values, hasPercent, hasBaselineZero };
}

function buildCategoryAxisScale(
  axis: ParsedCategoryAxis,
  groups: readonly ParsedChartGroup[],
): CategoryAxisScale {
  let categories: readonly string[] = [];
  let hasCategoryGroup = false;
  let allArea = true;

  for (const group of groups) {
    if (
      group.kind === "scatter" ||
      group.kind === "pie" ||
      group.kind === "doughnut"
    ) {
      continue;
    }
    if (group.axes.category !== axis) continue;
    if (!hasCategoryGroup) {
      categories = group.data.categories;
    }
    hasCategoryGroup = true;
    if (group.kind !== "area") allArea = false;
  }

  const count = Math.max(1, categories.length);
  const crossBetween: "between" | "midCat" =
    hasCategoryGroup && allArea && count > 1 ? "midCat" : "between";
  const reversed = axis.spec.reversed === true;
  const horizontal = axis.direction === "horizontal";

  const mapIndex = (index: number): number => {
    const rawT =
      crossBetween === "midCat" && count > 1
        ? index / (count - 1)
        : (index + 0.5) / count;
    return reversed ? 1 - rawT : rawT;
  };

  const mapBoundary = (boundary: number): number => {
    const rawT = boundary / count;
    return reversed ? 1 - rawT : rawT;
  };

  const mapScreen = (t: number, plot: GeometryBox): number =>
    horizontal ? plot.x + t * plot.width : plot.y + (1 - t) * plot.height;

  return {
    kind: "category",
    axis,
    categories,
    count,
    crossBetween,
    reversed,
    mapIndex,
    mapBoundary,
    mapScreen,
    indexToScreen(index, plot) {
      return mapScreen(mapIndex(index), plot);
    },
    slotSpan(plot) {
      return (horizontal ? plot.width : plot.height) / count;
    },
  };
}

function buildValueAxisScale(
  axis: ParsedValueAxis,
  groups: readonly ParsedChartGroup[],
): ValueAxisScale {
  const { values, hasPercent, hasBaselineZero } = collectAxisValues(
    axis,
    groups,
  );
  const spec = axis.spec;
  const reversed = spec.reversed === true;
  const horizontal = axis.direction === "horizontal";
  const numberFormat = spec.numberFormat ?? (hasPercent ? "0%" : undefined);

  if (spec.logBase !== undefined) {
    const base = spec.logBase;
    const positive = values.filter((v) => v > 0);
    const rawMin = positive.length > 0 ? Math.min(...positive) : 1;
    const rawMax = positive.length > 0 ? Math.max(...positive) : base * base;
    const minExp = Math.floor(Math.log(rawMin) / Math.log(base) + 1e-12);
    const maxExp = Math.ceil(Math.log(rawMax) / Math.log(base) - 1e-12);
    let min =
      spec.min !== undefined && spec.min > 0
        ? spec.min
        : cleanNumber(base ** minExp);
    if (min <= 0 || !Number.isFinite(min)) min = 1;
    let max =
      spec.max !== undefined && spec.max > min
        ? spec.max
        : cleanNumber(base ** maxExp);
    if (max <= min || !Number.isFinite(max)) max = cleanNumber(min * base);

    const startExp = Math.floor(Math.log(min) / Math.log(base) + 1e-9);
    const endExp = Math.ceil(Math.log(max) / Math.log(base) - 1e-9);
    const spanExp = Math.max(1, endExp - startExp + 1);
    const stride = spanExp > 10 ? Math.ceil(spanExp / 8) : 1;
    const majorTicks: number[] = [];
    const minorTicks: number[] = [];
    for (let e = startExp; e <= endExp; e += 1) {
      const tick = cleanNumber(base ** e);
      if ((e - startExp) % stride === 0) {
        if (tick >= min * (1 - 1e-9) && tick <= max * (1 + 1e-9)) {
          majorTicks.push(tick);
        }
      }
      const subMax = Math.min(9, Math.floor(base) - 1);
      for (let m = 2; m <= subMax; m += 1) {
        const sub = cleanNumber(m * base ** e);
        if (sub > min * (1 + 1e-9) && sub < max * (1 - 1e-9)) {
          minorTicks.push(sub);
        }
      }
    }
    if (majorTicks.length < 2) {
      majorTicks.splice(0, majorTicks.length, min, max);
    }

    const logSpan = Math.log(max / min);
    const mapValue = (v: number): number => {
      if (v <= 0 || !Number.isFinite(v)) return reversed ? 1 : 0;
      const rawT = Math.log(v / min) / logSpan;
      return reversed ? 1 - rawT : rawT;
    };
    const mapScreen = (t: number, plot: GeometryBox): number =>
      horizontal ? plot.x + t * plot.width : plot.y + (1 - t) * plot.height;

    return {
      kind: "value",
      axis,
      min,
      max,
      logBase: base,
      majorTicks,
      minorTicks,
      reversed,
      numberFormat,
      mapValue,
      mapScreen,
      valueToScreen(v, plot) {
        return mapScreen(mapValue(v), plot);
      },
      zeroValue() {
        return min;
      },
      zeroScreen(plot) {
        return mapScreen(mapValue(min), plot);
      },
    };
  }

  let rawMin = values.length > 0 ? Math.min(...values) : 0;
  let rawMax = values.length > 0 ? Math.max(...values) : 1;
  if (hasBaselineZero || values.length === 0) {
    rawMin = Math.min(0, rawMin);
    rawMax = Math.max(0, rawMax);
  }
  if (rawMin === rawMax) {
    if (rawMax === 0) {
      rawMin = 0;
      rawMax = 1;
    } else if (rawMax > 0) {
      rawMin = 0;
    } else {
      rawMax = 0;
    }
  }

  let majorUnit: number;
  let min: number;
  let max: number;

  if (hasPercent && spec.min === undefined && spec.max === undefined) {
    majorUnit = spec.majorUnit ?? 0.2;
    min =
      rawMin >= 0
        ? 0
        : Math.max(
            -1,
            cleanNumber(Math.floor(rawMin / majorUnit + 1e-12) * majorUnit),
          );
    max =
      rawMax <= 1
        ? 1
        : Math.min(
            1,
            cleanNumber(Math.ceil(rawMax / majorUnit - 1e-12) * majorUnit),
          );
    if (min >= max) max = cleanNumber(min + majorUnit);
  } else {
    const initialSpan = (spec.max ?? rawMax) - (spec.min ?? rawMin);
    majorUnit = spec.majorUnit ?? niceStep(initialSpan);
    min =
      spec.min ??
      cleanNumber(Math.floor(rawMin / majorUnit + 1e-12) * majorUnit);
    max =
      spec.max ??
      cleanNumber(Math.ceil(rawMax / majorUnit - 1e-12) * majorUnit);
    if (min >= max) max = cleanNumber(min + majorUnit);
    if (
      spec.majorUnit === undefined &&
      (spec.min !== undefined || spec.max !== undefined)
    ) {
      const tickEstimate = (max - min) / majorUnit;
      if (tickEstimate < 2 || tickEstimate > 12) {
        majorUnit = niceStep(max - min);
      }
    }
  }

  const minorUnit = spec.minorUnit ?? cleanNumber(majorUnit / 5);
  const majorTicks = generateLinearTicks(min, max, majorUnit, 40);
  const minorTicks =
    minorUnit > 0
      ? generateLinearTicks(min, max, minorUnit, 200).filter(
          (t) => !majorTicks.some((m) => Math.abs(m - t) <= majorUnit * 1e-6),
        )
      : [];

  const span = max - min;
  const mapValue = (v: number): number => {
    const rawT = (v - min) / span;
    return reversed ? 1 - rawT : rawT;
  };
  const mapScreen = (t: number, plot: GeometryBox): number =>
    horizontal ? plot.x + t * plot.width : plot.y + (1 - t) * plot.height;
  const zeroVal = Math.max(min, Math.min(max, 0));

  return {
    kind: "value",
    axis,
    min,
    max,
    majorTicks,
    minorTicks,
    reversed,
    numberFormat,
    mapValue,
    mapScreen,
    valueToScreen(v, plot) {
      return mapScreen(mapValue(v), plot);
    },
    zeroValue() {
      return zeroVal;
    },
    zeroScreen(plot) {
      return mapScreen(mapValue(zeroVal), plot);
    },
  };
}

/** Computes axis scales for all axes in the parsed chart. */
export function buildAxisScales(
  parsed: ParsedChart,
): ReadonlyMap<string, AxisScale> {
  const result = new Map<string, AxisScale>();
  for (const axis of parsed.axes) {
    switch (axis.kind) {
      case "category":
        result.set(axis.name, buildCategoryAxisScale(axis, parsed.groups));
        break;
      case "value":
        result.set(axis.name, buildValueAxisScale(axis, parsed.groups));
        break;
      default:
        assertNever(axis);
    }
  }
  return result;
}

/** Resolves the normalized unreversed/reversed position of `axis` inside `crossScale`. */
export function resolveAxisCrossingNormalized(
  axis: ParsedChartAxis,
  crossScale: AxisScale,
): number {
  const crossing = axis.crossing;
  switch (crossScale.kind) {
    case "value": {
      if (crossing.kind === "crossesAt") {
        const clamped = Math.max(
          crossScale.min,
          Math.min(crossScale.max, crossing.value),
        );
        return crossScale.mapValue(clamped);
      }
      if (crossing.value === "min") {
        return crossScale.mapValue(crossScale.min);
      }
      if (crossing.value === "max") {
        return crossScale.mapValue(crossScale.max);
      }
      return crossScale.mapValue(crossScale.zeroValue());
    }
    case "category": {
      let rawT: number;
      if (crossing.kind === "crossesAt") {
        rawT =
          crossScale.crossBetween === "midCat" && crossScale.count > 1
            ? Math.max(
                0,
                Math.min(1, (crossing.value - 1) / (crossScale.count - 1)),
              )
            : Math.max(0, Math.min(1, (crossing.value - 1) / crossScale.count));
      } else if (crossing.value === "max") {
        rawT = 1;
      } else {
        rawT = 0;
      }
      return crossScale.reversed ? 1 - rawT : rawT;
    }
    default:
      return assertNever(crossScale);
  }
}

/** Resolves the screen coordinate of `axis` inside `crossScale` and `plot`. */
export function resolveAxisCrossingScreen(
  axis: ParsedChartAxis,
  crossScale: AxisScale,
  plot: GeometryBox,
): number {
  const t = resolveAxisCrossingNormalized(axis, crossScale);
  return crossScale.mapScreen(t, plot);
}
