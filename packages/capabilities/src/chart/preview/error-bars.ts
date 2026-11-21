import type { GeometryBox, ThemeContext } from "@pptxdk/core/capability";

import type { ParsedChart } from "../chart-model.ts";
import type { ChartErrorBars } from "../chart-types.ts";
import { resolveChartLineStroke, strokeAttrs } from "./style.ts";
import type { AxisScale, PlottedPointAnchor, ValueAxisScale } from "./types.ts";

export type ComputedErrorBarBound = {
  readonly index: number;
  readonly low: number;
  readonly high: number;
  readonly hasPlus: boolean;
  readonly hasMinus: boolean;
};

/**
 * Computes per-slot error bar bounds in data space for any of the 5 OOXML error bar types.
 */
export function computeErrorBarBounds(
  bars: ChartErrorBars,
  values: readonly (number | null)[],
  options?: {
    readonly centers?: readonly (number | undefined)[];
    readonly skipMask?: readonly boolean[];
  },
): ComputedErrorBarBound[] {
  const valid: number[] = [];
  for (let i = 0; i < values.length; i += 1) {
    if (options?.skipMask?.[i]) continue;
    const v = values[i];
    if (v !== null && v !== undefined && Number.isFinite(v)) {
      valid.push(v);
    }
  }
  const n = valid.length;
  if (n === 0) return [];

  let sum = 0;
  for (const v of valid) sum += v;
  const mean = sum / n;

  let sqSum = 0;
  for (const v of valid) {
    const d = v - mean;
    sqSum += d * d;
  }
  const stdDev = n > 1 ? Math.sqrt(sqSum / (n - 1)) : 0;
  const stdErr = n > 1 ? stdDev / Math.sqrt(n) : 0;
  const include = bars.include ?? "both";

  const results: ComputedErrorBarBound[] = [];
  for (let i = 0; i < values.length; i += 1) {
    if (options?.skipMask?.[i]) continue;
    const raw = values[i];
    if (raw === null || raw === undefined || !Number.isFinite(raw)) continue;

    let center = options?.centers?.[i] ?? raw;
    let plusMag: number | undefined;
    let minusMag: number | undefined;

    switch (bars.type) {
      case "fixed":
        plusMag = bars.value;
        minusMag = bars.value;
        break;
      case "percentage": {
        const mag = Math.abs(raw) * (bars.value / 100);
        plusMag = mag;
        minusMag = mag;
        break;
      }
      case "standardError":
        plusMag = stdErr;
        minusMag = stdErr;
        break;
      case "standardDeviation": {
        center = mean;
        const mag = bars.value * stdDev;
        plusMag = mag;
        minusMag = mag;
        break;
      }
      case "custom":
        plusMag =
          bars.plus?.[i] !== null && bars.plus?.[i] !== undefined
            ? bars.plus[i]!
            : undefined;
        minusMag =
          bars.minus?.[i] !== null && bars.minus?.[i] !== undefined
            ? bars.minus[i]!
            : undefined;
        break;
    }

    const hasPlus =
      (include === "both" || include === "plus") &&
      plusMag !== undefined &&
      Number.isFinite(plusMag);
    const hasMinus =
      (include === "both" || include === "minus") &&
      minusMag !== undefined &&
      Number.isFinite(minusMag);

    if (!hasPlus && !hasMinus) continue;

    results.push({
      index: i,
      low: hasMinus ? center - minusMag! : center,
      high: hasPlus ? center + plusMag! : center,
      hasPlus,
      hasMinus,
    });
  }

  return results;
}

function clampForScale(value: number, scale: ValueAxisScale): number {
  if (scale.logBase !== undefined && value <= 0) {
    return scale.min;
  }
  return value;
}

function renderSingleErrorBarSvg(options: {
  readonly bound: ComputedErrorBarBound;
  readonly crossCoord: number;
  readonly axisScale: ValueAxisScale;
  readonly horizontal: boolean;
  readonly plotRect: GeometryBox;
  readonly endCap: boolean;
  readonly strokeAttr: string;
}): string {
  const {
    bound,
    crossCoord,
    axisScale,
    horizontal,
    plotRect,
    endCap,
    strokeAttr,
  } = options;
  const lowCoord = axisScale.valueToScreen(
    clampForScale(bound.low, axisScale),
    plotRect,
  );
  const highCoord = axisScale.valueToScreen(
    clampForScale(bound.high, axisScale),
    plotRect,
  );
  const capHalf = 3;

  if (horizontal) {
    const segs = [
      `<line x1="${lowCoord}" y1="${crossCoord}" x2="${highCoord}" y2="${crossCoord}" ${strokeAttr} fill="none"/>`,
    ];
    if (endCap) {
      if (bound.hasPlus) {
        segs.push(
          `<line x1="${highCoord}" y1="${crossCoord - capHalf}" x2="${highCoord}" y2="${crossCoord + capHalf}" ${strokeAttr} fill="none"/>`,
        );
      }
      if (bound.hasMinus) {
        segs.push(
          `<line x1="${lowCoord}" y1="${crossCoord - capHalf}" x2="${lowCoord}" y2="${crossCoord + capHalf}" ${strokeAttr} fill="none"/>`,
        );
      }
    }
    return segs.join("");
  }

  const segs = [
    `<line x1="${crossCoord}" y1="${highCoord}" x2="${crossCoord}" y2="${lowCoord}" ${strokeAttr} fill="none"/>`,
  ];
  if (endCap) {
    if (bound.hasPlus) {
      segs.push(
        `<line x1="${crossCoord - capHalf}" y1="${highCoord}" x2="${crossCoord + capHalf}" y2="${highCoord}" ${strokeAttr} fill="none"/>`,
      );
    }
    if (bound.hasMinus) {
      segs.push(
        `<line x1="${crossCoord - capHalf}" y1="${lowCoord}" x2="${crossCoord + capHalf}" y2="${lowCoord}" ${strokeAttr} fill="none"/>`,
      );
    }
  }
  return segs.join("");
}

export function renderErrorBarsSvg(options: {
  readonly parsed: ParsedChart;
  readonly anchors: readonly PlottedPointAnchor[];
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: GeometryBox;
  readonly theme: ThemeContext;
}): string {
  const { parsed, anchors, scales, plotRect, theme } = options;
  const anchorMap = new Map<string, PlottedPointAnchor>();
  for (const anchor of anchors) {
    if (anchor.kind === "slice") continue;
    anchorMap.set(
      `${anchor.groupIndex}:${anchor.groupSeriesIndex}:${anchor.pointIndex}`,
      anchor,
    );
  }

  const parts: string[] = [];

  for (let groupIndex = 0; groupIndex < parsed.groups.length; groupIndex += 1) {
    const group = parsed.groups[groupIndex]!;
    if (group.kind === "pie" || group.kind === "doughnut") continue;

    if (group.kind === "scatter") {
      const xScale = scales.get(group.axes.x.name);
      const yScale = scales.get(group.axes.y.name);
      if (xScale?.kind !== "value" || yScale?.kind !== "value") continue;

      for (let s = 0; s < group.series.length; s += 1) {
        const series = group.series[s]!;
        const xBars = series.errorBars?.x;
        const yBars = series.errorBars?.y;
        if (xBars === undefined && yBars === undefined) continue;

        const skipMask = series.x.map(
          (_, i) => !anchorMap.has(`${groupIndex}:${s}:${i}`),
        );

        if (xBars !== undefined) {
          const stroke = resolveChartLineStroke(xBars.line, theme, {
            color: "dk1",
            width: 1,
          });
          if (stroke !== undefined) {
            const strokeAttr = strokeAttrs(stroke);
            const xVals = series.x.map((v, i) => {
              const a = anchorMap.get(`${groupIndex}:${s}:${i}`);
              return a?.kind === "point" ? (a.xValue ?? v) : v;
            });
            const bounds = computeErrorBarBounds(xBars, xVals, { skipMask });
            for (const bound of bounds) {
              const a = anchorMap.get(`${groupIndex}:${s}:${bound.index}`);
              if (a?.kind !== "point") continue;
              parts.push(
                renderSingleErrorBarSvg({
                  bound,
                  crossCoord: a.y,
                  axisScale: xScale,
                  horizontal: true,
                  plotRect,
                  endCap: xBars.endCap !== false,
                  strokeAttr,
                }),
              );
            }
          }
        }

        if (yBars !== undefined) {
          const stroke = resolveChartLineStroke(yBars.line, theme, {
            color: "dk1",
            width: 1,
          });
          if (stroke !== undefined) {
            const strokeAttr = strokeAttrs(stroke);
            const yVals = series.y.map((v, i) => {
              const a = anchorMap.get(`${groupIndex}:${s}:${i}`);
              return a?.kind === "point" ? a.value : v;
            });
            const bounds = computeErrorBarBounds(yBars, yVals, { skipMask });
            for (const bound of bounds) {
              const a = anchorMap.get(`${groupIndex}:${s}:${bound.index}`);
              if (a?.kind !== "point") continue;
              parts.push(
                renderSingleErrorBarSvg({
                  bound,
                  crossCoord: a.x,
                  axisScale: yScale,
                  horizontal: false,
                  plotRect,
                  endCap: yBars.endCap !== false,
                  strokeAttr,
                }),
              );
            }
          }
        }
      }
      continue;
    }

    const valScale = scales.get(group.axes.value.name);
    if (valScale?.kind !== "value") continue;
    const isHorizontal = group.kind === "bar";

    for (let s = 0; s < group.series.length; s += 1) {
      const series = group.series[s]!;
      const bars = series.errorBars;
      if (bars === undefined) continue;

      const stroke = resolveChartLineStroke(bars.line, theme, {
        color: "dk1",
        width: 1,
      });
      if (stroke === undefined) continue;
      const strokeAttr = strokeAttrs(stroke);

      const skipMask = series.values.map(
        (_, i) => !anchorMap.has(`${groupIndex}:${s}:${i}`),
      );
      const centers = series.values.map((_, i) => {
        const a = anchorMap.get(`${groupIndex}:${s}:${i}`);
        return a !== undefined && a.kind !== "slice"
          ? a.plottedValue
          : undefined;
      });
      const bounds = computeErrorBarBounds(bars, series.values, {
        centers,
        skipMask,
      });

      for (const bound of bounds) {
        const a = anchorMap.get(`${groupIndex}:${s}:${bound.index}`);
        if (a === undefined || a.kind === "slice") continue;
        const crossCoord =
          a.kind === "bar"
            ? isHorizontal
              ? a.y + a.height / 2
              : a.x + a.width / 2
            : a.x;
        parts.push(
          renderSingleErrorBarSvg({
            bound,
            crossCoord,
            axisScale: valScale,
            horizontal: isHorizontal,
            plotRect,
            endCap: bars.endCap !== false,
            strokeAttr,
          }),
        );
      }
    }
  }

  return parts.join("");
}
