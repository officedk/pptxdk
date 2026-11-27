import type { ColorValue, Rect, ThemeContext } from "@pptxdk/core/capability";

import type { ParsedAreaGroup, ParsedLineGroup } from "../chart-model.ts";
import type { ChartDisplayBlanksAs, ChartLine } from "../chart-types.ts";
import {
  renderMarkerSvg,
  resolveEffectiveMarker,
  type ResolvedMarker,
} from "./markers.ts";
import {
  chartAccentSlot,
  resolveChartLineStroke,
  resolveSeriesColor,
  strokeAttrs,
  type ResolvedStroke,
} from "./style.ts";
import type { AxisScale, PlottedCartesianPointAnchor } from "./types.ts";

export type Point2D = {
  readonly x: number;
  readonly y: number;
};

/**
 * Computes Catmull-Rom cubic Bezier control points (tension = 1/6) for segment `pCurr -> pNext`.
 */
export function catmullRomControlPoints(
  pPrev: Point2D,
  pCurr: Point2D,
  pNext: Point2D,
  pNext2: Point2D,
): { readonly c1: Point2D; readonly c2: Point2D } {
  return {
    c1: {
      x: pCurr.x + (pNext.x - pPrev.x) / 6,
      y: pCurr.y + (pNext.y - pPrev.y) / 6,
    },
    c2: {
      x: pNext.x - (pNext2.x - pCurr.x) / 6,
      y: pNext.y - (pNext2.y - pCurr.y) / 6,
    },
  };
}

/**
 * Builds an SVG polyline or Catmull-Rom cubic Bezier path for a contiguous point sequence.
 */
export function buildSmoothOrLinearPath(
  points: readonly Point2D[],
  smooth: boolean,
): string {
  if (points.length === 0) return "";
  const first = points[0]!;
  if (points.length === 1) return `M${first.x} ${first.y}`;
  if (!smooth || points.length === 2) {
    return (
      `M${first.x} ${first.y}` +
      points
        .slice(1)
        .map((p) => `L${p.x} ${p.y}`)
        .join("")
    );
  }
  const cmds: string[] = [`M${first.x} ${first.y}`];
  for (let i = 0; i < points.length - 1; i += 1) {
    const pPrev = points[Math.max(0, i - 1)]!;
    const pCurr = points[i]!;
    const pNext = points[i + 1]!;
    const pNext2 = points[Math.min(points.length - 1, i + 2)]!;
    const { c1, c2 } = catmullRomControlPoints(pPrev, pCurr, pNext, pNext2);
    cmds.push(`C${c1.x} ${c1.y} ${c2.x} ${c2.y} ${pNext.x} ${pNext.y}`);
  }
  return cmds.join("");
}

function strokesEqual(a: ResolvedStroke, b: ResolvedStroke): boolean {
  return (
    a.hex === b.hex &&
    a.alpha === b.alpha &&
    a.width === b.width &&
    a.dashArray === b.dashArray
  );
}

/**
 * Renders a contiguous line run while respecting per-point line/color overrides on incoming segments.
 */
export function renderStyledLineRun<T extends Point2D>(options: {
  readonly run: readonly T[];
  readonly smooth: boolean;
  readonly baseStroke: ResolvedStroke;
  readonly segmentStrokeOf: (point: T) => ResolvedStroke;
}): string {
  const { run, smooth, baseStroke, segmentStrokeOf } = options;
  if (run.length === 0) return "";
  if (run.length === 1) {
    return "";
  }

  const segStrokes = run.slice(1).map((pt) => segmentStrokeOf(pt));
  const allMatchBase = segStrokes.every((st) => strokesEqual(st, baseStroke));
  if (allMatchBase) {
    const d = buildSmoothOrLinearPath(run, smooth);
    return `<path d="${d}" fill="none" ${strokeAttrs(baseStroke)} stroke-linejoin="round" stroke-linecap="round"/>`;
  }

  // Batch consecutive segments sharing the same resolved stroke to keep SVG compact.
  const parts: string[] = [];
  let segIdx = 0;
  while (segIdx < segStrokes.length) {
    const st = segStrokes[segIdx]!;
    const startSeg = segIdx;
    while (
      segIdx < segStrokes.length &&
      strokesEqual(segStrokes[segIdx]!, st)
    ) {
      segIdx += 1;
    }
    const endSeg = segIdx;
    const cmds: string[] = [`M${run[startSeg]!.x} ${run[startSeg]!.y}`];
    for (let i = startSeg; i < endSeg; i += 1) {
      const pCurr = run[i]!;
      const pNext = run[i + 1]!;
      if (!smooth || run.length === 2) {
        cmds.push(`L${pNext.x} ${pNext.y}`);
      } else {
        const pPrev = run[Math.max(0, i - 1)]!;
        const pNext2 = run[Math.min(run.length - 1, i + 2)]!;
        const { c1, c2 } = catmullRomControlPoints(pPrev, pCurr, pNext, pNext2);
        cmds.push(`C${c1.x} ${c1.y} ${c2.x} ${c2.y} ${pNext.x} ${pNext.y}`);
      }
    }
    parts.push(
      `<path d="${cmds.join("")}" fill="none" ${strokeAttrs(st)} stroke-linejoin="round" stroke-linecap="round"/>`,
    );
  }
  return parts.join("");
}

export function resolvePointSegmentStroke(options: {
  readonly pointLine: ChartLine | undefined;
  readonly pointColor: ColorValue | undefined;
  readonly seriesLine: ChartLine | undefined;
  readonly seriesColor: ColorValue | undefined;
  readonly seriesIndex: number;
  readonly baseStroke: ResolvedStroke;
  readonly theme: ThemeContext;
}): ResolvedStroke {
  const {
    pointLine,
    pointColor,
    seriesLine,
    seriesColor,
    seriesIndex,
    baseStroke,
    theme,
  } = options;
  if (pointLine === undefined && pointColor === undefined) {
    return baseStroke;
  }
  const mergedLine: ChartLine | undefined =
    pointLine !== undefined
      ? {
          ...pointLine,
          ...(pointLine.color === undefined && pointColor !== undefined
            ? { color: pointColor }
            : {}),
        }
      : {
          ...seriesLine,
          color: pointColor!,
        };
  return resolveChartLineStroke(mergedLine, theme, {
    color:
      pointColor ??
      seriesLine?.color ??
      seriesColor ??
      chartAccentSlot(seriesIndex),
    width: baseStroke.width,
    alpha: baseStroke.alpha,
  })!;
}

type LinePlottedNode = Point2D & {
  readonly pointIndex: number;
  readonly marker: ResolvedMarker | undefined;
  readonly segmentStroke: ResolvedStroke;
};

function renderLineGroup(options: {
  readonly group: ParsedLineGroup;
  readonly groupIndex: number;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: Rect;
  readonly displayBlanksAs: ChartDisplayBlanksAs | undefined;
  readonly theme: ThemeContext;
}): {
  readonly svg: string;
  readonly anchors: readonly PlottedCartesianPointAnchor[];
} {
  const { group, groupIndex, scales, plotRect, displayBlanksAs, theme } =
    options;
  const catScale = scales.get(group.axes.category.name);
  const valScale = scales.get(group.axes.value.name);
  if (catScale?.kind !== "category" || valScale?.kind !== "value") {
    return { svg: "", anchors: [] };
  }

  const categories = group.data.categories;
  const seriesCount = group.series.length;
  if (categories.length === 0 || seriesCount === 0) {
    return { svg: "", anchors: [] };
  }

  const isStacked = group.stacking !== "standard";
  const isPercent = group.stacking === "percent";
  const effectiveBlanks: ChartDisplayBlanksAs = isStacked
    ? "zero"
    : (displayBlanksAs ?? "gap");

  const percentTotals: number[] = isPercent
    ? categories.map((_, catIdx) => {
        let sum = 0;
        for (const s of group.series) {
          const v = s.values[catIdx];
          if (v !== null && v !== undefined && Number.isFinite(v)) {
            sum += Math.abs(v);
          }
        }
        return sum;
      })
    : [];

  const posAcc = Array.from({ length: categories.length }, () => 0);
  const negAcc = Array.from({ length: categories.length }, () => 0);

  const parts: string[] = [];
  const anchors: PlottedCartesianPointAnchor[] = [];

  for (let s = 0; s < seriesCount; s += 1) {
    const series = group.series[s]!;
    const baseStroke = resolveChartLineStroke(series.line, theme, {
      color: series.color ?? chartAccentSlot(series.seriesIndex),
      width: 2,
    })!;

    const runs: LinePlottedNode[][] = [];
    let currentRun: LinePlottedNode[] = [];
    const markerQueue: {
      readonly x: number;
      readonly y: number;
      readonly marker: ResolvedMarker | undefined;
      readonly isolated: boolean;
    }[] = [];

    for (let catIdx = 0; catIdx < categories.length; catIdx += 1) {
      const raw = series.values[catIdx];
      const isMissing =
        raw === null || raw === undefined || !Number.isFinite(raw);

      let plotted: number | undefined;
      let percent: number | undefined;

      if (isMissing) {
        if (effectiveBlanks === "zero") {
          plotted = 0;
        } else if (effectiveBlanks === "span") {
          continue;
        } else {
          if (currentRun.length > 0) {
            runs.push(currentRun);
            currentRun = [];
          }
          continue;
        }
      } else {
        if (isPercent) {
          const totalAbs = percentTotals[catIdx] ?? 0;
          plotted = totalAbs === 0 ? 0 : raw / totalAbs;
          percent = totalAbs === 0 ? 0 : Math.abs(raw) / totalAbs;
        } else {
          plotted = raw;
        }
      }

      let stackedVal = plotted;
      if (isStacked) {
        if (plotted >= 0) {
          posAcc[catIdx] = (posAcc[catIdx] ?? 0) + plotted;
          stackedVal = posAcc[catIdx]!;
        } else {
          negAcc[catIdx] = (negAcc[catIdx] ?? 0) + plotted;
          stackedVal = negAcc[catIdx]!;
        }
      }

      if (valScale.logBase !== undefined && stackedVal <= 0) {
        if (effectiveBlanks === "gap" && currentRun.length > 0) {
          runs.push(currentRun);
          currentRun = [];
        }
        continue;
      }

      const x = catScale.indexToScreen(catIdx, plotRect);
      const y = valScale.valueToScreen(stackedVal, plotRect);

      const point = series.points?.[catIdx];
      const marker = resolveEffectiveMarker({
        groupDefaultEnabled: false,
        seriesMarker: series.marker,
        pointMarker: point?.marker,
        seriesColor: series.color,
        pointColor: point?.color,
        seriesIndex: series.seriesIndex,
        theme,
      });

      const segmentStroke = resolvePointSegmentStroke({
        pointLine: point?.line,
        pointColor: point?.color,
        seriesLine: series.line,
        seriesColor: series.color,
        seriesIndex: series.seriesIndex,
        baseStroke,
        theme,
      });

      currentRun.push({
        x,
        y,
        pointIndex: catIdx,
        marker,
        segmentStroke,
      });

      if (!isMissing) {
        anchors.push({
          kind: "point",
          groupIndex,
          seriesIndex: series.seriesIndex,
          groupSeriesIndex: s,
          pointIndex: catIdx,
          value: raw,
          plottedValue: stackedVal,
          category: categories[catIdx] ?? "",
          seriesName: series.name,
          x,
          y,
          ...(marker === undefined ? {} : { markerRadius: marker.radius }),
          ...(percent === undefined ? {} : { percent }),
        });
      }
    }

    if (currentRun.length > 0) {
      runs.push(currentRun);
    }

    for (const run of runs) {
      if (run.length >= 2) {
        parts.push(
          renderStyledLineRun({
            run,
            smooth: series.smooth === true,
            baseStroke,
            segmentStrokeOf: (pt) => pt.segmentStroke,
          }),
        );
      }
      for (const node of run) {
        markerQueue.push({
          x: node.x,
          y: node.y,
          marker: node.marker,
          isolated: run.length === 1,
        });
      }
    }

    for (const item of markerQueue) {
      if (item.marker !== undefined) {
        parts.push(renderMarkerSvg(item.x, item.y, item.marker));
      } else if (item.isolated) {
        // Isolated valid point in gap mode without a marker renders a minimal dot so it stays visible.
        const dotColor = resolveSeriesColor(
          series.line?.color ?? series.color,
          series.seriesIndex,
          theme,
        );
        parts.push(
          `<circle cx="${item.x}" cy="${item.y}" r="${Math.max(1.5, baseStroke.width * 0.75)}" fill="#${dotColor.hex}" fill-opacity="${dotColor.alpha ?? 1}"/>`,
        );
      }
    }
  }

  return { svg: parts.join(""), anchors };
}

type AreaRunNode = Point2D & {
  readonly pointIndex: number;
  readonly rawValue: number;
  readonly isMissing: boolean;
  readonly percent?: number;
};

function renderAreaGroup(options: {
  readonly group: ParsedAreaGroup;
  readonly groupIndex: number;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: Rect;
  readonly displayBlanksAs: ChartDisplayBlanksAs | undefined;
  readonly theme: ThemeContext;
}): {
  readonly svg: string;
  readonly anchors: readonly PlottedCartesianPointAnchor[];
} {
  const { group, groupIndex, scales, plotRect, displayBlanksAs, theme } =
    options;
  const catScale = scales.get(group.axes.category.name);
  const valScale = scales.get(group.axes.value.name);
  if (catScale?.kind !== "category" || valScale?.kind !== "value") {
    return { svg: "", anchors: [] };
  }

  const categories = group.data.categories;
  const seriesCount = group.series.length;
  if (categories.length === 0 || seriesCount === 0) {
    return { svg: "", anchors: [] };
  }

  const isStacked = group.stacking !== "standard";
  const isPercent = group.stacking === "percent";
  const effectiveBlanks: "gap" | "zero" =
    isStacked || displayBlanksAs === "zero" ? "zero" : "gap";

  const percentTotals: number[] = isPercent
    ? categories.map((_, catIdx) => {
        let sum = 0;
        for (const s of group.series) {
          const v = s.values[catIdx];
          if (v !== null && v !== undefined && Number.isFinite(v)) {
            sum += Math.abs(v);
          }
        }
        return sum;
      })
    : [];

  const posAcc = Array.from({ length: categories.length }, () => 0);
  const negAcc = Array.from({ length: categories.length }, () => 0);

  const parts: string[] = [];
  const anchors: PlottedCartesianPointAnchor[] = [];
  const baseY = valScale.zeroScreen(plotRect);

  for (let s = 0; s < seriesCount; s += 1) {
    const series = group.series[s]!;
    const fill = resolveSeriesColor(series.color, series.seriesIndex, theme);
    const stroke =
      series.line === undefined
        ? undefined
        : resolveChartLineStroke(series.line, theme, {
            color: "dk1",
            width: 1,
          });
    const strokeAttr = stroke === undefined ? "" : ` ${strokeAttrs(stroke)}`;
    const fillAttr = `fill="#${fill.hex}" fill-opacity="${fill.alpha ?? 1}"`;

    if (!isStacked) {
      const runs: AreaRunNode[][] = [];
      let currentRun: AreaRunNode[] = [];

      for (let catIdx = 0; catIdx < categories.length; catIdx += 1) {
        const raw = series.values[catIdx];
        const isMissing =
          raw === null || raw === undefined || !Number.isFinite(raw);
        if (isMissing && effectiveBlanks === "gap") {
          if (currentRun.length > 0) {
            runs.push(currentRun);
            currentRun = [];
          }
          continue;
        }
        const plotted = isMissing ? 0 : raw;
        if (valScale.logBase !== undefined && plotted <= 0) {
          if (currentRun.length > 0) {
            runs.push(currentRun);
            currentRun = [];
          }
          continue;
        }
        const x = catScale.indexToScreen(catIdx, plotRect);
        const y = valScale.valueToScreen(plotted, plotRect);
        currentRun.push({
          x,
          y,
          pointIndex: catIdx,
          rawValue: plotted,
          isMissing,
        });
      }
      if (currentRun.length > 0) runs.push(currentRun);

      for (const run of runs) {
        if (run.length >= 2) {
          const first = run[0]!;
          const last = run[run.length - 1]!;
          const topSegs = run.map((pt) => `L${pt.x} ${pt.y}`).join("");
          parts.push(
            `<path d="M${first.x} ${baseY}${topSegs}L${last.x} ${baseY}Z" ${fillAttr}${strokeAttr} stroke-linejoin="round"/>`,
          );
        } else if (run.length === 1 && stroke !== undefined) {
          const pt = run[0]!;
          parts.push(
            `<path d="M${pt.x} ${baseY}V${pt.y}" fill="none" ${strokeAttrs(stroke)}/>`,
          );
        }
        for (const node of run) {
          if (!node.isMissing) {
            anchors.push({
              kind: "point",
              groupIndex,
              seriesIndex: series.seriesIndex,
              groupSeriesIndex: s,
              pointIndex: node.pointIndex,
              value: node.rawValue,
              category: categories[node.pointIndex] ?? "",
              seriesName: series.name,
              x: node.x,
              y: node.y,
            });
          }
        }
      }
    } else {
      const topPts: Point2D[] = [];
      const botPts: Point2D[] = [];

      for (let catIdx = 0; catIdx < categories.length; catIdx += 1) {
        const raw = series.values[catIdx];
        const isMissing =
          raw === null || raw === undefined || !Number.isFinite(raw);
        let plotted: number;
        let percent: number | undefined;
        if (isMissing) {
          plotted = 0;
        } else if (isPercent) {
          const totalAbs = percentTotals[catIdx] ?? 0;
          plotted = totalAbs === 0 ? 0 : raw / totalAbs;
          percent = totalAbs === 0 ? 0 : Math.abs(raw) / totalAbs;
        } else {
          plotted = raw;
        }

        let botVal: number;
        let topVal: number;
        if (plotted >= 0) {
          botVal = posAcc[catIdx] ?? 0;
          topVal = botVal + plotted;
          posAcc[catIdx] = topVal;
        } else {
          botVal = negAcc[catIdx] ?? 0;
          topVal = botVal + plotted;
          negAcc[catIdx] = topVal;
        }

        if (valScale.logBase !== undefined && topVal <= 0) {
          continue;
        }
        const safeBotVal =
          valScale.logBase !== undefined
            ? Math.max(valScale.min, botVal)
            : botVal;

        const x = catScale.indexToScreen(catIdx, plotRect);
        const topY = valScale.valueToScreen(topVal, plotRect);
        const botY = valScale.valueToScreen(safeBotVal, plotRect);
        topPts.push({ x, y: topY });
        botPts.push({ x, y: botY });

        if (!isMissing) {
          anchors.push({
            kind: "point",
            groupIndex,
            seriesIndex: series.seriesIndex,
            groupSeriesIndex: s,
            pointIndex: catIdx,
            value: raw,
            plottedValue: topVal,
            category: categories[catIdx] ?? "",
            seriesName: series.name,
            x,
            y: topY,
            ...(percent === undefined ? {} : { percent }),
          });
        }
      }

      if (topPts.length >= 2) {
        const firstTop = topPts[0]!;
        const topCmds = topPts
          .slice(1)
          .map((pt) => `L${pt.x} ${pt.y}`)
          .join("");
        const botCmds = botPts
          .toReversed()
          .map((pt) => `L${pt.x} ${pt.y}`)
          .join("");
        parts.push(
          `<path d="M${firstTop.x} ${firstTop.y}${topCmds}${botCmds}Z" ${fillAttr}${strokeAttr} stroke-linejoin="round"/>`,
        );
      }
    }
  }

  return { svg: parts.join(""), anchors };
}

export function renderLineOrAreaGroup(options: {
  readonly group: ParsedLineGroup | ParsedAreaGroup;
  readonly groupIndex: number;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: Rect;
  readonly displayBlanksAs: ChartDisplayBlanksAs | undefined;
  readonly theme: ThemeContext;
}): {
  readonly svg: string;
  readonly anchors: readonly PlottedCartesianPointAnchor[];
} {
  if (options.group.kind === "line") {
    return renderLineGroup({
      ...options,
      group: options.group,
    });
  }
  return renderAreaGroup({
    ...options,
    group: options.group,
  });
}
