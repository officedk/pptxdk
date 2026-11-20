import type { GeometryBox, ThemeContext } from "@react-pptx/core/capability";

import type { ParsedScatterGroup } from "../chart-model.ts";
import type { ChartDisplayBlanksAs } from "../chart-types.ts";
import {
  renderStyledLineRun,
  resolvePointSegmentStroke,
  type Point2D,
} from "./line-area.ts";
import {
  renderMarkerSvg,
  resolveEffectiveMarker,
  type ResolvedMarker,
} from "./markers.ts";
import {
  resolveChartLineStroke,
  resolveSeriesColor,
  type ResolvedStroke,
} from "./style.ts";
import type { AxisScale, PlottedCartesianPointAnchor } from "./types.ts";

type ScatterPlottedNode = Point2D & {
  readonly pointIndex: number;
  readonly marker: ResolvedMarker | undefined;
  readonly segmentStroke: ResolvedStroke;
};

export function renderScatterGroup(options: {
  readonly group: ParsedScatterGroup;
  readonly groupIndex: number;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: GeometryBox;
  readonly displayBlanksAs: ChartDisplayBlanksAs | undefined;
  readonly theme: ThemeContext;
}): {
  readonly svg: string;
  readonly anchors: readonly PlottedCartesianPointAnchor[];
} {
  const { group, groupIndex, scales, plotRect, displayBlanksAs, theme } =
    options;
  const xScale = scales.get(group.axes.x.name);
  const yScale = scales.get(group.axes.y.name);
  if (xScale?.kind !== "value" || yScale?.kind !== "value") {
    return { svg: "", anchors: [] };
  }
  if (group.series.length === 0) {
    return { svg: "", anchors: [] };
  }

  const hasLine = group.line !== false;
  const isSmooth = group.line === "smooth";
  const blanksMode: ChartDisplayBlanksAs = displayBlanksAs ?? "gap";
  const parts: string[] = [];
  const anchors: PlottedCartesianPointAnchor[] = [];

  for (let s = 0; s < group.series.length; s += 1) {
    const series = group.series[s]!;
    if (series.x.length === 0 || series.y.length === 0) continue;

    const baseStroke = resolveChartLineStroke(series.line, theme, {
      color: series.color ?? `accent${(series.seriesIndex % 6) + 1}`,
      width: 2,
    })!;

    const runs: ScatterPlottedNode[][] = [];
    let currentRun: ScatterPlottedNode[] = [];
    const markerQueue: {
      readonly x: number;
      readonly y: number;
      readonly marker: ResolvedMarker | undefined;
      readonly isolated: boolean;
    }[] = [];

    for (let i = 0; i < series.x.length; i += 1) {
      const rawX = series.x[i];
      const rawY = series.y[i];
      const missingX =
        rawX === null || rawX === undefined || !Number.isFinite(rawX);
      const missingY =
        rawY === null || rawY === undefined || !Number.isFinite(rawY);
      const isMissing = missingX || missingY;

      let xVal: number;
      let yVal: number;
      if (isMissing) {
        if (blanksMode === "zero") {
          xVal = missingX ? 0 : rawX;
          yVal = missingY ? 0 : rawY;
        } else if (blanksMode === "span") {
          continue;
        } else {
          if (currentRun.length > 0) {
            runs.push(currentRun);
            currentRun = [];
          }
          continue;
        }
      } else {
        xVal = rawX;
        yVal = rawY;
      }

      if (
        (xScale.logBase !== undefined && xVal <= 0) ||
        (yScale.logBase !== undefined && yVal <= 0)
      ) {
        if (blanksMode === "gap" && currentRun.length > 0) {
          runs.push(currentRun);
          currentRun = [];
        }
        continue;
      }

      const x = xScale.valueToScreen(xVal, plotRect);
      const y = yScale.valueToScreen(yVal, plotRect);
      const point = series.points?.[i];

      const marker = group.markers
        ? resolveEffectiveMarker({
            groupDefaultEnabled: true,
            seriesMarker: series.marker,
            pointMarker: point?.marker,
            seriesColor: series.color,
            pointColor: point?.color,
            seriesIndex: series.seriesIndex,
            theme,
          })
        : undefined;

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
        pointIndex: i,
        marker,
        segmentStroke,
      });

      if (!isMissing) {
        anchors.push({
          kind: "point",
          groupIndex,
          seriesIndex: series.seriesIndex,
          groupSeriesIndex: s,
          pointIndex: i,
          value: rawY,
          xValue: rawX,
          seriesName: series.name,
          x,
          y,
          ...(marker === undefined ? {} : { markerRadius: marker.radius }),
        });
      }
    }

    if (currentRun.length > 0) {
      runs.push(currentRun);
    }

    for (const run of runs) {
      if (hasLine && run.length >= 2) {
        parts.push(
          renderStyledLineRun({
            run,
            smooth: isSmooth,
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
          isolated: hasLine && run.length === 1,
        });
      }
    }

    for (const item of markerQueue) {
      if (item.marker !== undefined) {
        parts.push(renderMarkerSvg(item.x, item.y, item.marker));
      } else if (item.isolated) {
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
