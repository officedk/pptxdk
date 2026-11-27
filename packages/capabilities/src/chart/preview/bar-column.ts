import type { Rect, ThemeContext } from "@pptxdk/core/capability";

import type { ParsedColumnOrBarGroup } from "../chart-model.ts";
import {
  resolveChartLineStroke,
  resolveSeriesColor,
  strokeAttrs,
} from "./style.ts";
import type { AxisScale, PlottedBarAnchor } from "./types.ts";

export function renderColumnOrBarGroup(options: {
  readonly group: ParsedColumnOrBarGroup;
  readonly groupIndex: number;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: Rect;
  readonly theme: ThemeContext;
}): {
  readonly svg: string;
  readonly anchors: readonly PlottedBarAnchor[];
} {
  const { group, groupIndex, scales, plotRect, theme } = options;
  const catScale = scales.get(group.axes.category.name);
  const valScale = scales.get(group.axes.value.name);
  if (catScale?.kind !== "category" || valScale?.kind !== "value") {
    return { svg: "", anchors: [] };
  }

  const isHorizontal = group.kind === "bar";
  const categories = group.data.categories;
  const seriesCount = group.series.length;
  if (categories.length === 0 || seriesCount === 0) {
    return { svg: "", anchors: [] };
  }

  const slotSpan = catScale.slotSpan(plotRect);
  const g = group.gapWidth / 100;
  const parts: string[] = [];
  const anchors: PlottedBarAnchor[] = [];

  const emitBar = (params: {
    readonly seriesIndex: number;
    readonly groupSeriesIndex: number;
    readonly seriesName: string;
    readonly pointIndex: number;
    readonly rawValue: number;
    readonly plottedValue?: number;
    readonly percent?: number;
    readonly bandStart: number;
    readonly bandThickness: number;
    readonly startCoord: number;
    readonly endCoord: number;
  }) => {
    const series = group.series[params.groupSeriesIndex]!;
    const point = series.points?.[params.pointIndex];
    const fill = resolveSeriesColor(
      point?.color ?? series.color,
      series.seriesIndex,
      theme,
    );
    const lineSpec = point?.line ?? series.line;
    const stroke =
      lineSpec === undefined
        ? undefined
        : resolveChartLineStroke(lineSpec, theme, { color: "dk1", width: 1 });

    const x = isHorizontal
      ? Math.min(params.startCoord, params.endCoord)
      : params.bandStart;
    const y = isHorizontal
      ? params.bandStart
      : Math.min(params.startCoord, params.endCoord);
    const width = isHorizontal
      ? Math.abs(params.endCoord - params.startCoord)
      : params.bandThickness;
    const height = isHorizontal
      ? params.bandThickness
      : Math.abs(params.endCoord - params.startCoord);

    const valueEndSide: PlottedBarAnchor["valueEndSide"] = isHorizontal
      ? params.endCoord >= params.startCoord
        ? "right"
        : "left"
      : params.endCoord <= params.startCoord
        ? "top"
        : "bottom";

    if (width > 0 && height > 0) {
      const strokeAttr = stroke === undefined ? "" : ` ${strokeAttrs(stroke)}`;
      parts.push(
        `<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="#${fill.hex}" fill-opacity="${fill.alpha ?? 1}"${strokeAttr}/>`,
      );
    } else if (stroke !== undefined) {
      parts.push(
        `<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="none" ${strokeAttrs(stroke)}/>`,
      );
    }

    anchors.push({
      kind: "bar",
      groupIndex,
      seriesIndex: params.seriesIndex,
      groupSeriesIndex: params.groupSeriesIndex,
      pointIndex: params.pointIndex,
      value: params.rawValue,
      ...(params.plottedValue === undefined
        ? {}
        : { plottedValue: params.plottedValue }),
      category: categories[params.pointIndex] ?? "",
      seriesName: params.seriesName,
      x,
      y,
      width,
      height,
      isNegative: params.rawValue < 0,
      horizontal: isHorizontal,
      valueEndSide,
      ...(params.percent === undefined ? {} : { percent: params.percent }),
    });
  };

  if (group.stacking === "clustered") {
    const o = (group.overlap ?? 0) / 100;
    const denom = Math.max(0.1, g + 1 + Math.max(0, seriesCount - 1) * (1 - o));
    const barThickness = Math.max(0.5, slotSpan / denom);
    const clusterSpan =
      barThickness * (1 + Math.max(0, seriesCount - 1) * (1 - o));
    const zeroCoord = valScale.zeroScreen(plotRect);

    for (let catIdx = 0; catIdx < categories.length; catIdx += 1) {
      const catCenter = catScale.indexToScreen(catIdx, plotRect);
      for (let s = 0; s < seriesCount; s += 1) {
        const series = group.series[s]!;
        const value = series.values[catIdx];
        if (value === null || value === undefined || !Number.isFinite(value)) {
          continue;
        }
        if (valScale.logBase !== undefined && value <= 0) {
          continue;
        }
        const bandStart =
          catCenter - clusterSpan / 2 + s * barThickness * (1 - o);
        const endCoord = valScale.valueToScreen(value, plotRect);
        emitBar({
          seriesIndex: series.seriesIndex,
          groupSeriesIndex: s,
          seriesName: series.name,
          pointIndex: catIdx,
          rawValue: value,
          bandStart,
          bandThickness: barThickness,
          startCoord: zeroCoord,
          endCoord,
        });
      }
    }
  } else {
    const barThickness = Math.max(0.5, slotSpan / Math.max(0.1, g + 1));
    const isPercent = group.stacking === "percent";

    for (let catIdx = 0; catIdx < categories.length; catIdx += 1) {
      const catCenter = catScale.indexToScreen(catIdx, plotRect);
      const bandStart = catCenter - barThickness / 2;

      let totalAbs = 0;
      if (isPercent) {
        for (const series of group.series) {
          const v = series.values[catIdx];
          if (v !== null && v !== undefined && Number.isFinite(v)) {
            totalAbs += Math.abs(v);
          }
        }
        if (totalAbs === 0) continue;
      }

      let posAcc = 0;
      let negAcc = 0;

      for (let s = 0; s < seriesCount; s += 1) {
        const series = group.series[s]!;
        const raw = series.values[catIdx];
        if (raw === null || raw === undefined || !Number.isFinite(raw)) {
          continue;
        }
        const plotted = isPercent ? raw / totalAbs : raw;
        if (valScale.logBase !== undefined && plotted <= 0) {
          continue;
        }
        let startVal: number;
        let endVal: number;
        if (plotted >= 0) {
          startVal = posAcc;
          endVal = posAcc + plotted;
          posAcc = endVal;
        } else {
          startVal = negAcc;
          endVal = negAcc + plotted;
          negAcc = endVal;
        }
        const startCoord = valScale.valueToScreen(
          valScale.logBase !== undefined
            ? Math.max(valScale.min, startVal)
            : startVal,
          plotRect,
        );
        const endCoord = valScale.valueToScreen(endVal, plotRect);
        emitBar({
          seriesIndex: series.seriesIndex,
          groupSeriesIndex: s,
          seriesName: series.name,
          pointIndex: catIdx,
          rawValue: raw,
          plottedValue: endVal,
          ...(isPercent ? { percent: Math.abs(raw) / totalAbs } : {}),
          bandStart,
          bandThickness: barThickness,
          startCoord,
          endCoord,
        });
      }
    }
  }

  return { svg: parts.join(""), anchors };
}
