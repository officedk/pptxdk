import type { GeometryBox, ThemeContext } from "@react-pptx/core/capability";

import type { TextStyle } from "../../shared/text-service.ts";
import type { ParsedChart, ParsedChartAxis } from "../chart-model.ts";
import {
  collectLegendEntries,
  measureLegend,
  type MeasuredLegend,
} from "./legend.ts";
import { formatChartNumber } from "./number-format.ts";
import { resolveAxisCrossingNormalized } from "./scales.ts";
import {
  measureChartText,
  mergeChartTextStyle,
  resolveChartFontSizePixels,
} from "./text-render.ts";
import type { AxisScale, ChartLayoutResult } from "./types.ts";

export type ComputedChartLayout = ChartLayoutResult & {
  readonly measuredLegend?: MeasuredLegend;
};

type EdgeSide = "top" | "bottom" | "left" | "right";

function titleTextOf(title: NonNullable<ParsedChart["title"]>): {
  readonly text: string;
  readonly style?: TextStyle;
} {
  return typeof title === "string" ? { text: title } : title;
}

export function resolveAxisLabelSide(
  axis: ParsedChartAxis,
  crossScale: AxisScale | undefined,
): EdgeSide | "interior" | "none" {
  const tickPos = axis.spec.tickLabelPosition ?? "nextTo";
  if (tickPos === "none") return "none";
  if (tickPos === "low") {
    return axis.direction === "horizontal" ? "bottom" : "left";
  }
  if (tickPos === "high") {
    return axis.direction === "horizontal" ? "top" : "right";
  }
  if (crossScale !== undefined) {
    const t = resolveAxisCrossingNormalized(axis, crossScale);
    if (t > 0.02 && t < 0.98) {
      return "interior";
    }
    if (axis.direction === "horizontal") {
      return t >= 0.98 ? "top" : "bottom";
    }
    return t >= 0.98 ? "right" : "left";
  }
  return axis.position;
}

export function resolveAxisTitleSide(
  axis: ParsedChartAxis,
  crossScale: AxisScale | undefined,
): EdgeSide {
  const labelSide = resolveAxisLabelSide(axis, crossScale);
  if (labelSide !== "none" && labelSide !== "interior") {
    return labelSide;
  }
  return axis.position;
}

function hasOutsideEndBarLabels(parsed: ParsedChart): {
  readonly columnPos: boolean;
  readonly columnNeg: boolean;
  readonly barPos: boolean;
  readonly barNeg: boolean;
} {
  let columnPos = false;
  let columnNeg = false;
  let barPos = false;
  let barNeg = false;
  for (const group of parsed.groups) {
    if (group.kind !== "column" && group.kind !== "bar") continue;
    for (const series of group.series) {
      const serLabels =
        series.dataLabels === false
          ? false
          : series.dataLabels !== undefined
            ? {
                ...(group.dataLabels === false ? {} : group.dataLabels),
                ...series.dataLabels,
              }
            : group.dataLabels;
      for (let i = 0; i < series.values.length; i += 1) {
        const val = series.values[i];
        if (val === null || val === undefined || !Number.isFinite(val))
          continue;
        const ptOverride = series.points?.[i]?.dataLabels;
        const eff =
          ptOverride === false
            ? false
            : ptOverride !== undefined
              ? { ...(serLabels === false ? {} : serLabels), ...ptOverride }
              : serLabels;
        if (eff === undefined || eff === false) continue;
        if (!eff.value && !eff.category && !eff.percent && !eff.seriesName)
          continue;
        const pos =
          eff.position ??
          (group.stacking === "clustered" ? "outsideEnd" : "center");
        if (pos !== "outsideEnd") continue;
        if (group.kind === "column") {
          if (val < 0) columnNeg = true;
          else columnPos = true;
        } else {
          if (val < 0) barNeg = true;
          else barPos = true;
        }
      }
    }
  }
  return { columnPos, columnNeg, barPos, barNeg };
}

function hasPointDataLabels(parsed: ParsedChart): {
  readonly top: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
  readonly right: boolean;
} {
  let top = false;
  let bottom = false;
  let left = false;
  let right = false;
  for (const group of parsed.groups) {
    if (
      group.kind !== "line" &&
      group.kind !== "area" &&
      group.kind !== "scatter"
    ) {
      continue;
    }
    const defaultPos =
      group.kind === "scatter"
        ? "right"
        : group.kind === "area"
          ? "center"
          : "top";
    for (const series of group.series) {
      const serLabels =
        series.dataLabels === false
          ? false
          : series.dataLabels !== undefined
            ? {
                ...(group.dataLabels === false ? {} : group.dataLabels),
                ...series.dataLabels,
              }
            : group.dataLabels;
      const count = "x" in series ? series.x.length : series.values.length;
      for (let i = 0; i < count; i += 1) {
        const ptOverride = series.points?.[i]?.dataLabels;
        const eff =
          ptOverride === false
            ? false
            : ptOverride !== undefined
              ? { ...(serLabels === false ? {} : serLabels), ...ptOverride }
              : serLabels;
        if (eff === undefined || eff === false) continue;
        if (
          !eff.value &&
          !eff.category &&
          !eff.percent &&
          !eff.xValue &&
          !eff.seriesName
        ) {
          continue;
        }
        const pos = eff.position ?? defaultPos;
        if (pos === "top") top = true;
        else if (pos === "bottom") bottom = true;
        else if (pos === "left") left = true;
        else if (pos === "right") right = true;
      }
    }
  }
  return { top, bottom, left, right };
}

export function computeChartLayout(options: {
  readonly parsed: ParsedChart;
  readonly box: GeometryBox;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly baseChartTextStyle: TextStyle | undefined;
  readonly baseFontSizePx: number;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
}): ComputedChartLayout {
  const {
    parsed,
    box,
    scales,
    baseChartTextStyle,
    baseFontSizePx,
    theme,
    slideNumber,
  } = options;

  const legendEntries =
    parsed.legend !== undefined && parsed.legend !== false
      ? collectLegendEntries(parsed, theme)
      : [];

  if (parsed.plotArea !== undefined) {
    const plotRect: GeometryBox = {
      x: box.x + parsed.plotArea.left,
      y: box.y + parsed.plotArea.top,
      width: parsed.plotArea.width,
      height: parsed.plotArea.height,
    };
    let titleBox: GeometryBox | undefined;
    if (parsed.title !== undefined) {
      const topRoom = Math.max(0, parsed.plotArea.top - 4);
      const t = titleTextOf(parsed.title);
      const measured = measureChartText({
        text: t.text,
        style: mergeChartTextStyle(baseChartTextStyle, t.style),
        theme,
        slideNumber,
        labelPath: "Chart.title",
        availableWidth: Math.max(10, box.width - 16),
      });
      const h = Math.min(measured.height, Math.max(12, topRoom));
      titleBox = {
        x: box.x + 8,
        y: box.y + 2,
        width: Math.max(10, box.width - 16),
        height: h,
      };
    }
    let measuredLegend: MeasuredLegend | undefined;
    let legendBox: GeometryBox | undefined;
    if (
      parsed.legend !== undefined &&
      parsed.legend !== false &&
      legendEntries.length > 0
    ) {
      measuredLegend = measureLegend({
        entries: legendEntries,
        position: parsed.legend,
        maxWidth: Math.max(20, box.width - 16),
        fontSizePx: baseFontSizePx,
        style: baseChartTextStyle,
        theme,
        slideNumber,
      });
      const bottomRoom = Math.max(
        14,
        box.height - (parsed.plotArea.top + parsed.plotArea.height) - 4,
      );
      if (parsed.legend === "bottom") {
        legendBox = {
          x: box.x + 8,
          y: box.y + box.height - bottomRoom - 2,
          width: Math.max(10, box.width - 16),
          height: bottomRoom,
        };
      } else if (parsed.legend === "top") {
        legendBox = {
          x: box.x + 8,
          y: box.y + (titleBox ? titleBox.height + 4 : 2),
          width: Math.max(10, box.width - 16),
          height: Math.min(
            measuredLegend.height,
            Math.max(14, parsed.plotArea.top - 4),
          ),
        };
      } else if (parsed.legend === "left") {
        legendBox = {
          x: box.x + 4,
          y: plotRect.y,
          width: Math.max(14, parsed.plotArea.left - 8),
          height: plotRect.height,
        };
      } else {
        const rightRoom = Math.max(
          14,
          box.width - (parsed.plotArea.left + parsed.plotArea.width) - 8,
        );
        legendBox = {
          x: plotRect.x + plotRect.width + 4,
          y: plotRect.y,
          width: rightRoom,
          height:
            parsed.legend === "topRight"
              ? Math.min(measuredLegend.height, plotRect.height)
              : plotRect.height,
        };
      }
    }

    const axisTitleBoxes = new Map<string, GeometryBox>();
    const labelZones = new Map<
      string,
      {
        readonly side: EdgeSide;
        readonly depth: number;
        readonly offset: number;
      }
    >();
    for (const axis of parsed.axes) {
      if (axis.spec.visible === false) continue;
      const crossAxis = parsed.axes.find((a) => a.id === axis.crossAxisId);
      const crossScale = crossAxis ? scales.get(crossAxis.name) : undefined;
      const labelSide = resolveAxisLabelSide(axis, crossScale);
      if (labelSide !== "none" && labelSide !== "interior") {
        labelZones.set(axis.name, { side: labelSide, depth: 20, offset: 0 });
      }
      if (axis.spec.title !== undefined) {
        const titleSide = resolveAxisTitleSide(axis, crossScale);
        const t = titleTextOf(axis.spec.title);
        const titleStyle = mergeChartTextStyle(
          mergeChartTextStyle(baseChartTextStyle, axis.spec.style),
          t.style,
        );
        const titlePx = resolveChartFontSizePixels(titleStyle?.fontSize);
        const thickness = titlePx * 1.3;
        if (titleSide === "bottom") {
          axisTitleBoxes.set(axis.name, {
            x: plotRect.x,
            y: Math.min(
              box.y + box.height - thickness,
              plotRect.y + plotRect.height + 18,
            ),
            width: plotRect.width,
            height: thickness,
          });
        } else if (titleSide === "top") {
          axisTitleBoxes.set(axis.name, {
            x: plotRect.x,
            y: Math.max(box.y, plotRect.y - 18 - thickness),
            width: plotRect.width,
            height: thickness,
          });
        } else if (titleSide === "left") {
          axisTitleBoxes.set(axis.name, {
            x: Math.max(box.x, plotRect.x - 24 - thickness),
            y: plotRect.y,
            width: thickness,
            height: plotRect.height,
          });
        } else {
          axisTitleBoxes.set(axis.name, {
            x: Math.min(
              box.x + box.width - thickness,
              plotRect.x + plotRect.width + 24,
            ),
            y: plotRect.y,
            width: thickness,
            height: plotRect.height,
          });
        }
      }
    }

    return {
      plotRect,
      ...(titleBox === undefined ? {} : { titleBox }),
      ...(legendBox === undefined ? {} : { legendBox }),
      ...(measuredLegend === undefined ? {} : { measuredLegend }),
      axisTitleBoxes,
      labelZones,
    };
  }

  // Automatic layout
  const padX = Math.min(8, box.width * 0.025);
  const padY = Math.min(8, box.height * 0.025);
  const avail = {
    x: box.x + padX,
    y: box.y + padY,
    width: Math.max(4, box.width - 2 * padX),
    height: Math.max(4, box.height - 2 * padY),
  };

  let titleBox: GeometryBox | undefined;
  if (parsed.title !== undefined) {
    const t = titleTextOf(parsed.title);
    const titleStyle = mergeChartTextStyle(baseChartTextStyle, t.style);
    const measured = measureChartText({
      text: t.text,
      style: titleStyle,
      theme,
      slideNumber,
      labelPath: "Chart.title",
      availableWidth: avail.width,
    });
    const gap = Math.min(6, avail.height * 0.03);
    const maxTitleH = Math.max(4, avail.height * 0.25);
    const h = Math.min(measured.height, maxTitleH);
    titleBox = {
      x: avail.x,
      y: avail.y,
      width: avail.width,
      height: h,
    };
    const consume = Math.min(h + gap, Math.max(0, avail.height - 4));
    avail.y += consume;
    avail.height = Math.max(4, avail.height - consume);
  }

  let measuredLegend: MeasuredLegend | undefined;
  let legendBox: GeometryBox | undefined;
  if (
    parsed.legend !== undefined &&
    parsed.legend !== false &&
    legendEntries.length > 0
  ) {
    const pos = parsed.legend;
    measuredLegend = measureLegend({
      entries: legendEntries,
      position: pos,
      maxWidth: avail.width,
      fontSizePx: baseFontSizePx,
      style: baseChartTextStyle,
      theme,
      slideNumber,
    });
    const gap = Math.min(6, Math.min(avail.width, avail.height) * 0.03);
    if (pos === "top" || pos === "bottom") {
      const h = Math.min(
        measuredLegend.height,
        Math.max(4, avail.height * 0.28),
      );
      const consume = Math.min(h + gap, Math.max(0, avail.height - 4));
      if (pos === "top") {
        legendBox = { x: avail.x, y: avail.y, width: avail.width, height: h };
        avail.y += consume;
      } else {
        legendBox = {
          x: avail.x,
          y: avail.y + avail.height - h,
          width: avail.width,
          height: h,
        };
      }
      avail.height = Math.max(4, avail.height - consume);
    } else if (pos === "left" || pos === "right") {
      const w = Math.min(measuredLegend.width, Math.max(4, avail.width * 0.32));
      const consume = Math.min(w + gap, Math.max(0, avail.width - 4));
      if (pos === "left") {
        legendBox = { x: avail.x, y: avail.y, width: w, height: avail.height };
        avail.x += consume;
      } else {
        legendBox = {
          x: avail.x + avail.width - w,
          y: avail.y,
          width: w,
          height: avail.height,
        };
      }
      avail.width = Math.max(4, avail.width - consume);
    } else {
      // topRight
      const w = Math.min(measuredLegend.width, Math.max(4, avail.width * 0.32));
      const h = Math.min(
        measuredLegend.height,
        Math.max(4, avail.height * 0.5),
      );
      const consume = Math.min(w + gap, Math.max(0, avail.width - 4));
      legendBox = {
        x: avail.x + avail.width - w,
        y: avail.y,
        width: w,
        height: h,
      };
      avail.width = Math.max(4, avail.width - consume);
    }
  }

  // Reserve edge space for axes ticks, tick marks, and axis titles.
  const edges: Record<EdgeSide, number> = {
    top: 0,
    bottom: 0,
    left: 0,
    right: 0,
  };
  const labelDepths = new Map<
    string,
    { readonly side: EdgeSide; readonly depth: number; readonly offset: number }
  >();
  const addLabelZone = (
    axisName: string,
    side: EdgeSide,
    depth: number,
  ): void => {
    const existing = edges[side];
    const offset = existing > 0 ? existing + 2 : 0;
    edges[side] = offset + depth;
    labelDepths.set(axisName, { side, depth, offset });
  };
  const axisTitleSpecs: {
    readonly axis: ParsedChartAxis;
    readonly side: EdgeSide;
    readonly thickness: number;
  }[] = [];

  for (const axis of parsed.axes) {
    if (axis.spec.visible === false) continue;
    const scale = scales.get(axis.name);
    if (scale === undefined) continue;
    const crossAxis = parsed.axes.find((a) => a.id === axis.crossAxisId);
    const crossScale = crossAxis ? scales.get(crossAxis.name) : undefined;
    const labelSide = resolveAxisLabelSide(axis, crossScale);
    const axisTextStyle = mergeChartTextStyle(
      baseChartTextStyle,
      axis.spec.style,
    );
    const tickProtrusion =
      axis.spec.majorTickMark === "out" || axis.spec.majorTickMark === "cross"
        ? 4
        : 0;
    const labelGap = 4;

    if (labelSide !== "none" && labelSide !== "interior") {
      if (scale.kind === "value") {
        const formattedTicks = scale.majorTicks.map((v) =>
          formatChartNumber(v, scale.numberFormat),
        );
        let maxW = 0;
        let maxH = 0;
        for (const text of formattedTicks) {
          const m = measureChartText({
            text,
            style: axisTextStyle,
            theme,
            slideNumber,
            labelPath: `Chart.axes.${axis.name}.style`,
          });
          maxW = Math.max(maxW, m.width);
          maxH = Math.max(maxH, m.height);
        }
        if (axis.direction === "horizontal") {
          const depth = maxH + labelGap + tickProtrusion;
          addLabelZone(axis.name, labelSide, depth);
          const overhang = Math.min(maxW / 2, avail.width * 0.1);
          edges.left = Math.max(edges.left, overhang);
          edges.right = Math.max(edges.right, overhang);
        } else {
          const depth = Math.min(
            maxW + labelGap + tickProtrusion,
            avail.width * 0.32,
          );
          addLabelZone(axis.name, labelSide, depth);
          const overhang = Math.min(maxH / 2, avail.height * 0.1);
          edges.top = Math.max(edges.top, overhang);
          edges.bottom = Math.max(edges.bottom, overhang);
        }
      } else {
        let maxW = 0;
        let maxH = 0;
        const estSlot = Math.max(12, (avail.width * 0.8) / scale.count);
        for (const cat of scale.categories) {
          const m = measureChartText({
            text: cat,
            style: axisTextStyle,
            theme,
            slideNumber,
            labelPath: `Chart.axes.${axis.name}.style`,
            ...(axis.direction === "horizontal"
              ? { availableWidth: estSlot }
              : {}),
          });
          maxW = Math.max(maxW, m.width);
          maxH = Math.max(maxH, m.height);
        }
        if (axis.direction === "horizontal") {
          const singleLineH = measureChartText({
            text: "Ay",
            style: axisTextStyle,
            theme,
            slideNumber,
            labelPath: `Chart.axes.${axis.name}.style`,
          }).height;
          const cappedH = Math.min(
            maxH,
            singleLineH * 2.2,
            avail.height * 0.25,
          );
          const depth = cappedH + labelGap + tickProtrusion;
          addLabelZone(axis.name, labelSide, depth);
        } else {
          const depth = Math.min(
            maxW + labelGap + tickProtrusion,
            avail.width * 0.34,
          );
          addLabelZone(axis.name, labelSide, depth);
        }
      }
    }

    if (axis.spec.title !== undefined) {
      const titleSide = resolveAxisTitleSide(axis, crossScale);
      const t = titleTextOf(axis.spec.title);
      const titleStyle = mergeChartTextStyle(axisTextStyle, t.style);
      const m = measureChartText({
        text: t.text,
        style: titleStyle,
        theme,
        slideNumber,
        labelPath: `Chart.axes.${axis.name}.title`,
      });
      const thickness = m.height + 4;
      axisTitleSpecs.push({ axis, side: titleSide, thickness });
    }
  }

  // Reserve outer band for axis titles beyond label bands.
  const axisTitleBoxes = new Map<string, GeometryBox>();
  const titleEdgeOffset: Record<EdgeSide, number> = {
    top: edges.top,
    bottom: edges.bottom,
    left: edges.left,
    right: edges.right,
  };
  for (const item of axisTitleSpecs) {
    edges[item.side] += item.thickness;
  }

  // Reserve space for outsideEnd bar labels and edge-adjacent point data labels.
  const outside = hasOutsideEndBarLabels(parsed);
  const pointLabels = hasPointDataLabels(parsed);
  if (outside.columnPos || pointLabels.top) {
    edges.top = Math.max(
      edges.top,
      Math.min(baseFontSizePx * 1.4, avail.height * 0.16),
    );
  }
  if (outside.columnNeg || pointLabels.bottom) {
    edges.bottom = Math.max(
      edges.bottom,
      Math.min(baseFontSizePx * 1.4, avail.height * 0.16),
    );
  }
  if (outside.barPos || pointLabels.right) {
    edges.right = Math.max(
      edges.right,
      Math.min(baseFontSizePx * 2.5, avail.width * 0.16),
    );
  }
  if (outside.barNeg || pointLabels.left) {
    edges.left = Math.max(
      edges.left,
      Math.min(baseFontSizePx * 2.5, avail.width * 0.16),
    );
  }

  // Ensure a small visual breathing room on edges without labels/titles.
  for (const side of ["top", "bottom", "left", "right"] as const) {
    if (edges[side] === 0) {
      edges[side] = Math.min(
        6,
        (side === "left" || side === "right" ? avail.width : avail.height) *
          0.03,
      );
    }
  }

  // Proportional safety cap for compact frames so plotRect always has positive dimensions.
  const maxHorizSum = Math.max(0, avail.width - 4);
  if (edges.left + edges.right > maxHorizSum && edges.left + edges.right > 0) {
    const ratio = maxHorizSum / (edges.left + edges.right);
    edges.left *= ratio;
    edges.right *= ratio;
    titleEdgeOffset.left *= ratio;
    titleEdgeOffset.right *= ratio;
  }
  const maxVertSum = Math.max(0, avail.height - 4);
  if (edges.top + edges.bottom > maxVertSum && edges.top + edges.bottom > 0) {
    const ratio = maxVertSum / (edges.top + edges.bottom);
    edges.top *= ratio;
    edges.bottom *= ratio;
    titleEdgeOffset.top *= ratio;
    titleEdgeOffset.bottom *= ratio;
  }

  const plotRect: GeometryBox = {
    x: avail.x + edges.left,
    y: avail.y + edges.top,
    width: Math.max(4, avail.width - edges.left - edges.right),
    height: Math.max(4, avail.height - edges.top - edges.bottom),
  };

  // Position axis title boxes around plotRect.
  const titleCursors: Record<EdgeSide, number> = {
    top: plotRect.y - titleEdgeOffset.top,
    bottom: plotRect.y + plotRect.height + titleEdgeOffset.bottom,
    left: plotRect.x - titleEdgeOffset.left,
    right: plotRect.x + plotRect.width + titleEdgeOffset.right,
  };
  for (const item of axisTitleSpecs) {
    const hRatio =
      avail.width > 4 && edges.left + edges.right > 0
        ? Math.min(1, maxHorizSum / (edges.left + edges.right))
        : 1;
    const vRatio =
      avail.height > 4 && edges.top + edges.bottom > 0
        ? Math.min(1, maxVertSum / (edges.top + edges.bottom))
        : 1;
    if (item.side === "top") {
      const th = item.thickness * vRatio;
      const y = titleCursors.top - th;
      axisTitleBoxes.set(item.axis.name, {
        x: plotRect.x,
        y,
        width: plotRect.width,
        height: th,
      });
      titleCursors.top = y;
    } else if (item.side === "bottom") {
      const th = item.thickness * vRatio;
      const y = titleCursors.bottom;
      axisTitleBoxes.set(item.axis.name, {
        x: plotRect.x,
        y,
        width: plotRect.width,
        height: th,
      });
      titleCursors.bottom = y + th;
    } else if (item.side === "left") {
      const tw = item.thickness * hRatio;
      const x = titleCursors.left - tw;
      axisTitleBoxes.set(item.axis.name, {
        x,
        y: plotRect.y,
        width: tw,
        height: plotRect.height,
      });
      titleCursors.left = x;
    } else {
      const tw = item.thickness * hRatio;
      const x = titleCursors.right;
      axisTitleBoxes.set(item.axis.name, {
        x,
        y: plotRect.y,
        width: tw,
        height: plotRect.height,
      });
      titleCursors.right = x + tw;
    }
  }

  return {
    plotRect,
    ...(titleBox === undefined ? {} : { titleBox }),
    ...(legendBox === undefined ? {} : { legendBox }),
    ...(measuredLegend === undefined ? {} : { measuredLegend }),
    axisTitleBoxes,
    labelZones: labelDepths,
  };
}
