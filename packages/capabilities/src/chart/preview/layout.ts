import type { Rect, ThemeContext } from "@pptxdk/core/capability";

import type { TextStyle } from "../../shared/text-service.ts";
import { mergeDataLabels } from "../chart-labels.ts";
import type { ParsedChart, ParsedChartAxis } from "../chart-model.ts";
import type { ChartTextStyle } from "../chart-schemas.ts";
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
  readonly style?: ChartTextStyle | undefined;
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
      const serLabels = mergeDataLabels(group.dataLabels, series.dataLabels);
      for (let i = 0; i < series.values.length; i += 1) {
        const val = series.values[i];
        if (val === null || val === undefined || !Number.isFinite(val))
          continue;
        const ptOverride = series.points?.[i]?.dataLabels;
        const eff = mergeDataLabels(serLabels, ptOverride);
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
      const serLabels = mergeDataLabels(group.dataLabels, series.dataLabels);
      const count = "x" in series ? series.x.length : series.values.length;
      for (let i = 0; i < count; i += 1) {
        const ptOverride = series.points?.[i]?.dataLabels;
        const eff = mergeDataLabels(serLabels, ptOverride);
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
  readonly box: Rect;
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
    const plotRect: Rect = {
      x: box.x + parsed.plotArea.left,
      y: box.y + parsed.plotArea.top,
      width: parsed.plotArea.width,
      height: parsed.plotArea.height,
    };
    let titleBox: Rect | undefined;
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
    let legendBox: Rect | undefined;
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

    const axisTitleBoxes = new Map<string, Rect>();
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

  let titleBox: Rect | undefined;
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
  let legendBox: Rect | undefined;
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
  const AXIS_TITLE_GAP = 6;
  const AXIS_TITLE_BOX_PAD = 4;
  const VERTICAL_VALUE_LABEL_PAD = 4;
  const edges: Record<EdgeSide, number> = {
    top: 0,
    bottom: 0,
    left: 0,
    right: 0,
  };
  const overhang: Record<EdgeSide, number> = {
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
  const pendingTitles: {
    readonly axis: ParsedChartAxis;
    readonly side: EdgeSide;
    readonly text: string;
    readonly style: TextStyle | undefined;
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
    const effectiveFontPx =
      axisTextStyle?.fontSize !== undefined
        ? resolveChartFontSizePixels(axisTextStyle.fontSize)
        : baseFontSizePx;
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
        let firstW = 0;
        let lastW = 0;
        for (let i = 0; i < formattedTicks.length; i += 1) {
          const text = formattedTicks[i]!;
          const m = measureChartText({
            text,
            style: axisTextStyle,
            theme,
            slideNumber,
            labelPath: `Chart.axes.${axis.name}.style`,
          });
          if (i === 0) firstW = m.width;
          if (i === formattedTicks.length - 1) lastW = m.width;
          maxW = Math.max(maxW, m.width);
          maxH = Math.max(maxH, m.height);
        }
        if (axis.direction === "horizontal") {
          const depth = Math.ceil(maxH) + labelGap + tickProtrusion;
          addLabelZone(axis.name, labelSide, depth);
          const leftTickW = scale.reversed ? lastW : firstW;
          const rightTickW = scale.reversed ? firstW : lastW;
          overhang.left = Math.max(
            overhang.left,
            Math.min(Math.ceil(leftTickW / 2) + 2, avail.width * 0.15),
          );
          overhang.right = Math.max(
            overhang.right,
            Math.min(Math.ceil(rightTickW / 2) + 2, avail.width * 0.15),
          );
        } else {
          const depth = Math.min(
            Math.ceil(maxW) +
              VERTICAL_VALUE_LABEL_PAD +
              labelGap +
              tickProtrusion,
            avail.width * 0.42,
          );
          addLabelZone(axis.name, labelSide, depth);
          const vOverhang = Math.min(
            Math.ceil(Math.max(maxH, effectiveFontPx * 1.4) / 2) + 2,
            avail.height * 0.15,
          );
          overhang.top = Math.max(overhang.top, vOverhang);
          overhang.bottom = Math.max(overhang.bottom, vOverhang);
        }
      } else {
        let maxW = 0;
        let maxH = 0;
        let firstW = 0;
        let lastW = 0;
        const estSlot = Math.max(12, (avail.width * 0.8) / scale.count);
        for (let i = 0; i < scale.categories.length; i += 1) {
          const cat = scale.categories[i]!;
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
          if (i === 0) firstW = m.width;
          if (i === scale.categories.length - 1) lastW = m.width;
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
          if (scale.crossBetween === "midCat" && scale.categories.length > 1) {
            const leftCatW = scale.reversed ? lastW : firstW;
            const rightCatW = scale.reversed ? firstW : lastW;
            overhang.left = Math.max(
              overhang.left,
              Math.min(Math.ceil(leftCatW / 2) + 2, avail.width * 0.15),
            );
            overhang.right = Math.max(
              overhang.right,
              Math.min(Math.ceil(rightCatW / 2) + 2, avail.width * 0.15),
            );
          }
        } else {
          const depth = Math.min(
            Math.ceil(maxW) +
              VERTICAL_VALUE_LABEL_PAD +
              labelGap +
              tickProtrusion,
            avail.width * 0.34,
          );
          addLabelZone(axis.name, labelSide, depth);
          if (scale.crossBetween === "midCat" && scale.categories.length > 1) {
            const vOverhang = Math.min(
              Math.ceil(maxH / 2) + 2,
              avail.height * 0.15,
            );
            overhang.top = Math.max(overhang.top, vOverhang);
            overhang.bottom = Math.max(overhang.bottom, vOverhang);
          }
        }
      }
    } else if (tickProtrusion > 0) {
      overhang[axis.position] = Math.max(
        overhang[axis.position],
        tickProtrusion,
      );
    }

    if (axis.spec.title !== undefined) {
      const titleSide = resolveAxisTitleSide(axis, crossScale);
      const t = titleTextOf(axis.spec.title);
      const titleStyle = mergeChartTextStyle(axisTextStyle, t.style);
      pendingTitles.push({
        axis,
        side: titleSide,
        text: t.text,
        style: titleStyle,
      });
    }
  }

  // Merge cross-axis end-label overhangs after all label zone offsets are fixed,
  // so perpendicular overhangs never shift labelZones into axis titles or frame edges.
  for (const side of ["top", "bottom", "left", "right"] as const) {
    edges[side] = Math.max(edges[side], overhang[side]);
  }

  // Reserve outer band for axis titles beyond label bands.
  // Converge available length in up to 3 passes so wrapped horizontal/vertical titles
  // reflect the resulting plotRect width and height.
  const axisTitleBoxes = new Map<string, Rect>();
  const titleEdgeOffset: Record<EdgeSide, number> = {
    top: edges.top,
    bottom: edges.bottom,
    left: edges.left,
    right: edges.right,
  };
  let axisTitleSpecs: {
    readonly axis: ParsedChartAxis;
    readonly side: EdgeSide;
    readonly thickness: number;
  }[] = [];

  if (pendingTitles.length > 0) {
    let estEdges = { ...titleEdgeOffset };
    for (let pass = 0; pass < 3; pass += 1) {
      const estPlotW = Math.max(
        24,
        avail.width - estEdges.left - estEdges.right,
      );
      const estPlotH = Math.max(
        24,
        avail.height - estEdges.top - estEdges.bottom,
      );
      const nextSpecs = pendingTitles.map((item) => {
        const availableWidth =
          item.side === "top" || item.side === "bottom" ? estPlotW : estPlotH;
        const m = measureChartText({
          text: item.text,
          style: item.style,
          theme,
          slideNumber,
          labelPath: `Chart.axes.${item.axis.name}.title`,
          availableWidth,
        });
        return {
          axis: item.axis,
          side: item.side,
          thickness: Math.ceil(m.height) + AXIS_TITLE_BOX_PAD,
        };
      });
      const nextEdges = { ...titleEdgeOffset };
      for (const item of nextSpecs) {
        nextEdges[item.side] += AXIS_TITLE_GAP + item.thickness;
      }
      const converged =
        axisTitleSpecs.length === nextSpecs.length &&
        nextSpecs.every(
          (spec, idx) => spec.thickness === axisTitleSpecs[idx]?.thickness,
        );
      axisTitleSpecs = nextSpecs;
      estEdges = nextEdges;
      if (converged) break;
    }
    for (const side of ["top", "bottom", "left", "right"] as const) {
      edges[side] = estEdges[side];
    }
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
  const rawHorizSum = edges.left + edges.right;
  const hRatio =
    rawHorizSum > maxHorizSum && rawHorizSum > 0
      ? maxHorizSum / rawHorizSum
      : 1;
  if (hRatio < 1) {
    edges.left *= hRatio;
    edges.right *= hRatio;
    titleEdgeOffset.left *= hRatio;
    titleEdgeOffset.right *= hRatio;
  }
  const maxVertSum = Math.max(0, avail.height - 4);
  const rawVertSum = edges.top + edges.bottom;
  const vRatio =
    rawVertSum > maxVertSum && rawVertSum > 0 ? maxVertSum / rawVertSum : 1;
  if (vRatio < 1) {
    edges.top *= vRatio;
    edges.bottom *= vRatio;
    titleEdgeOffset.top *= vRatio;
    titleEdgeOffset.bottom *= vRatio;
  }

  const plotRect: Rect = {
    x: avail.x + edges.left,
    y: avail.y + edges.top,
    width: Math.max(4, avail.width - edges.left - edges.right),
    height: Math.max(4, avail.height - edges.top - edges.bottom),
  };

  // Position axis title boxes around plotRect with explicit clearance from label bands.
  const titleCursors: Record<EdgeSide, number> = {
    top: plotRect.y - titleEdgeOffset.top,
    bottom: plotRect.y + plotRect.height + titleEdgeOffset.bottom,
    left: plotRect.x - titleEdgeOffset.left,
    right: plotRect.x + plotRect.width + titleEdgeOffset.right,
  };
  const hTitleGap = AXIS_TITLE_GAP * hRatio;
  const vTitleGap = AXIS_TITLE_GAP * vRatio;
  for (const item of axisTitleSpecs) {
    if (item.side === "top") {
      const th = item.thickness * vRatio;
      const y = titleCursors.top - vTitleGap - th;
      axisTitleBoxes.set(item.axis.name, {
        x: plotRect.x,
        y,
        width: plotRect.width,
        height: th,
      });
      titleCursors.top = y;
    } else if (item.side === "bottom") {
      const th = item.thickness * vRatio;
      const y = titleCursors.bottom + vTitleGap;
      axisTitleBoxes.set(item.axis.name, {
        x: plotRect.x,
        y,
        width: plotRect.width,
        height: th,
      });
      titleCursors.bottom = y + th;
    } else if (item.side === "left") {
      const tw = item.thickness * hRatio;
      const x = titleCursors.left - hTitleGap - tw;
      axisTitleBoxes.set(item.axis.name, {
        x,
        y: plotRect.y,
        width: tw,
        height: plotRect.height,
      });
      titleCursors.left = x;
    } else {
      const tw = item.thickness * hRatio;
      const x = titleCursors.right + hTitleGap;
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
