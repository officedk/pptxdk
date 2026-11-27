import type { Rect, ThemeContext } from "@pptxdk/core/capability";

import type { TextStyle } from "../../shared/text-service.ts";
import type { ParsedChart, ParsedChartAxis } from "../chart-model.ts";
import type { ChartTextStyle } from "../chart-schemas.ts";
import type { ChartLine, ChartTickMark } from "../chart-types.ts";
import { resolveAxisLabelSide, type ComputedChartLayout } from "./layout.ts";
import { formatChartNumber } from "./number-format.ts";
import { resolveAxisCrossingScreen } from "./scales.ts";
import {
  resolveChartLineStroke,
  strokeAttrs,
  type ResolvedStroke,
} from "./style.ts";
import {
  mergeChartTextStyle,
  renderChartTextSvg,
  resolveChartFontSizePixels,
  type ChartTextClipAllocator,
} from "./text-render.ts";
import {
  assertNever,
  type AxisScale,
  type CategoryAxisScale,
} from "./types.ts";

function categoryGridCoordinates(
  scale: CategoryAxisScale,
  plot: Rect,
  minor: boolean,
): number[] {
  const coords: number[] = [];
  if (minor) {
    if (scale.crossBetween === "midCat" && scale.count > 1) {
      for (let i = 0; i < scale.count - 1; i += 1) {
        const rawT = (i + 0.5) / (scale.count - 1);
        coords.push(scale.mapScreen(scale.reversed ? 1 - rawT : rawT, plot));
      }
    } else {
      for (let i = 0; i < scale.count; i += 1) {
        coords.push(scale.indexToScreen(i, plot));
      }
    }
    return coords;
  }
  if (scale.crossBetween === "midCat") {
    for (let i = 0; i < scale.count; i += 1) {
      coords.push(scale.indexToScreen(i, plot));
    }
  } else {
    for (let b = 0; b <= scale.count; b += 1) {
      coords.push(scale.mapScreen(scale.mapBoundary(b), plot));
    }
  }
  return coords;
}

function drawGridlineSet(
  coords: readonly number[],
  direction: "horizontal" | "vertical",
  plot: Rect,
  stroke: ResolvedStroke,
): string {
  if (coords.length === 0) return "";
  const segs: string[] = [];
  for (const coord of coords) {
    if (direction === "vertical") {
      segs.push(`M${plot.x} ${coord}H${plot.x + plot.width}`);
    } else {
      segs.push(`M${coord} ${plot.y}V${plot.y + plot.height}`);
    }
  }
  return `<path d="${segs.join("")}" ${strokeAttrs(stroke)} fill="none"/>`;
}

export function renderGridlinesSvg(options: {
  readonly parsed: ParsedChart;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly plotRect: Rect;
  readonly theme: ThemeContext;
}): string {
  const { parsed, scales, plotRect, theme } = options;
  const parts: string[] = [];

  for (const axis of parsed.axes) {
    if (axis.spec.visible === false) continue;
    const scale = scales.get(axis.name);
    if (scale === undefined) continue;

    const emitFor = (
      setting: boolean | ChartLine | undefined,
      minor: boolean,
    ) => {
      if (setting === undefined || setting === false) return;
      const stroke = resolveChartLineStroke(setting, theme, {
        color: "dk1",
        width: minor ? 0.75 : 1,
        alpha: minor ? 0.1 : 0.2,
      });
      if (stroke === undefined) return;
      const coords =
        scale.kind === "value"
          ? (minor ? scale.minorTicks : scale.majorTicks).map((v) =>
              scale.valueToScreen(v, plotRect),
            )
          : categoryGridCoordinates(scale, plotRect, minor);
      const svg = drawGridlineSet(coords, axis.direction, plotRect, stroke);
      if (svg !== "") parts.push(svg);
    };

    emitFor(axis.spec.minorGridlines, true);
    emitFor(axis.spec.gridlines, false);
  }

  return parts.join("");
}

function tickMarkRange(
  mark: ChartTickMark | undefined,
  length: number,
  outDir: -1 | 1,
): { readonly startOffset: number; readonly endOffset: number } | undefined {
  const effective = mark ?? "none";
  switch (effective) {
    case "none":
      return undefined;
    case "out":
      return outDir > 0
        ? { startOffset: 0, endOffset: length }
        : { startOffset: -length, endOffset: 0 };
    case "in":
      return outDir > 0
        ? { startOffset: -length, endOffset: 0 }
        : { startOffset: 0, endOffset: length };
    case "cross":
      return { startOffset: -length, endOffset: length };
    default:
      return assertNever(effective);
  }
}

function renderTickMarks(options: {
  readonly coords: readonly number[];
  readonly direction: "horizontal" | "vertical";
  readonly crossCoord: number;
  readonly mark: ChartTickMark | undefined;
  readonly length: number;
  readonly outDir: -1 | 1;
  readonly stroke: ResolvedStroke;
}): string {
  const range = tickMarkRange(options.mark, options.length, options.outDir);
  if (range === undefined || options.coords.length === 0) return "";
  const segs: string[] = [];
  for (const coord of options.coords) {
    if (options.direction === "horizontal") {
      segs.push(
        `M${coord} ${options.crossCoord + range.startOffset}V${options.crossCoord + range.endOffset}`,
      );
    } else {
      segs.push(
        `M${options.crossCoord + range.startOffset} ${coord}H${options.crossCoord + range.endOffset}`,
      );
    }
  }
  return `<path d="${segs.join("")}" ${strokeAttrs(options.stroke)} fill="none"/>`;
}

function axisTitleParts(
  axis: ParsedChartAxis,
):
  | { readonly text: string; readonly style?: ChartTextStyle | undefined }
  | undefined {
  if (axis.spec.title === undefined) return undefined;
  return typeof axis.spec.title === "string"
    ? { text: axis.spec.title }
    : axis.spec.title;
}

export function renderAxesSvg(options: {
  readonly parsed: ParsedChart;
  readonly chartBox: Rect;
  readonly layout: ComputedChartLayout;
  readonly scales: ReadonlyMap<string, AxisScale>;
  readonly baseChartTextStyle: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly clipAllocator: ChartTextClipAllocator;
}): string {
  const {
    parsed,
    chartBox,
    layout,
    scales,
    baseChartTextStyle,
    theme,
    slideNumber,
    clipAllocator,
  } = options;
  const { plotRect } = layout;
  const parts: string[] = [];

  for (const axis of parsed.axes) {
    if (axis.spec.visible === false) continue;
    const scale = scales.get(axis.name);
    if (scale === undefined) continue;
    const crossAxis = parsed.axes.find((a) => a.id === axis.crossAxisId);
    const crossScale = crossAxis ? scales.get(crossAxis.name) : undefined;
    const crossCoord =
      crossScale !== undefined
        ? resolveAxisCrossingScreen(axis, crossScale, plotRect)
        : axis.position === "top"
          ? plotRect.y
          : axis.position === "bottom"
            ? plotRect.y + plotRect.height
            : axis.position === "right"
              ? plotRect.x + plotRect.width
              : plotRect.x;

    const lineStroke = resolveChartLineStroke(axis.spec.line, theme, {
      color: "dk1",
      width: 1,
    });
    const tickStroke =
      lineStroke ??
      resolveChartLineStroke(true, theme, { color: "dk1", width: 1 })!;

    if (lineStroke !== undefined) {
      const path =
        axis.direction === "horizontal"
          ? `M${plotRect.x} ${crossCoord}H${plotRect.x + plotRect.width}`
          : `M${crossCoord} ${plotRect.y}V${plotRect.y + plotRect.height}`;
      parts.push(`<path d="${path}" ${strokeAttrs(lineStroke)} fill="none"/>`);
    }

    const outDir: -1 | 1 =
      axis.direction === "horizontal"
        ? axis.position === "top" || crossCoord <= plotRect.y + 1
          ? -1
          : 1
        : axis.position === "right" ||
            crossCoord >= plotRect.x + plotRect.width - 1
          ? 1
          : -1;

    const majorTickCoords =
      scale.kind === "value"
        ? scale.majorTicks.map((v) => scale.valueToScreen(v, plotRect))
        : categoryGridCoordinates(scale, plotRect, false);
    const minorTickCoords =
      scale.kind === "value"
        ? scale.minorTicks.map((v) => scale.valueToScreen(v, plotRect))
        : categoryGridCoordinates(scale, plotRect, true);

    parts.push(
      renderTickMarks({
        coords: majorTickCoords,
        direction: axis.direction,
        crossCoord,
        mark: axis.spec.majorTickMark,
        length: 4,
        outDir,
        stroke: tickStroke,
      }),
    );
    parts.push(
      renderTickMarks({
        coords: minorTickCoords,
        direction: axis.direction,
        crossCoord,
        mark: axis.spec.minorTickMark,
        length: 2.5,
        outDir,
        stroke: tickStroke,
      }),
    );

    const axisTextStyle = mergeChartTextStyle(
      baseChartTextStyle,
      axis.spec.style,
    );
    const axisFontPx = resolveChartFontSizePixels(axisTextStyle?.fontSize);
    const labelSide = resolveAxisLabelSide(axis, crossScale);
    const tickProtrusion =
      axis.spec.majorTickMark === "out" || axis.spec.majorTickMark === "cross"
        ? 4
        : 0;
    const labelGap = 4;

    if (labelSide !== "none") {
      if (axis.direction === "horizontal") {
        const zone = layout.labelZones.get(axis.name);
        const zoneOffset = zone?.offset ?? 0;
        const zoneH = Math.max(
          axisFontPx * 1.3,
          (zone?.depth ?? axisFontPx * 1.4) - labelGap - tickProtrusion,
        );
        const labelY =
          labelSide === "top"
            ? plotRect.y - labelGap - tickProtrusion - zoneOffset - zoneH
            : labelSide === "bottom"
              ? plotRect.y +
                plotRect.height +
                labelGap +
                tickProtrusion +
                zoneOffset
              : crossCoord + 4;

        if (scale.kind === "category") {
          const slotW = scale.slotSpan(plotRect);
          for (let i = 0; i < scale.categories.length; i += 1) {
            const cx = scale.indexToScreen(i, plotRect);
            const labelBox: Rect = {
              x: cx - slotW / 2,
              y: labelY,
              width: Math.max(4, slotW),
              height: zoneH,
            };
            parts.push(
              renderChartTextSvg({
                text: scale.categories[i]!,
                box: labelBox,
                style: axisTextStyle,
                theme,
                slideNumber,
                labelPath: `Chart.axes.${axis.name}.style`,
                align: "center",
                verticalAlign: labelSide === "top" ? "bottom" : "top",
                wrap: true,
                clipBox: labelBox,
                clipAllocator,
              }),
            );
          }
        } else {
          const tickSpacing = Math.max(
            24,
            plotRect.width / Math.max(1, scale.majorTicks.length - 1),
          );
          for (const val of scale.majorTicks) {
            const cx = scale.valueToScreen(val, plotRect);
            const text = formatChartNumber(val, scale.numberFormat);
            const labelBox: Rect = {
              x: cx - tickSpacing / 2,
              y: labelY,
              width: tickSpacing,
              height: zoneH,
            };
            parts.push(
              renderChartTextSvg({
                text,
                box: labelBox,
                style: axisTextStyle,
                theme,
                slideNumber,
                labelPath: `Chart.axes.${axis.name}.style`,
                align: "center",
                verticalAlign: labelSide === "top" ? "bottom" : "top",
                wrap: false,
                clipBox: chartBox,
                clipAllocator,
              }),
            );
          }
        }
      } else {
        // Vertical axis
        const titleBox = layout.axisTitleBoxes.get(axis.name);
        const leftBound =
          titleBox !== undefined && titleBox.x < plotRect.x
            ? titleBox.x + titleBox.width + 2
            : chartBox.x + 2;
        const rightBound =
          titleBox !== undefined && titleBox.x >= plotRect.x + plotRect.width
            ? titleBox.x - 2
            : chartBox.x + chartBox.width - 2;

        const zoneOffset = layout.labelZones.get(axis.name)?.offset ?? 0;
        let boxX: number;
        let boxW: number;
        let align: "left" | "right";
        if (labelSide === "right") {
          boxX =
            plotRect.x +
            plotRect.width +
            labelGap +
            tickProtrusion +
            zoneOffset;
          boxW = Math.max(6, rightBound - boxX);
          align = "left";
        } else if (labelSide === "left") {
          const rightEdge = plotRect.x - labelGap - tickProtrusion - zoneOffset;
          boxX = leftBound;
          boxW = Math.max(6, rightEdge - leftBound);
          align = "right";
        } else {
          const leftRoom = crossCoord - plotRect.x - 4;
          if (leftRoom >= 16) {
            boxW = Math.min(56, leftRoom);
            boxX = crossCoord - 4 - boxW;
            align = "right";
          } else {
            boxX = crossCoord + 4;
            boxW = Math.max(
              6,
              Math.min(56, plotRect.x + plotRect.width - boxX),
            );
            align = "left";
          }
        }

        if (scale.kind === "value") {
          const labelH = Math.max(axisFontPx * 1.4, 12);
          for (const val of scale.majorTicks) {
            const cy = scale.valueToScreen(val, plotRect);
            const text = formatChartNumber(val, scale.numberFormat);
            const labelBox: Rect = {
              x: boxX,
              y: cy - labelH / 2,
              width: boxW,
              height: labelH,
            };
            parts.push(
              renderChartTextSvg({
                text,
                box: labelBox,
                style: axisTextStyle,
                theme,
                slideNumber,
                labelPath: `Chart.axes.${axis.name}.style`,
                align,
                verticalAlign: "middle",
                wrap: false,
                clipBox: labelBox,
                clipAllocator,
              }),
            );
          }
        } else {
          const slotH = scale.slotSpan(plotRect);
          const labelH = Math.max(slotH, axisFontPx * 1.2);
          for (let i = 0; i < scale.categories.length; i += 1) {
            const cy = scale.indexToScreen(i, plotRect);
            const labelBox: Rect = {
              x: boxX,
              y: cy - labelH / 2,
              width: boxW,
              height: labelH,
            };
            parts.push(
              renderChartTextSvg({
                text: scale.categories[i]!,
                box: labelBox,
                style: axisTextStyle,
                theme,
                slideNumber,
                labelPath: `Chart.axes.${axis.name}.style`,
                align,
                verticalAlign: "middle",
                wrap: slotH >= axisFontPx * 2.2,
                clipBox: labelBox,
                clipAllocator,
              }),
            );
          }
        }
      }
    }

    const title = axisTitleParts(axis);
    const titleBox = layout.axisTitleBoxes.get(axis.name);
    if (title !== undefined && titleBox !== undefined) {
      const titleStyle = mergeChartTextStyle(axisTextStyle, title.style);
      const rotationDegrees =
        axis.direction === "vertical"
          ? titleBox.x < plotRect.x
            ? -90
            : 90
          : undefined;
      parts.push(
        renderChartTextSvg({
          text: title.text,
          box: titleBox,
          style: titleStyle,
          theme,
          slideNumber,
          labelPath: `Chart.axes.${axis.name}.title`,
          align: "center",
          verticalAlign: "middle",
          wrap: true,
          clipBox: titleBox,
          clipAllocator,
          ...(rotationDegrees === undefined ? {} : { rotationDegrees }),
        }),
      );
    }
  }

  return parts.join("");
}
