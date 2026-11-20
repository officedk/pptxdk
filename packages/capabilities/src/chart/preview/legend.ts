import type { GeometryBox, ThemeContext } from "@react-pptx/core/capability";

import type { TextStyle } from "../../shared/text-service.ts";
import type { ParsedChart } from "../chart-model.ts";
import { renderMarkerSvg, resolveEffectiveMarker } from "./markers.ts";
import { resolveChartLineStroke, resolveSeriesColor } from "./style.ts";
import {
  measureChartText,
  renderChartTextSvg,
  type ChartTextClipAllocator,
} from "./text-render.ts";
import {
  assertNever,
  type ChartLegendEntry,
  type ChartLegendMarker,
} from "./types.ts";

export type LegendPosition = "top" | "bottom" | "left" | "right" | "topRight";

function toLegendMarker(
  resolved: ReturnType<typeof resolveEffectiveMarker>,
): ChartLegendMarker | undefined {
  if (resolved === undefined) return undefined;
  return {
    symbol: resolved.symbol,
    fillHex: resolved.fillHex,
    fillAlpha: resolved.fillAlpha,
    ...(resolved.stroke === undefined
      ? {}
      : {
          strokeHex: resolved.stroke.hex,
          strokeAlpha: resolved.stroke.alpha,
          strokeWidth: resolved.stroke.width,
        }),
  };
}

export function collectLegendEntries(
  parsed: ParsedChart,
  theme: ThemeContext,
): ChartLegendEntry[] {
  const entries: ChartLegendEntry[] = [];
  for (const group of parsed.groups) {
    switch (group.kind) {
      case "column":
      case "bar":
      case "area": {
        for (const series of group.series) {
          const color = resolveSeriesColor(
            series.color,
            series.seriesIndex,
            theme,
          );
          const stroke =
            series.line === undefined
              ? undefined
              : resolveChartLineStroke(series.line, theme, {
                  color: "dk1",
                  width: 1,
                });
          entries.push({
            label: series.name,
            colorHex: color.hex,
            colorAlpha: color.alpha ?? 1,
            kind: "swatch",
            ...(stroke === undefined
              ? {}
              : {
                  strokeHex: stroke.hex,
                  strokeAlpha: stroke.alpha,
                  strokeWidth: stroke.width,
                  ...(stroke.dashArray === undefined
                    ? {}
                    : { dashArray: stroke.dashArray }),
                }),
          });
        }
        break;
      }
      case "pie":
      case "doughnut": {
        // Pie and doughnut legends explain category slices; independent-category doughnuts
        // strictly use the first declared ring's categories.
        const firstSeries = group.series[0];
        if (firstSeries === undefined) break;
        for (let j = 0; j < firstSeries.categories.length; j += 1) {
          const point = firstSeries.points?.[j];
          const color = resolveSeriesColor(
            point?.color ?? firstSeries.color,
            j,
            theme,
          );
          const lineSpec = point?.line ?? firstSeries.line;
          const stroke =
            lineSpec === undefined
              ? undefined
              : resolveChartLineStroke(lineSpec, theme, {
                  color: "dk1",
                  width: 1,
                });
          entries.push({
            label: firstSeries.categories[j] ?? "",
            colorHex: color.hex,
            colorAlpha: color.alpha ?? 1,
            kind: "swatch",
            ...(stroke === undefined
              ? {}
              : {
                  strokeHex: stroke.hex,
                  strokeAlpha: stroke.alpha,
                  strokeWidth: stroke.width,
                  ...(stroke.dashArray === undefined
                    ? {}
                    : { dashArray: stroke.dashArray }),
                }),
          });
        }
        break;
      }
      case "line": {
        for (const series of group.series) {
          const fallbackColor =
            series.line?.color ??
            series.color ??
            `accent${(series.seriesIndex % 6) + 1}`;
          const stroke = resolveChartLineStroke(series.line, theme, {
            color: fallbackColor,
            width: 2,
          })!;
          const marker = toLegendMarker(
            resolveEffectiveMarker({
              groupDefaultEnabled: false,
              seriesMarker: series.marker,
              pointMarker: undefined,
              seriesColor: series.color,
              pointColor: undefined,
              seriesIndex: series.seriesIndex,
              theme,
            }),
          );
          entries.push({
            label: series.name,
            colorHex: stroke.hex,
            colorAlpha: stroke.alpha,
            kind: "line",
            strokeWidth: stroke.width,
            ...(stroke.dashArray === undefined
              ? {}
              : { dashArray: stroke.dashArray }),
            ...(marker === undefined ? {} : { marker }),
          });
        }
        break;
      }
      case "scatter": {
        const hasLine = group.line !== false;
        for (const series of group.series) {
          const fallbackColor =
            series.line?.color ??
            series.color ??
            `accent${(series.seriesIndex % 6) + 1}`;
          const stroke = hasLine
            ? resolveChartLineStroke(series.line, theme, {
                color: fallbackColor,
                width: 2,
              })!
            : undefined;
          const color = resolveSeriesColor(
            series.color,
            series.seriesIndex,
            theme,
          );
          const marker = toLegendMarker(
            group.markers
              ? resolveEffectiveMarker({
                  groupDefaultEnabled: true,
                  seriesMarker: series.marker,
                  pointMarker: undefined,
                  seriesColor: series.color,
                  pointColor: undefined,
                  seriesIndex: series.seriesIndex,
                  theme,
                })
              : undefined,
          );
          entries.push({
            label: series.name,
            colorHex: stroke?.hex ?? marker?.fillHex ?? color.hex,
            colorAlpha: stroke?.alpha ?? marker?.fillAlpha ?? color.alpha ?? 1,
            kind: hasLine ? "line" : "marker",
            ...(stroke === undefined
              ? {}
              : {
                  strokeWidth: stroke.width,
                  ...(stroke.dashArray === undefined
                    ? {}
                    : { dashArray: stroke.dashArray }),
                }),
            ...(marker === undefined ? {} : { marker }),
          });
        }
        break;
      }
      default:
        assertNever(group);
    }
  }
  return entries;
}

type MeasuredLegendItem = {
  readonly entry: ChartLegendEntry;
  readonly itemWidth: number;
  readonly textHeight: number;
};

export type MeasuredLegend = {
  readonly width: number;
  readonly height: number;
  readonly swatchSize: number;
  readonly iconWidth: number;
  readonly swatchGap: number;
  readonly itemGapX: number;
  readonly rowGapY: number;
  readonly rowHeight: number;
  readonly rows: readonly (readonly MeasuredLegendItem[])[];
};

export function measureLegend(options: {
  readonly entries: readonly ChartLegendEntry[];
  readonly position: LegendPosition;
  readonly maxWidth: number;
  readonly fontSizePx: number;
  readonly style: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
}): MeasuredLegend {
  const { entries, position, maxWidth, fontSizePx, style, theme, slideNumber } =
    options;
  const swatchSize = Math.max(5, Math.min(fontSizePx * 0.72, 12));
  const hasLineEntry = entries.some((entry) => entry.kind === "line");
  const iconWidth = hasLineEntry
    ? Math.max(swatchSize, Math.min(fontSizePx * 1.3, 18))
    : swatchSize;
  const swatchGap = Math.max(3, swatchSize * 0.45);
  const itemGapX = Math.max(8, fontSizePx * 0.75);
  const rowGapY = Math.max(2, fontSizePx * 0.25);

  const items: MeasuredLegendItem[] = entries.map((entry) => {
    const measured = measureChartText({
      text: entry.label,
      style,
      theme,
      slideNumber,
      labelPath: "Chart.legend",
    });
    return {
      entry,
      itemWidth: iconWidth + swatchGap + measured.width,
      textHeight: Math.max(swatchSize, measured.height),
    };
  });

  const rowHeight = Math.max(
    swatchSize,
    ...items.map((item) => item.textHeight),
  );

  if (position === "top" || position === "bottom") {
    const rows: MeasuredLegendItem[][] = [];
    let currentRow: MeasuredLegendItem[] = [];
    let currentWidth = 0;
    const limit = Math.max(20, maxWidth);
    for (const item of items) {
      const nextWidth =
        currentRow.length === 0
          ? item.itemWidth
          : currentWidth + itemGapX + item.itemWidth;
      if (currentRow.length > 0 && nextWidth > limit) {
        rows.push(currentRow);
        currentRow = [item];
        currentWidth = item.itemWidth;
      } else {
        currentRow.push(item);
        currentWidth = nextWidth;
      }
    }
    if (currentRow.length > 0) rows.push(currentRow);
    const maxRowWidth = Math.max(
      0,
      ...rows.map((row) =>
        row.reduce(
          (sum, item, idx) => sum + item.itemWidth + (idx > 0 ? itemGapX : 0),
          0,
        ),
      ),
    );
    const totalHeight =
      rows.length === 0
        ? 0
        : rows.length * rowHeight + (rows.length - 1) * rowGapY;
    return {
      width: Math.min(maxWidth, maxRowWidth),
      height: totalHeight,
      swatchSize,
      iconWidth,
      swatchGap,
      itemGapX,
      rowGapY,
      rowHeight,
      rows,
    };
  }

  const rows = items.map((item) => [item]);
  const widest = Math.max(0, ...items.map((item) => item.itemWidth));
  const totalHeight =
    rows.length === 0
      ? 0
      : rows.length * rowHeight + (rows.length - 1) * rowGapY;
  return {
    width: widest,
    height: totalHeight,
    swatchSize,
    iconWidth,
    swatchGap,
    itemGapX,
    rowGapY,
    rowHeight,
    rows,
  };
}

function renderLegendIconSvg(
  entry: ChartLegendEntry,
  x: number,
  rowY: number,
  measured: MeasuredLegend,
): string {
  const cy = rowY + measured.rowHeight / 2;
  const cx = x + measured.iconWidth / 2;

  if (entry.kind === "swatch") {
    const size = measured.swatchSize;
    const swatchX = cx - size / 2;
    const swatchY = cy - size / 2;
    const dashAttr =
      entry.dashArray === undefined
        ? ""
        : ` stroke-dasharray="${entry.dashArray}"`;
    const strokeAttr =
      entry.strokeWidth === undefined
        ? ""
        : ` stroke="#${entry.strokeHex ?? entry.colorHex}" stroke-opacity="${entry.strokeAlpha ?? 1}" stroke-width="${Math.min(1.5, entry.strokeWidth)}"${dashAttr}`;
    return `<rect x="${swatchX}" y="${swatchY}" width="${size}" height="${size}" fill="#${entry.colorHex}" fill-opacity="${entry.colorAlpha}"${strokeAttr}/>`;
  }

  const parts: string[] = [];
  if (entry.kind === "line") {
    const strokeW = Math.min(2.5, Math.max(1.25, entry.strokeWidth ?? 2));
    const dashAttr =
      entry.dashArray === undefined
        ? ""
        : ` stroke-dasharray="${entry.dashArray}"`;
    parts.push(
      `<line x1="${x}" y1="${cy}" x2="${x + measured.iconWidth}" y2="${cy}" stroke="#${entry.colorHex}" stroke-opacity="${entry.colorAlpha}" stroke-width="${strokeW}"${dashAttr} stroke-linecap="round"/>`,
    );
  }

  if (entry.marker !== undefined) {
    const r = Math.min(measured.swatchSize * 0.42, 4);
    parts.push(
      renderMarkerSvg(cx, cy, {
        symbol: entry.marker.symbol,
        radius: r,
        fillHex: entry.marker.fillHex,
        fillAlpha: entry.marker.fillAlpha,
        ...(entry.marker.strokeHex === undefined
          ? {}
          : {
              stroke: {
                hex: entry.marker.strokeHex,
                alpha: entry.marker.strokeAlpha ?? 1,
                width: Math.min(1.5, entry.marker.strokeWidth ?? 1),
              },
            }),
      }),
    );
  }

  return parts.join("");
}

export function renderLegendSvg(options: {
  readonly measured: MeasuredLegend;
  readonly position: LegendPosition;
  readonly legendBox: GeometryBox;
  readonly style: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly clipAllocator: ChartTextClipAllocator;
}): string {
  const {
    measured,
    position,
    legendBox,
    style,
    theme,
    slideNumber,
    clipAllocator,
  } = options;
  if (
    measured.rows.length === 0 ||
    legendBox.width <= 0 ||
    legendBox.height <= 0
  ) {
    return "";
  }

  const parts: string[] = [];
  const totalContentHeight =
    measured.rows.length * measured.rowHeight +
    Math.max(0, measured.rows.length - 1) * measured.rowGapY;
  const startY =
    position === "topRight"
      ? legendBox.y
      : legendBox.y + Math.max(0, (legendBox.height - totalContentHeight) / 2);

  for (let r = 0; r < measured.rows.length; r += 1) {
    const row = measured.rows[r]!;
    const rowY = startY + r * (measured.rowHeight + measured.rowGapY);
    if (rowY >= legendBox.y + legendBox.height) break;

    const rowWidth = row.reduce(
      (sum, item, idx) =>
        sum + item.itemWidth + (idx > 0 ? measured.itemGapX : 0),
      0,
    );
    let cursorX =
      position === "top" || position === "bottom"
        ? legendBox.x + Math.max(0, (legendBox.width - rowWidth) / 2)
        : legendBox.x;

    for (const item of row) {
      parts.push(renderLegendIconSvg(item.entry, cursorX, rowY, measured));
      const textX = cursorX + measured.iconWidth + measured.swatchGap;
      const availableTextWidth = Math.max(
        1,
        Math.min(
          item.itemWidth - measured.iconWidth - measured.swatchGap,
          legendBox.x + legendBox.width - textX,
        ),
      );
      parts.push(
        renderChartTextSvg({
          text: item.entry.label,
          box: {
            x: textX,
            y: rowY,
            width: availableTextWidth,
            height: measured.rowHeight,
          },
          style,
          theme,
          slideNumber,
          labelPath: "Chart.legend",
          align: "left",
          verticalAlign: "middle",
          wrap: false,
          clipBox: legendBox,
          clipAllocator,
        }),
      );
      cursorX += item.itemWidth + measured.itemGapX;
    }
  }

  return parts.join("");
}
