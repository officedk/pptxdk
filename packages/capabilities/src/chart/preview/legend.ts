import type {
  GeometryBox,
  ThemeContext,
} from "@react-pptx/core/capability";

import type { ParsedChart } from "../chart-model.ts";
import type { TextStyle } from "../../shared/text-service.ts";
import { resolveSeriesColor } from "./style.ts";
import {
  measureChartText,
  renderChartTextSvg,
  type ChartTextClipAllocator,
} from "./text-render.ts";
import type { ChartLegendEntry } from "./types.ts";

export type LegendPosition =
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "topRight";

export function collectLegendEntries(
  parsed: ParsedChart,
  theme: ThemeContext,
): ChartLegendEntry[] {
  const entries: ChartLegendEntry[] = [];
  for (const group of parsed.groups) {
    for (const series of group.series) {
      const color = resolveSeriesColor(
        series.color,
        series.seriesIndex,
        theme,
      );
      entries.push({
        label: series.name,
        colorHex: color.hex,
        colorAlpha: color.alpha ?? 1,
        kind: "swatch",
      });
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
      itemWidth: swatchSize + swatchGap + measured.width,
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
          (sum, item, idx) =>
            sum + item.itemWidth + (idx > 0 ? itemGapX : 0),
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
    swatchGap,
    itemGapX,
    rowGapY,
    rowHeight,
    rows,
  };
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
      const swatchY = rowY + (measured.rowHeight - measured.swatchSize) / 2;
      parts.push(
        `<rect x="${cursorX}" y="${swatchY}" width="${measured.swatchSize}" height="${measured.swatchSize}" fill="#${item.entry.colorHex}" fill-opacity="${item.entry.colorAlpha}"/>`,
      );
      const textX = cursorX + measured.swatchSize + measured.swatchGap;
      const availableTextWidth = Math.max(
        1,
        Math.min(
          item.itemWidth - measured.swatchSize - measured.swatchGap,
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
