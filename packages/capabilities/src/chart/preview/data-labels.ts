import type { GeometryBox, ThemeContext } from "@react-pptx/core/capability";

import { mergeDataLabels } from "../chart-labels.ts";
import type { ParsedChart } from "../chart-model.ts";
import type { TextStyle } from "../../shared/text-service.ts";
import { formatChartNumber } from "./number-format.ts";
import {
  measureChartText,
  mergeChartTextStyle,
  renderChartTextSvg,
  resolveChartFontSizePixels,
  type ChartTextClipAllocator,
} from "./text-render.ts";
import type { PlottedBarAnchor, PlottedPointAnchor } from "./types.ts";

function barLabelBox(
  anchor: PlottedBarAnchor,
  position: string,
  textWidth: number,
  textHeight: number,
): {
  readonly box: GeometryBox;
  readonly align: "left" | "center" | "right";
  readonly verticalAlign: "top" | "middle" | "bottom";
} {
  const pad = 3;
  const labelH = Math.max(textHeight, 10);

  if (!anchor.horizontal) {
    const colW = Math.max(anchor.width, textWidth + 2);
    const colX = anchor.x + (anchor.width - colW) / 2;
    if (position === "center") {
      return {
        box: { x: colX, y: anchor.y, width: colW, height: Math.max(anchor.height, labelH) },
        align: "center",
        verticalAlign: "middle",
      };
    }
    const endIsTop = anchor.valueEndSide === "top";
    if (position === "insideEnd") {
      const y = endIsTop ? anchor.y + pad : anchor.y + anchor.height - labelH - pad;
      return {
        box: { x: colX, y, width: colW, height: labelH },
        align: "center",
        verticalAlign: endIsTop ? "top" : "bottom",
      };
    }
    if (position === "insideBase") {
      const y = endIsTop ? anchor.y + anchor.height - labelH - pad : anchor.y + pad;
      return {
        box: { x: colX, y, width: colW, height: labelH },
        align: "center",
        verticalAlign: endIsTop ? "bottom" : "top",
      };
    }
    // outsideEnd
    const y = endIsTop ? anchor.y - labelH - pad : anchor.y + anchor.height + pad;
    return {
      box: { x: colX, y, width: colW, height: labelH },
      align: "center",
      verticalAlign: endIsTop ? "bottom" : "top",
    };
  }

  // Horizontal bar
  const barW = Math.max(textWidth + pad * 2, 12);
  const barY = anchor.y + (anchor.height - labelH) / 2;
  if (position === "center") {
    return {
      box: { x: anchor.x, y: barY, width: Math.max(anchor.width, barW), height: labelH },
      align: "center",
      verticalAlign: "middle",
    };
  }
  const endIsRight = anchor.valueEndSide === "right";
  if (position === "insideEnd") {
    const x = endIsRight
      ? anchor.x + anchor.width - barW - pad
      : anchor.x + pad;
    return {
      box: { x, y: barY, width: barW, height: labelH },
      align: endIsRight ? "right" : "left",
      verticalAlign: "middle",
    };
  }
  if (position === "insideBase") {
    const x = endIsRight
      ? anchor.x + pad
      : anchor.x + anchor.width - barW - pad;
    return {
      box: { x, y: barY, width: barW, height: labelH },
      align: endIsRight ? "left" : "right",
      verticalAlign: "middle",
    };
  }
  // outsideEnd
  const x = endIsRight ? anchor.x + anchor.width + pad : anchor.x - barW - pad;
  return {
    box: { x, y: barY, width: barW, height: labelH },
    align: endIsRight ? "left" : "right",
    verticalAlign: "middle",
  };
}

export function renderDataLabelsSvg(options: {
  readonly parsed: ParsedChart;
  readonly anchors: readonly PlottedPointAnchor[];
  readonly chartBox: GeometryBox;
  readonly baseChartTextStyle: TextStyle | undefined;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
  readonly clipAllocator: ChartTextClipAllocator;
}): string {
  const {
    parsed,
    anchors,
    chartBox,
    baseChartTextStyle,
    theme,
    slideNumber,
    clipAllocator,
  } = options;
  const parts: string[] = [];

  for (const anchor of anchors) {
    if (anchor.kind !== "bar") continue;
    const group = parsed.groups[anchor.groupIndex];
    if (group === undefined || (group.kind !== "column" && group.kind !== "bar")) {
      continue;
    }
    const series = group.series[anchor.groupSeriesIndex];
    if (series === undefined) continue;

    const seriesMerged = mergeDataLabels(group.dataLabels, series.dataLabels);
    const pointMerged = mergeDataLabels(
      seriesMerged,
      series.points?.[anchor.pointIndex]?.dataLabels,
    );
    if (pointMerged === undefined || pointMerged === false) continue;

    const labelSegments: string[] = [];
    if (pointMerged.category === true) {
      labelSegments.push(anchor.category);
    }
    if (pointMerged.value === true) {
      labelSegments.push(
        formatChartNumber(anchor.value, pointMerged.numberFormat),
      );
    }
    if (pointMerged.percent === true && anchor.percent !== undefined) {
      labelSegments.push(
        formatChartNumber(
          anchor.percent,
          pointMerged.numberFormat ?? "0%",
        ),
      );
    }
    if (labelSegments.length === 0) continue;

    const text = labelSegments.join(", ");
    const groupStyle =
      group.dataLabels !== undefined && group.dataLabels !== false
        ? group.dataLabels.style
        : undefined;
    const pointColor = series.points?.[anchor.pointIndex]?.labelColor;
    const style = mergeChartTextStyle(
      mergeChartTextStyle(baseChartTextStyle, groupStyle),
      pointColor === undefined ? undefined : { color: pointColor },
    );

    const measured = measureChartText({
      text,
      style,
      theme,
      slideNumber,
      labelPath: `Chart.groups[${anchor.groupIndex}].dataLabels`,
    });
    const fontPx = resolveChartFontSizePixels(style?.fontSize);
    const position =
      pointMerged.position ??
      (group.stacking === "clustered" ? "outsideEnd" : "center");
    const placement = barLabelBox(
      anchor,
      position,
      measured.width,
      Math.max(measured.height, fontPx * 1.15),
    );

    parts.push(
      renderChartTextSvg({
        text,
        box: placement.box,
        style,
        theme,
        slideNumber,
        labelPath: `Chart.groups[${anchor.groupIndex}].dataLabels`,
        align: placement.align,
        verticalAlign: placement.verticalAlign,
        wrap: false,
        clipBox: chartBox,
        clipAllocator,
      }),
    );
  }

  return parts.join("");
}
