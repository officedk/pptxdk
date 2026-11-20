import type { GeometryBox, ThemeContext } from "@react-pptx/core/capability";

import type { TextStyle } from "../../shared/text-service.ts";
import { mergeDataLabels } from "../chart-labels.ts";
import type { ParsedChart } from "../chart-model.ts";
import { formatChartNumber } from "./number-format.ts";
import {
  measureChartText,
  mergeChartTextStyle,
  renderChartTextSvg,
  resolveChartFontSizePixels,
  type ChartTextClipAllocator,
} from "./text-render.ts";
import type {
  PlottedBarAnchor,
  PlottedCartesianPointAnchor,
  PlottedPointAnchor,
} from "./types.ts";

type LabelPlacement = {
  readonly box: GeometryBox;
  readonly align: "left" | "center" | "right";
  readonly verticalAlign: "top" | "middle" | "bottom";
};

function barLabelBox(
  anchor: PlottedBarAnchor,
  position: string,
  textWidth: number,
  textHeight: number,
): LabelPlacement {
  const pad = 3;
  const labelH = Math.max(textHeight, 10);

  if (!anchor.horizontal) {
    const colW = Math.max(anchor.width, textWidth + 2);
    const colX = anchor.x + (anchor.width - colW) / 2;
    if (position === "center") {
      return {
        box: {
          x: colX,
          y: anchor.y,
          width: colW,
          height: Math.max(anchor.height, labelH),
        },
        align: "center",
        verticalAlign: "middle",
      };
    }
    const endIsTop = anchor.valueEndSide === "top";
    if (position === "insideEnd") {
      const y = endIsTop
        ? anchor.y + pad
        : anchor.y + anchor.height - labelH - pad;
      return {
        box: { x: colX, y, width: colW, height: labelH },
        align: "center",
        verticalAlign: endIsTop ? "top" : "bottom",
      };
    }
    if (position === "insideBase") {
      const y = endIsTop
        ? anchor.y + anchor.height - labelH - pad
        : anchor.y + pad;
      return {
        box: { x: colX, y, width: colW, height: labelH },
        align: "center",
        verticalAlign: endIsTop ? "bottom" : "top",
      };
    }
    // outsideEnd
    const y = endIsTop
      ? anchor.y - labelH - pad
      : anchor.y + anchor.height + pad;
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
      box: {
        x: anchor.x,
        y: barY,
        width: Math.max(anchor.width, barW),
        height: labelH,
      },
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

function pointLabelBox(
  anchor: PlottedCartesianPointAnchor,
  position: string,
  textWidth: number,
  textHeight: number,
): LabelPlacement {
  const offset = (anchor.markerRadius ?? 1.5) + 4;
  const w = Math.max(textWidth + 4, 12);
  const h = Math.max(textHeight, 10);

  switch (position) {
    case "center":
      return {
        box: { x: anchor.x - w / 2, y: anchor.y - h / 2, width: w, height: h },
        align: "center",
        verticalAlign: "middle",
      };
    case "bottom":
      return {
        box: { x: anchor.x - w / 2, y: anchor.y + offset, width: w, height: h },
        align: "center",
        verticalAlign: "top",
      };
    case "left":
      return {
        box: {
          x: anchor.x - offset - w,
          y: anchor.y - h / 2,
          width: w,
          height: h,
        },
        align: "right",
        verticalAlign: "middle",
      };
    case "right":
      return {
        box: {
          x: anchor.x + offset,
          y: anchor.y - h / 2,
          width: w,
          height: h,
        },
        align: "left",
        verticalAlign: "middle",
      };
    case "top":
    default:
      return {
        box: {
          x: anchor.x - w / 2,
          y: anchor.y - offset - h,
          width: w,
          height: h,
        },
        align: "center",
        verticalAlign: "bottom",
      };
  }
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
    if (anchor.kind === "slice") continue;
    const group = parsed.groups[anchor.groupIndex];
    if (group === undefined) continue;

    let series;
    if (anchor.kind === "bar") {
      if (group.kind !== "column" && group.kind !== "bar") continue;
      series = group.series[anchor.groupSeriesIndex];
    } else {
      if (
        group.kind !== "line" &&
        group.kind !== "area" &&
        group.kind !== "scatter"
      ) {
        continue;
      }
      series = group.series[anchor.groupSeriesIndex];
    }
    if (series === undefined) continue;

    const seriesMerged = mergeDataLabels(group.dataLabels, series.dataLabels);
    const pointMerged = mergeDataLabels(
      seriesMerged,
      series.points?.[anchor.pointIndex]?.dataLabels,
    );
    if (pointMerged === undefined || pointMerged === false) continue;

    const labelSegments: string[] = [];
    if (group.kind === "scatter") {
      if (
        pointMerged.xValue === true &&
        anchor.kind === "point" &&
        anchor.xValue !== undefined
      ) {
        labelSegments.push(
          formatChartNumber(anchor.xValue, pointMerged.numberFormat),
        );
      }
    } else if (pointMerged.category === true && anchor.category !== undefined) {
      labelSegments.push(anchor.category);
    }
    if (pointMerged.value === true) {
      labelSegments.push(
        formatChartNumber(anchor.value, pointMerged.numberFormat),
      );
    }
    if (pointMerged.percent === true && anchor.percent !== undefined) {
      labelSegments.push(
        formatChartNumber(anchor.percent, pointMerged.numberFormat ?? "0%"),
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
    const textH = Math.max(measured.height, fontPx * 1.15);

    const placement: LabelPlacement =
      anchor.kind === "bar"
        ? barLabelBox(
            anchor,
            pointMerged.position ??
              (group.kind === "column" || group.kind === "bar"
                ? group.stacking === "clustered"
                  ? "outsideEnd"
                  : "center"
                : "center"),
            measured.width,
            textH,
          )
        : pointLabelBox(
            anchor,
            pointMerged.position ??
              (group.kind === "scatter"
                ? "right"
                : group.kind === "area"
                  ? "center"
                  : "top"),
            measured.width,
            textH,
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
