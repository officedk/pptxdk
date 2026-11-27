import type { Rect, PaintContext, PaintOutput } from "@pptxdk/core/capability";

import { parseTextContent } from "../../shared/text-service.ts";
import { textOnlyChartStyle, type ParsedChart } from "../chart-model.ts";
import { renderAxesSvg, renderGridlinesSvg } from "./axes-and-grid.ts";
import { renderColumnOrBarGroup } from "./bar-column.ts";
import { renderDataLabelsSvg } from "./data-labels.ts";
import { renderErrorBarsSvg } from "./error-bars.ts";
import { computeChartLayout } from "./layout.ts";
import { renderLegendSvg } from "./legend.ts";
import { renderLineOrAreaGroup } from "./line-area.ts";
import { renderPieOrDoughnutGroup } from "./pie-doughnut.ts";
import { buildAxisScales } from "./scales.ts";
import { renderScatterGroup } from "./scatter.ts";
import { resolveChartLineStroke, strokeAttrs } from "./style.ts";
import {
  createChartClipAllocator,
  mergeChartTextStyle,
  renderChartTextSvg,
  resolveChartFontSizePixels,
} from "./text-render.ts";
import { renderTrendlinesSvg } from "./trendlines.ts";
import { assertNever, type PlottedPointAnchor } from "./types.ts";

/**
 * Renders a static SVG preview for a resolved chart without mutating or leaking into native emit.
 */
export function drawChartPreview(
  parsed: ParsedChart,
  box: Rect,
  context: PaintContext,
): PaintOutput {
  const baseChartTextStyle = textOnlyChartStyle(parsed.style);
  const { style: resolvedChartStyle } = parseTextContent(
    [],
    baseChartTextStyle,
    context.theme,
    { slideNumber: context.slideNumber },
    "Chart.style",
  );
  const baseFontSizePx = resolveChartFontSizePixels(
    resolvedChartStyle.fontSize,
  );

  const scales = buildAxisScales(parsed);
  const layout = computeChartLayout({
    parsed,
    box,
    scales,
    baseChartTextStyle,
    baseFontSizePx,
    theme: context.theme,
    slideNumber: context.slideNumber,
  });

  const clipAllocator = createChartClipAllocator(
    context.slideNumber,
    context.objectId,
  );
  const outerClipId = clipAllocator.nextClipId();
  clipAllocator.defs.push(
    `<clipPath id="${outerClipId}" clipPathUnits="userSpaceOnUse"><rect x="${box.x}" y="${box.y}" width="${Math.max(0, box.width)}" height="${Math.max(0, box.height)}"/></clipPath>`,
  );
  const plotClipId = clipAllocator.nextClipId();
  clipAllocator.defs.push(
    `<clipPath id="${plotClipId}" clipPathUnits="userSpaceOnUse"><rect x="${layout.plotRect.x}" y="${layout.plotRect.y}" width="${Math.max(0, layout.plotRect.width)}" height="${Math.max(0, layout.plotRect.height)}"/></clipPath>`,
  );

  const bodyParts: string[] = [];

  // 1. Background gridlines
  bodyParts.push(
    renderGridlinesSvg({
      parsed,
      scales,
      plotRect: layout.plotRect,
      theme: context.theme,
    }),
  );

  // 2. Plotted series clipped to plotRect
  const anchors: PlottedPointAnchor[] = [];
  const plotParts: string[] = [];
  for (let groupIndex = 0; groupIndex < parsed.groups.length; groupIndex += 1) {
    const group = parsed.groups[groupIndex]!;
    switch (group.kind) {
      case "column":
      case "bar": {
        const rendered = renderColumnOrBarGroup({
          group,
          groupIndex,
          scales,
          plotRect: layout.plotRect,
          theme: context.theme,
        });
        if (rendered.svg !== "") plotParts.push(rendered.svg);
        anchors.push(...rendered.anchors);
        break;
      }
      case "line":
      case "area": {
        const rendered = renderLineOrAreaGroup({
          group,
          groupIndex,
          scales,
          plotRect: layout.plotRect,
          displayBlanksAs: parsed.displayBlanksAs,
          theme: context.theme,
        });
        if (rendered.svg !== "") plotParts.push(rendered.svg);
        anchors.push(...rendered.anchors);
        break;
      }
      case "scatter": {
        const rendered = renderScatterGroup({
          group,
          groupIndex,
          scales,
          plotRect: layout.plotRect,
          displayBlanksAs: parsed.displayBlanksAs,
          theme: context.theme,
        });
        if (rendered.svg !== "") plotParts.push(rendered.svg);
        anchors.push(...rendered.anchors);
        break;
      }
      case "pie":
      case "doughnut": {
        const rendered = renderPieOrDoughnutGroup({
          group,
          groupIndex,
          plotRect: layout.plotRect,
          theme: context.theme,
        });
        if (rendered.svg !== "") plotParts.push(rendered.svg);
        anchors.push(...rendered.anchors);
        break;
      }
      default:
        assertNever(group);
    }
  }
  const errorBarsSvg = renderErrorBarsSvg({
    parsed,
    anchors,
    scales,
    plotRect: layout.plotRect,
    theme: context.theme,
  });
  if (errorBarsSvg !== "") {
    plotParts.push(errorBarsSvg);
  }
  const trendlines = renderTrendlinesSvg({
    parsed,
    scales,
    plotRect: layout.plotRect,
    chartBox: box,
    baseChartTextStyle,
    theme: context.theme,
    slideNumber: context.slideNumber,
    clipAllocator,
  });
  if (trendlines.linesSvg !== "") {
    plotParts.push(trendlines.linesSvg);
  }
  if (plotParts.length > 0) {
    bodyParts.push(
      `<g clip-path="url(#${plotClipId})">${plotParts.join("")}</g>`,
    );
  }

  // Subtle zero baseline when a value axis spans negative and positive values without an axis line at zero.
  for (const axis of parsed.axes) {
    if (axis.kind !== "value") continue;
    const valScale = scales.get(axis.name);
    if (valScale?.kind !== "value") continue;
    if (valScale.logBase !== undefined) continue;
    if (valScale.min < 0 && valScale.max > 0) {
      const zeroCoord = valScale.zeroScreen(layout.plotRect);
      const zeroStroke = resolveChartLineStroke(true, context.theme, {
        color: "dk1",
        width: 0.75,
        alpha: 0.35,
      })!;
      const path =
        axis.direction === "vertical"
          ? `M${layout.plotRect.x} ${zeroCoord}H${layout.plotRect.x + layout.plotRect.width}`
          : `M${zeroCoord} ${layout.plotRect.y}V${layout.plotRect.y + layout.plotRect.height}`;
      bodyParts.push(
        `<path d="${path}" ${strokeAttrs(zeroStroke)} fill="none"/>`,
      );
    }
  }

  // 3. Axes, tick marks, tick labels, and axis titles
  bodyParts.push(
    renderAxesSvg({
      parsed,
      chartBox: box,
      layout,
      scales,
      baseChartTextStyle,
      theme: context.theme,
      slideNumber: context.slideNumber,
      clipAllocator,
    }),
  );

  // 4. Data labels & trendline annotations
  bodyParts.push(
    renderDataLabelsSvg({
      parsed,
      anchors,
      chartBox: box,
      baseChartTextStyle,
      theme: context.theme,
      slideNumber: context.slideNumber,
      clipAllocator,
    }),
  );
  if (trendlines.labelsSvg !== "") {
    bodyParts.push(trendlines.labelsSvg);
  }

  // 5. Chart title
  if (parsed.title !== undefined && layout.titleBox !== undefined) {
    const titleObj =
      typeof parsed.title === "string" ? { text: parsed.title } : parsed.title;
    bodyParts.push(
      renderChartTextSvg({
        text: titleObj.text,
        box: layout.titleBox,
        style: mergeChartTextStyle(baseChartTextStyle, titleObj.style),
        theme: context.theme,
        slideNumber: context.slideNumber,
        labelPath: "Chart.title",
        align: "center",
        verticalAlign: "middle",
        wrap: true,
        clipBox: layout.titleBox,
        clipAllocator,
      }),
    );
  }

  // 6. Legend
  if (
    parsed.legend !== undefined &&
    parsed.legend !== false &&
    layout.legendBox !== undefined &&
    layout.measuredLegend !== undefined
  ) {
    bodyParts.push(
      renderLegendSvg({
        measured: layout.measuredLegend,
        position: parsed.legend,
        legendBox: layout.legendBox,
        style: baseChartTextStyle,
        theme: context.theme,
        slideNumber: context.slideNumber,
        clipAllocator,
      }),
    );
  }

  const defsXml =
    clipAllocator.defs.length > 0
      ? `<defs>${clipAllocator.defs.join("")}</defs>`
      : "";
  return {
    svg: `${defsXml}<g clip-path="url(#${outerClipId})">${bodyParts.join("")}</g>`,
  };
}
