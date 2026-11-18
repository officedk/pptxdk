import {
  CapabilityError,
  type GeometryBox,
  type ObjectPreviewContext,
  type ObjectPreviewOutput,
} from "@react-pptx/core/capability";

import { textOnlyChartStyle, type ParsedChart } from "./chart-model.ts";
import {
  measureTextLayoutFromParagraphs,
  parseTextContent,
  previewTextBoxOutlines,
  renderGlyphPathSvg,
  type FontSize,
} from "./text-service.ts";

const allowedPreviewSeriesKeys = new Set([
  "name",
  "values",
  "categories",
  "groupIndex",
  "groupSeriesIndex",
  "seriesIndex",
]);

function resolveChartFontSizePixels(fontSize: FontSize | undefined): number {
  if (fontSize === undefined) {
    throw new CapabilityError(
      "preview-missing-font-size",
      "Preview requires a resolved chart font size",
    );
  }
  if (typeof fontSize === "number") return fontSize;
  const parsed = Number.parseFloat(fontSize);
  return fontSize.endsWith("pt") ? (parsed * 4) / 3 : parsed;
}

/**
 * Renders a static SVG preview for single-group finite non-negative clustered column charts.
 * Uses text-service directly so trimmed chart builds do not pull in textCapability.
 */
export function drawChartPreview(
  parsed: ParsedChart,
  box: GeometryBox,
  context: ObjectPreviewContext,
): ObjectPreviewOutput {
  const group = parsed.groups[0];
  if (
    parsed.groups.length !== 1 ||
    group?.kind !== "column" ||
    group.stacking !== undefined ||
    group.gapWidth !== undefined ||
    group.overlap !== undefined ||
    parsed.plotArea !== undefined ||
    parsed.hasExplicitAxes ||
    parsed.displayBlanksAs !== undefined ||
    (parsed.legend !== undefined &&
      parsed.legend !== false &&
      parsed.legend !== "bottom")
  ) {
    throw new CapabilityError(
      "preview-unsupported-chart",
      "Chart preview supports only default non-negative clustered columns with an optional bottom legend",
    );
  }
  if (
    group.series.length === 0 ||
    parsed.categories.length === 0 ||
    group.series.some(
      (series) =>
        Object.keys(series).some((key) => !allowedPreviewSeriesKeys.has(key)) ||
        series.values.some(
          (value) => value === null || !Number.isFinite(value) || value < 0,
        ),
    )
  ) {
    throw new CapabilityError(
      "preview-unsupported-chart",
      "Chart preview does not support empty, negative, missing, or styled series",
    );
  }
  if (
    group.dataLabels !== undefined &&
    group.dataLabels !== false &&
    Object.keys(group.dataLabels).some((key) => key !== "value")
  ) {
    throw new CapabilityError(
      "preview-unsupported-chart",
      "Chart preview supports only default value data labels",
    );
  }
  if (parsed.title !== undefined && typeof parsed.title !== "string") {
    throw new CapabilityError(
      "preview-unsupported-chart",
      "Chart preview does not support styled titles",
    );
  }

  const baseChartTextStyle = textOnlyChartStyle(parsed.style);
  const { style: resolvedChartStyle } = parseTextContent(
    [],
    baseChartTextStyle,
    context.theme,
    { slideNumber: context.slideNumber },
    "Chart.style",
  );
  // Unit conversion is only for chart spacing; the shared text service owns glyph metrics.
  const size = resolveChartFontSizePixels(resolvedChartStyle.fontSize);

  const parts: string[] = [];
  const renderLabel = (
    content: string,
    target: GeometryBox,
    align: "left" | "center" | "right" = "center",
  ): void => {
    const labelStyle = {
      ...baseChartTextStyle,
      align:
        baseChartTextStyle?.align ?? context.theme.textStyle.align ?? align,
    };
    const { paragraphs, style: resolvedLabelStyle } = parseTextContent(
      [{ type: "#text", text: content }],
      labelStyle,
      context.theme,
      { slideNumber: context.slideNumber },
      "Chart.style",
    );
    const measured = measureTextLayoutFromParagraphs(
      paragraphs,
      resolvedLabelStyle,
    );
    if (
      measured.width > target.width + 0.01 ||
      measured.height > target.height + 0.01
    ) {
      throw new CapabilityError(
        "preview-chart-label-overflow",
        "Preview chart label does not fit its allocated region",
      );
    }
    const lines = previewTextBoxOutlines({
      paragraphs,
      box: target,
      verticalAlign: resolvedLabelStyle.verticalAlign ?? "top",
    });
    if (lines.length > 1) {
      throw new CapabilityError(
        "preview-chart-label-overflow",
        "Preview chart requires single-line labels",
      );
    }
    for (const line of lines) {
      for (const glyph of line.glyphs) {
        parts.push(renderGlyphPathSvg(glyph, context.theme));
      }
    }
  };

  // Reservations are proportional to font size; native PowerPoint/WPS auto-layout remains authoritative in PPTX.
  const titleHeight = parsed.title === undefined ? 0 : size * 2.5;
  const legendHeight = parsed.legend ? size * 2.5 : 0;
  // Reserve full top value-label height above the plot so a maximum bar at the top tick never clips.
  const plotTopInset =
    group.dataLabels !== undefined &&
    group.dataLabels !== false &&
    group.dataLabels.value === true
      ? size * 1.5
      : size;
  const plot = {
    x: box.x + size * 3,
    y: box.y + titleHeight + plotTopInset,
    width: box.width - size * 4,
    height: box.height - titleHeight - legendHeight - plotTopInset - size * 2,
  };
  if (plot.width <= 0 || plot.height <= 0) {
    throw new CapabilityError(
      "preview-chart-too-small",
      "Preview chart frame is too small",
    );
  }

  if (parsed.title !== undefined) {
    renderLabel(parsed.title, {
      x: box.x,
      y: box.y,
      width: box.width,
      height: titleHeight,
    });
  }

  const maximum = Math.max(
    ...group.series.flatMap((series) => series.values as readonly number[]),
  );
  const rough = (maximum || 1) / 5;
  const magnitude = 10 ** Math.floor(Math.log10(rough));
  const step =
    [1, 2, 5, 10].find((candidate) => candidate * magnitude >= rough)! *
    magnitude;
  const top = Math.ceil((maximum || 1) / step) * step;
  if (
    !Number.isFinite(top) ||
    top <= 0 ||
    !Number.isFinite(step) ||
    step <= 0
  ) {
    throw new CapabilityError(
      "preview-chart-invalid-scale",
      "Preview chart values cannot form a finite positive scale",
    );
  }
  const tickCount = Math.round(top / step);
  for (let tick = 0; tick <= tickCount; tick += 1) {
    const value = Number((tick * step).toPrecision(15));
    const y = plot.y + plot.height * (1 - value / top);
    renderLabel(
      String(value),
      { x: box.x, y: y - size / 2, width: size * 2.5, height: size * 1.5 },
      "right",
    );
  }

  const axisColor = context.theme.resolveColor("dk1");
  parts.push(
    `<path d="M${plot.x} ${plot.y}v${plot.height}h${plot.width}" stroke="#${axisColor.hex}" stroke-opacity="${axisColor.alpha ?? 1}" fill="none"/>`,
  );

  const categoryWidth = plot.width / parsed.categories.length;
  // Native gapWidth=150 measures the inter-cluster gap as 150% of one bar width.
  const barWidth = categoryWidth / (group.series.length + 1.5);
  const clusterWidth = barWidth * group.series.length;
  const showValueLabels =
    group.dataLabels !== undefined &&
    group.dataLabels !== false &&
    group.dataLabels.value === true;

  for (let category = 0; category < parsed.categories.length; category += 1) {
    const center = plot.x + categoryWidth * (category + 0.5);
    renderLabel(parsed.categories[category]!, {
      x: center - categoryWidth / 2,
      y: plot.y + plot.height + size / 2,
      width: categoryWidth,
      height: size * 1.5,
    });
    for (let series = 0; series < group.series.length; series += 1) {
      const value = group.series[series]!.values[category]!;
      const barHeight = plot.height * (value / top);
      const x = center - clusterWidth / 2 + series * barWidth;
      const y = plot.y + plot.height - barHeight;
      const color = context.theme.resolveColor(`accent${(series % 6) + 1}`);
      parts.push(
        `<rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" fill="#${color.hex}" fill-opacity="${color.alpha ?? 1}"/>`,
      );
      if (showValueLabels) {
        renderLabel(String(value), {
          x,
          y: y - size * 1.5,
          width: barWidth,
          height: size * 1.5,
        });
      }
    }
  }

  if (parsed.legend) {
    const width = plot.width / group.series.length;
    for (let series = 0; series < group.series.length; series += 1) {
      const x = plot.x + series * width;
      const y = box.y + box.height - legendHeight + size / 2;
      const color = context.theme.resolveColor(`accent${(series % 6) + 1}`);
      parts.push(
        `<rect x="${x}" y="${y}" width="${size}" height="${size}" fill="#${color.hex}" fill-opacity="${color.alpha ?? 1}"/>`,
      );
      renderLabel(
        group.series[series]!.name,
        {
          x: x + size * 1.5,
          y,
          width: width - size * 1.5,
          height: size * 1.5,
        },
        "left",
      );
    }
  }

  return {
    svg: parts.join(""),
    warnings: [
      {
        message:
          "Chart preview uses an approximate automatic layout; plot area, axis scale ticks, margins, and legend spacing may differ from PowerPoint/WPS native auto-layout.",
      },
    ],
  };
}
