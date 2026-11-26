import type { Rect, ThemeContext } from "@pptxdk/core/capability";

import { mergeDataLabels } from "../chart-labels.ts";
import type { ParsedDoughnutGroup, ParsedPieGroup } from "../chart-model.ts";
import {
  resolveChartLineStroke,
  resolveSeriesColor,
  strokeAttrs,
} from "./style.ts";
import type { PlottedSliceAnchor } from "./types.ts";

function toScreenRadians(degreesClockwiseFromTop: number): number {
  return ((degreesClockwiseFromTop - 90) * Math.PI) / 180;
}

function hasOuterSliceLabels(
  group: ParsedPieGroup | ParsedDoughnutGroup,
): boolean {
  if (group.kind === "doughnut") return false;
  const series = group.series[0];
  for (let i = 0; i < series.values.length; i += 1) {
    const val = series.values[i];
    if (
      val === null ||
      val === undefined ||
      !Number.isFinite(val) ||
      val <= 0
    ) {
      continue;
    }
    const eff = mergeDataLabels(
      mergeDataLabels(group.dataLabels, series.dataLabels),
      series.points?.[i]?.dataLabels,
    );
    if (eff === undefined || eff === false) continue;
    const hasContent =
      eff.value === true ||
      eff.percent === true ||
      eff.category === true ||
      eff.seriesName === true;
    if (!hasContent) continue;
    const pos = eff.position ?? "bestFit";
    if (pos === "outsideEnd" || pos === "bestFit") {
      return true;
    }
  }
  return false;
}

function buildSlicePath(options: {
  readonly ox: number;
  readonly oy: number;
  readonly innerR: number;
  readonly outerR: number;
  readonly startRad: number;
  readonly endRad: number;
  readonly spanDeg: number;
}): { readonly d: string; readonly fullCircle: boolean } {
  const { ox, oy, innerR, outerR, startRad, endRad, spanDeg } = options;

  // Split 360-degree slices into two 180-degree arcs so SVG A does not collapse coincident endpoints.
  if (spanDeg >= 359.999) {
    const oppRad = startRad + Math.PI;
    const ox1 = ox + outerR * Math.cos(startRad);
    const oy1 = oy + outerR * Math.sin(startRad);
    const ox2 = ox + outerR * Math.cos(oppRad);
    const oy2 = oy + outerR * Math.sin(oppRad);
    if (innerR <= 0.001) {
      return {
        d: `M${ox1} ${oy1}A${outerR} ${outerR} 0 1 1 ${ox2} ${oy2}A${outerR} ${outerR} 0 1 1 ${ox1} ${oy1}Z`,
        fullCircle: true,
      };
    }
    const ix1 = ox + innerR * Math.cos(startRad);
    const iy1 = oy + innerR * Math.sin(startRad);
    const ix2 = ox + innerR * Math.cos(oppRad);
    const iy2 = oy + innerR * Math.sin(oppRad);
    return {
      d: `M${ox1} ${oy1}A${outerR} ${outerR} 0 1 1 ${ox2} ${oy2}A${outerR} ${outerR} 0 1 1 ${ox1} ${oy1}M${ix1} ${iy1}A${innerR} ${innerR} 0 1 0 ${ix2} ${iy2}A${innerR} ${innerR} 0 1 0 ${ix1} ${iy1}Z`,
      fullCircle: true,
    };
  }

  const largeArc = spanDeg > 180 ? 1 : 0;
  const x1 = ox + outerR * Math.cos(startRad);
  const y1 = oy + outerR * Math.sin(startRad);
  const x2 = ox + outerR * Math.cos(endRad);
  const y2 = oy + outerR * Math.sin(endRad);

  if (innerR <= 0.001) {
    return {
      d: `M${ox} ${oy}L${x1} ${y1}A${outerR} ${outerR} 0 ${largeArc} 1 ${x2} ${y2}Z`,
      fullCircle: false,
    };
  }

  const x3 = ox + innerR * Math.cos(endRad);
  const y3 = oy + innerR * Math.sin(endRad);
  const x4 = ox + innerR * Math.cos(startRad);
  const y4 = oy + innerR * Math.sin(startRad);
  return {
    d: `M${x1} ${y1}A${outerR} ${outerR} 0 ${largeArc} 1 ${x2} ${y2}L${x3} ${y3}A${innerR} ${innerR} 0 ${largeArc} 0 ${x4} ${y4}Z`,
    fullCircle: false,
  };
}

/**
 * Renders pie and single/multi-ring doughnut chart groups with explosion, holeSize,
 * custom starting angle, and slice anchors for data labels and leader lines.
 */
export function renderPieOrDoughnutGroup(options: {
  readonly group: ParsedPieGroup | ParsedDoughnutGroup;
  readonly groupIndex: number;
  readonly plotRect: Rect;
  readonly theme: ThemeContext;
}): {
  readonly svg: string;
  readonly anchors: readonly PlottedSliceAnchor[];
} {
  const { group, groupIndex, plotRect, theme } = options;
  if (
    group.series.length === 0 ||
    plotRect.width <= 0 ||
    plotRect.height <= 0
  ) {
    return { svg: "", anchors: [] };
  }

  let maxExpRatio = 0;
  for (const series of group.series) {
    for (let j = 0; j < series.values.length; j += 1) {
      const val = series.values[j];
      if (
        val === null ||
        val === undefined ||
        !Number.isFinite(val) ||
        val <= 0
      ) {
        continue;
      }
      const exp =
        series.points?.[j]?.explosion ??
        series.explosion ??
        group.explosion ??
        0;
      const ratio = Math.min(Math.max(0, exp), 100) / 100;
      if (ratio > maxExpRatio) maxExpRatio = ratio;
    }
  }

  const outerLabels = hasOuterSliceLabels(group);
  const radiusScale = outerLabels ? 0.68 : 0.88;
  const rAvail = (Math.min(plotRect.width, plotRect.height) / 2) * radiusScale;
  const rOuterTotal = Math.max(2, rAvail / (1 + maxExpRatio));
  const cx = plotRect.x + plotRect.width / 2;
  const cy = plotRect.y + plotRect.height / 2;

  const ringCount = group.kind === "pie" ? 1 : group.series.length;
  const holeRatio =
    group.kind === "doughnut"
      ? Math.max(0.1, Math.min(0.9, group.holeSize / 100))
      : 0;
  const rHole = rOuterTotal * holeRatio;
  const ringThickness = (rOuterTotal - rHole) / Math.max(1, ringCount);

  const parts: string[] = [];
  const anchors: PlottedSliceAnchor[] = [];

  for (let s = 0; s < group.series.length; s += 1) {
    const series = group.series[s]!;
    const rIn = group.kind === "pie" ? 0 : rHole + s * ringThickness;
    const rOut =
      group.kind === "pie" ? rOuterTotal : rHole + (s + 1) * ringThickness;

    let ringTotal = 0;
    let positiveCount = 0;
    for (const v of series.values) {
      if (v !== null && v !== undefined && Number.isFinite(v) && v > 0) {
        ringTotal += v;
        positiveCount += 1;
      }
    }
    if (ringTotal <= 0) continue;

    let currentDeg = group.firstSliceAngle;
    for (let j = 0; j < series.values.length; j += 1) {
      const val = series.values[j];
      if (
        val === null ||
        val === undefined ||
        !Number.isFinite(val) ||
        val <= 0
      ) {
        continue;
      }

      const spanDeg = (val / ringTotal) * 360;
      const startDeg = currentDeg;
      const endDeg = startDeg + spanDeg;
      const midDeg = startDeg + spanDeg / 2;
      currentDeg = endDeg;

      const startRad = toScreenRadians(startDeg);
      const endRad = toScreenRadians(endDeg);
      const midRad = toScreenRadians(midDeg);

      const point = series.points?.[j];
      const exp = point?.explosion ?? series.explosion ?? group.explosion ?? 0;
      const expOffset = rOuterTotal * (Math.min(Math.max(0, exp), 100) / 100);
      const ox = cx + expOffset * Math.cos(midRad);
      const oy = cy + expOffset * Math.sin(midRad);

      const { d, fullCircle } = buildSlicePath({
        ox,
        oy,
        innerR: rIn,
        outerR: rOut,
        startRad,
        endRad,
        spanDeg,
      });

      const fill = resolveSeriesColor(point?.color ?? series.color, j, theme);
      const lineSpec = point?.line ?? series.line;
      const stroke =
        lineSpec !== undefined
          ? resolveChartLineStroke(lineSpec, theme, { color: "dk1", width: 1 })
          : positiveCount > 1 || ringCount > 1
            ? { hex: "FFFFFF", alpha: 0.9, width: 1 }
            : undefined;
      const strokeAttr = stroke === undefined ? "" : ` ${strokeAttrs(stroke)}`;
      const fillRuleAttr =
        fullCircle && rIn > 0.001 ? ' fill-rule="evenodd"' : "";

      parts.push(
        `<path d="${d}" fill="#${fill.hex}" fill-opacity="${fill.alpha ?? 1}"${fillRuleAttr}${strokeAttr}/>`,
      );

      anchors.push({
        kind: "slice",
        groupIndex,
        seriesIndex: series.seriesIndex,
        groupSeriesIndex: s,
        pointIndex: j,
        value: val,
        category: series.categories[j] ?? "",
        seriesName: series.name,
        cx: ox,
        cy: oy,
        innerR: rIn,
        outerR: rOut,
        startAngleRad: startRad,
        endAngleRad: endRad,
        midAngleRad: midRad,
        sliceSum: ringTotal,
        percent: val / ringTotal,
        spanDeg,
      });
    }
  }

  return { svg: parts.join(""), anchors };
}
