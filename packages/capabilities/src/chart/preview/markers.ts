import type { ColorValue, ThemeContext } from "@react-pptx/core/capability";

import type { ChartMarker, ChartMarkerSymbol } from "../chart-types.ts";
import {
  resolveChartLineStroke,
  resolveSeriesColor,
  strokeAttrs,
  type ResolvedStroke,
} from "./style.ts";
import { assertNever } from "./types.ts";

export type VisibleMarkerSymbol = Exclude<ChartMarkerSymbol, "none" | "auto">;

const DEFAULT_MARKER_SIZE_PT = 5;

const AUTO_MARKER_SYMBOLS: readonly VisibleMarkerSymbol[] = [
  "diamond",
  "square",
  "triangle",
  "x",
  "star",
  "circle",
  "plus",
  "dot",
  "dash",
];

export type ResolvedMarker = {
  readonly symbol: VisibleMarkerSymbol;
  readonly radius: number;
  readonly fillHex: string;
  readonly fillAlpha: number;
  readonly stroke?: ResolvedStroke;
};

function asMarkerObject(
  value: boolean | ChartMarker | undefined,
): ChartMarker | undefined {
  return typeof value === "object" && value !== null ? value : undefined;
}

/**
 * Resolves effective marker geometry and styling by merging point-level overrides over series
 * and group defaults, converting size from points to pixel radius.
 */
export function resolveEffectiveMarker(options: {
  readonly groupDefaultEnabled: boolean;
  readonly seriesMarker: boolean | ChartMarker | undefined;
  readonly pointMarker: boolean | ChartMarker | undefined;
  readonly seriesColor: ColorValue | undefined;
  readonly pointColor: ColorValue | undefined;
  readonly seriesIndex: number;
  readonly theme: ThemeContext;
}): ResolvedMarker | undefined {
  const {
    groupDefaultEnabled,
    seriesMarker,
    pointMarker,
    seriesColor,
    pointColor,
    seriesIndex,
    theme,
  } = options;

  let enabled: boolean;
  if (pointMarker === false) {
    enabled = false;
  } else if (pointMarker !== undefined) {
    enabled = true;
  } else if (seriesMarker === false) {
    enabled = false;
  } else if (seriesMarker !== undefined) {
    enabled = true;
  } else {
    enabled = groupDefaultEnabled;
  }

  if (!enabled) return undefined;

  const seriesObj = asMarkerObject(seriesMarker);
  const pointObj = asMarkerObject(pointMarker);

  const rawSymbol =
    pointMarker === true
      ? (seriesObj?.symbol ?? "circle")
      : (pointObj?.symbol ?? seriesObj?.symbol ?? "circle");

  if (rawSymbol === "none") return undefined;

  const symbol: VisibleMarkerSymbol =
    rawSymbol === "auto"
      ? AUTO_MARKER_SYMBOLS[seriesIndex % AUTO_MARKER_SYMBOLS.length]!
      : rawSymbol;

  const sizePt = pointObj?.size ?? seriesObj?.size ?? DEFAULT_MARKER_SIZE_PT;
  const radius = (sizePt * (4 / 3)) / 2;

  const fillColor = resolveSeriesColor(
    pointObj?.fill ?? seriesObj?.fill ?? pointColor ?? seriesColor,
    seriesIndex,
    theme,
  );
  const fillHex = fillColor.hex;
  const fillAlpha = fillColor.alpha ?? 1;

  const lineSpec = pointObj?.line ?? seriesObj?.line;
  const explicitStroke =
    lineSpec === undefined
      ? undefined
      : resolveChartLineStroke(lineSpec, theme, {
          color: "dk1",
          width: 1,
        });

  const stroke: ResolvedStroke | undefined =
    symbol === "plus" || symbol === "x" || symbol === "star"
      ? (explicitStroke ?? {
          hex: fillHex,
          alpha: fillAlpha,
          width: Math.max(1.25, radius * 0.35),
        })
      : explicitStroke;

  return {
    symbol,
    radius,
    fillHex,
    fillAlpha,
    ...(stroke === undefined ? {} : { stroke }),
  };
}

/**
 * Renders one SVG marker glyph centered at `(cx, cy)` covering all 11 OOXML marker symbols.
 */
export function renderMarkerSvg(
  cx: number,
  cy: number,
  marker: ResolvedMarker,
): string {
  const r = marker.radius;
  const fillAttr = `fill="#${marker.fillHex}" fill-opacity="${marker.fillAlpha}"`;
  const strokeAttr =
    marker.stroke === undefined ? "" : ` ${strokeAttrs(marker.stroke)}`;

  switch (marker.symbol) {
    case "circle":
      return `<circle cx="${cx}" cy="${cy}" r="${r}" ${fillAttr}${strokeAttr}/>`;
    case "square":
      return `<rect x="${cx - r}" y="${cy - r}" width="${2 * r}" height="${2 * r}" ${fillAttr}${strokeAttr}/>`;
    case "diamond":
      return `<path d="M${cx} ${cy - r}L${cx + r} ${cy}L${cx} ${cy + r}L${cx - r} ${cy}Z" ${fillAttr}${strokeAttr}/>`;
    case "triangle":
      return `<path d="M${cx} ${cy - r}L${cx + r} ${cy + r}L${cx - r} ${cy + r}Z" ${fillAttr}${strokeAttr}/>`;
    case "plus": {
      const stroke = marker.stroke ?? {
        hex: marker.fillHex,
        alpha: marker.fillAlpha,
        width: Math.max(1.25, r * 0.35),
      };
      return `<path d="M${cx - r} ${cy}H${cx + r}M${cx} ${cy - r}V${cy + r}" fill="none" ${strokeAttrs(stroke)} stroke-linecap="round"/>`;
    }
    case "x": {
      const stroke = marker.stroke ?? {
        hex: marker.fillHex,
        alpha: marker.fillAlpha,
        width: Math.max(1.25, r * 0.35),
      };
      return `<path d="M${cx - r} ${cy - r}L${cx + r} ${cy + r}M${cx - r} ${cy + r}L${cx + r} ${cy - r}" fill="none" ${strokeAttrs(stroke)} stroke-linecap="round"/>`;
    }
    case "star": {
      const stroke = marker.stroke ?? {
        hex: marker.fillHex,
        alpha: marker.fillAlpha,
        width: Math.max(1.25, r * 0.35),
      };
      return `<path d="M${cx - r} ${cy - r}L${cx + r} ${cy + r}M${cx - r} ${cy + r}L${cx + r} ${cy - r}M${cx} ${cy - r}V${cy + r}" fill="none" ${strokeAttrs(stroke)} stroke-linecap="round"/>`;
    }
    case "dot":
      return `<rect x="${cx}" y="${cy - r * 0.25}" width="${r}" height="${r * 0.5}" ${fillAttr}${strokeAttr}/>`;
    case "dash":
      return `<rect x="${cx - r}" y="${cy - r * 0.25}" width="${2 * r}" height="${r * 0.5}" ${fillAttr}${strokeAttr}/>`;
    default:
      return assertNever(marker.symbol);
  }
}
