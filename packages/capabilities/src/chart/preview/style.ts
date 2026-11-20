import type {
  ColorValue,
  ResolvedColor,
  ThemeContext,
} from "@react-pptx/core/capability";

import type { ChartLine, ChartLineDash } from "../chart-types.ts";
import { assertNever } from "./types.ts";

/** Resolves a series or point fill color, falling back to the theme accent cycle. */
export function resolveSeriesColor(
  color: ColorValue | undefined,
  seriesIndex: number,
  theme: ThemeContext,
): ResolvedColor {
  return theme.resolveColor(color ?? `accent${(seriesIndex % 6) + 1}`);
}

/** Maps an OOXML preset chart line dash pattern to an SVG stroke-dasharray string. */
export function chartLineDashArray(
  dash: ChartLineDash | undefined,
  width: number,
): string | undefined {
  if (dash === undefined) return undefined;
  const w = Math.max(1, width);
  switch (dash) {
    case "solid":
      return undefined;
    case "dash":
      return `${4 * w} ${3 * w}`;
    case "dot":
      return `${w} ${2 * w}`;
    case "dashDot":
      return `${4 * w} ${2 * w} ${w} ${2 * w}`;
    case "longDash":
      return `${8 * w} ${3 * w}`;
    case "longDashDot":
      return `${8 * w} ${2 * w} ${w} ${2 * w}`;
    case "longDashDotDot":
      return `${8 * w} ${2 * w} ${w} ${2 * w} ${w} ${2 * w}`;
    default:
      return assertNever(dash);
  }
}

export type ResolvedStroke = {
  readonly hex: string;
  readonly alpha: number;
  readonly width: number;
  readonly dashArray?: string;
};

/** Resolves a ChartLine declaration into SVG stroke attributes. */
export function resolveChartLineStroke(
  line: boolean | ChartLine | undefined,
  theme: ThemeContext,
  defaults: {
    readonly color?: ColorValue;
    readonly width?: number;
    readonly alpha?: number;
  } = {},
): ResolvedStroke | undefined {
  if (line === false) return undefined;
  if (line === undefined || line === true) {
    const resolved = theme.resolveColor(defaults.color ?? "dk1");
    return {
      hex: resolved.hex,
      alpha: (resolved.alpha ?? 1) * (defaults.alpha ?? 1),
      width: defaults.width ?? 1,
    };
  }
  const resolved = theme.resolveColor(line.color ?? defaults.color ?? "dk1");
  const width = line.width ?? defaults.width ?? 1;
  const dashArray = chartLineDashArray(line.dash, width);
  return {
    hex: resolved.hex,
    alpha: resolved.alpha ?? 1,
    width,
    ...(dashArray === undefined ? {} : { dashArray }),
  };
}

/** Formats SVG stroke attributes string from a resolved stroke. */
export function strokeAttrs(stroke: ResolvedStroke): string {
  const dash =
    stroke.dashArray === undefined
      ? ""
      : ` stroke-dasharray="${stroke.dashArray}"`;
  return `stroke="#${stroke.hex}" stroke-opacity="${stroke.alpha}" stroke-width="${stroke.width}"${dash}`;
}
