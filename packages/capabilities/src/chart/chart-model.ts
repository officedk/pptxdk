import type {
  CapabilityLayoutContract,
  LayoutStyle,
  LinkTarget,
  MorphKey,
} from "@react-pptx/core/capability";

import type { TextStyle } from "../shared/text-service.ts";
import type {
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartDataLabelSettings,
  ChartDisplayBlanksAs,
  ChartGroupData,
  ChartKind,
  ChartProps,
  ChartSeries,
  ChartTitle,
  IndependentCategoryData,
  ScatterChartSeries,
  ScatterData,
  ScatterLine,
} from "./chart-types.ts";

export const chartTextStyleKeys = [
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "color",
  "lineHeight",
  "verticalAlign",
  "align",
  "spacingAfter",
  "letterSpacing",
  "textTransform",
  "textDecoration",
  "textShadow",
  "lang",
] as const;

export function textOnlyChartStyle(value: unknown): TextStyle | undefined {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return undefined;
  }
  const record = value as Record<string, unknown>;
  return Object.fromEntries(
    chartTextStyleKeys.flatMap((key) =>
      record[key] === undefined ? [] : [[key, record[key]]],
    ),
  );
}

export type ParsedChartSeries = ChartSeries & {
  /** Resolved ordered category domain: the shared group domain or this ring's own. */
  readonly categories: readonly string[];
  readonly groupIndex: number;
  readonly groupSeriesIndex: number;
  /** Derived once from group order and group-local series order. */
  readonly seriesIndex: number;
};

export type ParsedScatterSeries = ScatterChartSeries & {
  readonly groupIndex: number;
  readonly groupSeriesIndex: number;
  /** Derived once from group order and group-local series order. */
  readonly seriesIndex: number;
};

/** Native c:crosses/crossesAt on this axis; positions this axis in its crossAx partner domain. */
export type ParsedAxisCrossing =
  | { readonly kind: "crosses"; readonly value: ChartAxisCrossing }
  | { readonly kind: "crossesAt"; readonly value: number };

export type ParsedChartAxis = {
  readonly name: string;
  readonly id: number;
  readonly crossAxisId: number;
  /** Physical direction derived from every group reference to this axis. */
  readonly direction: "horizontal" | "vertical";
  /** Nominal axPos side; interior/autoZero crossings may not resolve to an edge. */
  readonly position: ChartAxisPosition;
  /** Effective native crossing control; explicit author values are preserved. */
  readonly crossing: ParsedAxisCrossing;
  readonly spec: ChartAxis;
};

type ParsedGroupBase = {
  readonly dataLabels?: false | ChartDataLabelSettings;
};

type ParsedCategoryChartGroup = ParsedGroupBase & {
  readonly kind: Exclude<ChartKind, "scatter">;
  readonly data: ChartGroupData | IndependentCategoryData;
  readonly series: readonly ParsedChartSeries[];
  readonly axes?: {
    readonly category: ParsedChartAxis;
    readonly value: ParsedChartAxis;
  };
  readonly stacking?: "stacked" | "percent";
  readonly gapWidth?: number;
  readonly overlap?: number;
  readonly explosion?: number;
  readonly firstSliceAngle?: number;
  readonly holeSize?: number;
};

type ParsedScatterChartGroup = ParsedGroupBase & {
  readonly kind: "scatter";
  readonly data: ScatterData;
  readonly series: readonly ParsedScatterSeries[];
  readonly axes?: {
    readonly x: ParsedChartAxis;
    readonly y: ParsedChartAxis;
  };
  readonly line?: ScatterLine;
  readonly markers: boolean;
};

export type ParsedChartGroup =
  | ParsedCategoryChartGroup
  | ParsedScatterChartGroup;

export type ParsedChart = {
  readonly name?: string;
  readonly morphKey?: MorphKey;
  readonly link?: LinkTarget;
  readonly declaredX?: number;
  readonly declaredY?: number;
  readonly declaredWidth?: number;
  readonly declaredHeight?: number;
  readonly axes: readonly ParsedChartAxis[];
  readonly hasExplicitAxes: boolean;
  readonly groups: readonly ParsedChartGroup[];
  /** Flattened series view with the same objects owned by their groups. */
  readonly series: readonly (ParsedChartSeries | ParsedScatterSeries)[];
  /** Labels of the first category group; each group keeps its own domain. */
  readonly categories: readonly string[];
  /** Largest category point count across category series; scatter groups contribute zero. */
  readonly categoryCount: number;
  readonly title?: ChartTitle;
  readonly displayBlanksAs?: ChartDisplayBlanksAs;
  readonly legend?: ChartProps["legend"];
  readonly plotArea?: ChartProps["plotArea"];
  readonly style?: TextStyle & LayoutStyle;
  readonly contract: CapabilityLayoutContract;
};
