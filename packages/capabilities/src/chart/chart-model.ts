import type {
  CapabilityLayoutContract,
  LayoutStyle,
  LinkTarget,
  MorphKey,
} from "@pptxdk/core/capability";

import type { TextStyle } from "../shared/text-service.ts";
import { chartTextStyleSchema } from "./chart-schemas.ts";
import type {
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartCategoryAxis,
  ChartDataLabelSettings,
  ChartDisplayBlanksAs,
  ChartProps,
  ChartSeries,
  ChartTitle,
  ChartValueAxis,
  IndependentCategoryData,
  ScatterChartSeries,
  ScatterData,
  ScatterLine,
  SharedCategoryData,
} from "./chart-types.ts";

export const chartTextStyleKeys = Object.keys(
  chartTextStyleSchema.unwrap().shape,
);

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

export type ParsedCategoryAxis = {
  readonly kind: "category";
  readonly name: string;
  readonly id: number;
  readonly crossAxisId: number;
  /** Physical direction derived from every group reference to this axis. */
  readonly direction: "horizontal" | "vertical";
  /** Nominal axPos side; interior/autoZero crossings may not resolve to an edge. */
  readonly position: ChartAxisPosition;
  /** Effective native crossing control; explicit author values are preserved. */
  readonly crossing: ParsedAxisCrossing;
  readonly spec: ChartCategoryAxis;
};

export type ParsedValueAxis = {
  readonly kind: "value";
  readonly name: string;
  readonly id: number;
  readonly crossAxisId: number;
  /** Physical direction derived from every group reference to this axis. */
  readonly direction: "horizontal" | "vertical";
  /** Nominal axPos side; interior/autoZero crossings may not resolve to an edge. */
  readonly position: ChartAxisPosition;
  /** Effective native crossing control; explicit author values are preserved. */
  readonly crossing: ParsedAxisCrossing;
  readonly spec: ChartValueAxis;
};

export type ParsedChartAxis = ParsedCategoryAxis | ParsedValueAxis;

export type ResolvedCartesianAxes = {
  readonly category: ParsedCategoryAxis;
  readonly value: ParsedValueAxis;
};

export type ResolvedScatterAxes = {
  readonly x: ParsedValueAxis;
  readonly y: ParsedValueAxis;
};

type ParsedGroupBase = {
  readonly dataLabels?: false | ChartDataLabelSettings;
};

export type ParsedColumnOrBarGroup = ParsedGroupBase & {
  readonly kind: "column" | "bar";
  readonly data: SharedCategoryData;
  readonly series: readonly ParsedChartSeries[];
  readonly axes: ResolvedCartesianAxes;
  readonly stacking: "clustered" | "stacked" | "percent";
  readonly gapWidth: number;
  readonly overlap?: number;
};

export type ParsedLineGroup = ParsedGroupBase & {
  readonly kind: "line";
  readonly data: SharedCategoryData;
  readonly series: readonly ParsedChartSeries[];
  readonly axes: ResolvedCartesianAxes;
  readonly stacking: "standard" | "stacked" | "percent";
};

export type ParsedAreaGroup = ParsedGroupBase & {
  readonly kind: "area";
  readonly data: SharedCategoryData;
  readonly series: readonly ParsedChartSeries[];
  readonly axes: ResolvedCartesianAxes;
  readonly stacking: "standard" | "stacked" | "percent";
};

export type ParsedPieGroup = ParsedGroupBase & {
  readonly kind: "pie";
  readonly data: SharedCategoryData;
  readonly series: readonly [ParsedChartSeries];
  readonly firstSliceAngle: number;
  readonly explosion?: number;
};

export type ParsedDoughnutGroup = ParsedGroupBase & {
  readonly kind: "doughnut";
  readonly data: SharedCategoryData | IndependentCategoryData;
  readonly series: readonly ParsedChartSeries[];
  readonly firstSliceAngle: number;
  readonly holeSize: number;
  readonly explosion?: number;
};

export type ParsedScatterGroup = ParsedGroupBase & {
  readonly kind: "scatter";
  readonly data: ScatterData;
  readonly series: readonly ParsedScatterSeries[];
  readonly axes: ResolvedScatterAxes;
  readonly line: ScatterLine;
  readonly markers: boolean;
};

export type ParsedChartGroup =
  | ParsedColumnOrBarGroup
  | ParsedLineGroup
  | ParsedAreaGroup
  | ParsedPieGroup
  | ParsedDoughnutGroup
  | ParsedScatterGroup;

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
