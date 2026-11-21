import type { GeometryBox } from "@pptxdk/core/capability";

import type { ParsedCategoryAxis, ParsedValueAxis } from "../chart-model.ts";
import type { ChartMarkerSymbol } from "../chart-types.ts";

/** Plotted anchor emitted by column/bar renderers for label and annotation placement. */
export type PlottedBarAnchor = {
  readonly kind: "bar";
  readonly groupIndex: number;
  readonly seriesIndex: number;
  readonly groupSeriesIndex: number;
  readonly pointIndex: number;
  readonly value: number;
  readonly plottedValue?: number;
  readonly category: string;
  readonly seriesName: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly isNegative: boolean;
  readonly horizontal: boolean;
  readonly valueEndSide: "top" | "bottom" | "left" | "right";
  readonly percent?: number;
};

/** Plotted anchor emitted by line, area, and scatter point renderers. */
export type PlottedCartesianPointAnchor = {
  readonly kind: "point";
  readonly groupIndex: number;
  readonly seriesIndex: number;
  readonly groupSeriesIndex: number;
  readonly pointIndex: number;
  readonly value: number;
  readonly plottedValue?: number;
  readonly xValue?: number;
  readonly category?: string;
  readonly seriesName: string;
  readonly x: number;
  readonly y: number;
  readonly markerRadius?: number;
  readonly percent?: number;
};

/** Plotted anchor emitted by pie and doughnut slice renderers. */
export type PlottedSliceAnchor = {
  readonly kind: "slice";
  readonly groupIndex: number;
  readonly seriesIndex: number;
  readonly groupSeriesIndex: number;
  readonly pointIndex: number;
  readonly value: number;
  readonly plottedValue?: number;
  readonly category: string;
  readonly seriesName: string;
  readonly cx: number;
  readonly cy: number;
  readonly innerR: number;
  readonly outerR: number;
  readonly startAngleRad: number;
  readonly endAngleRad: number;
  readonly midAngleRad: number;
  readonly sliceSum: number;
  readonly percent: number;
  readonly spanDeg: number;
};

export type PlottedPointAnchor =
  | PlottedBarAnchor
  | PlottedCartesianPointAnchor
  | PlottedSliceAnchor;

export type CategoryAxisScale = {
  readonly kind: "category";
  readonly axis: ParsedCategoryAxis;
  readonly categories: readonly string[];
  readonly count: number;
  readonly crossBetween: "between" | "midCat";
  readonly reversed: boolean;
  mapIndex(index: number): number;
  mapBoundary(boundary: number): number;
  mapScreen(t: number, plot: GeometryBox): number;
  indexToScreen(index: number, plot: GeometryBox): number;
  slotSpan(plot: GeometryBox): number;
};

export type ValueAxisScale = {
  readonly kind: "value";
  readonly axis: ParsedValueAxis;
  readonly min: number;
  readonly max: number;
  readonly logBase?: number;
  readonly majorTicks: readonly number[];
  readonly minorTicks: readonly number[];
  readonly reversed: boolean;
  readonly numberFormat: string | undefined;
  mapValue(value: number): number;
  mapScreen(t: number, plot: GeometryBox): number;
  valueToScreen(value: number, plot: GeometryBox): number;
  zeroValue(): number;
  zeroScreen(plot: GeometryBox): number;
};

export type AxisScale = CategoryAxisScale | ValueAxisScale;

export type ChartLegendMarker = {
  readonly symbol: Exclude<ChartMarkerSymbol, "none" | "auto">;
  readonly fillHex: string;
  readonly fillAlpha: number;
  readonly strokeHex?: string;
  readonly strokeAlpha?: number;
  readonly strokeWidth?: number;
};

export type ChartLegendEntry = {
  readonly label: string;
  readonly colorHex: string;
  readonly colorAlpha: number;
  readonly kind: "swatch" | "line" | "marker";
  readonly strokeHex?: string;
  readonly strokeAlpha?: number;
  readonly strokeWidth?: number;
  readonly dashArray?: string;
  readonly marker?: ChartLegendMarker;
};

export type ChartLayoutResult = {
  readonly plotRect: GeometryBox;
  readonly titleBox?: GeometryBox;
  readonly legendBox?: GeometryBox;
  readonly axisTitleBoxes: ReadonlyMap<string, GeometryBox>;
  readonly labelZones: ReadonlyMap<
    string,
    {
      readonly side: "top" | "bottom" | "left" | "right";
      readonly depth: number;
      readonly offset: number;
    }
  >;
};

export function assertNever(value: never): never {
  throw new Error(`Unhandled chart preview variant: ${JSON.stringify(value)}`);
}
