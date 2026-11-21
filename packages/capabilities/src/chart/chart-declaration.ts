import { CommonNode } from "@pptxdk/core/document";

import type { ChartProps } from "./chart-types.ts";

export type ChartNode = CommonNode<
  "chart",
  ChartProps,
  NonNullable<ChartProps["style"]>
>;

/** Creates the editable native chart shared by the command factory and the Chart JSX element. */
export function createChartNode(props: ChartProps): ChartNode {
  return new CommonNode<"chart", ChartProps, NonNullable<ChartProps["style"]>>(
    "chart",
    "rp:chart",
    props,
  );
}

export type {
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartAxisReferences,
  ChartCategoryAxis,
  ChartCategoryAxisReferences,
  ChartCustomErrorBars,
  ChartDataLabelOverride,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartErrorBarInclude,
  ChartErrorBars,
  ChartErrorBarValues,
  ChartGroup,
  ChartGroupData,
  ChartKind,
  ChartLine,
  ChartLineDash,
  ChartMarker,
  ChartMarkerSymbol,
  ChartPoint,
  ChartProps,
  ChartSeries,
  ChartSeriesDataLabelOverride,
  ChartTextStyle,
  ChartTickLabelPosition,
  ChartTickMark,
  ChartTitle,
  ChartTrendline,
  ChartTrendlineType,
  ChartValueAxis,
  ChartXYAxisReferences,
  ColumnOrBarChartGroup,
  DoughnutChartGroup,
  IndependentCategoryData,
  IndependentCategorySeries,
  LineOrAreaChartGroup,
  PieChartGroup,
  ScatterChartGroup,
  ScatterChartPoint,
  ScatterChartSeries,
  ScatterData,
  ScatterDataLabelOverride,
  ScatterDataLabels,
  ScatterErrorBars,
  ScatterLine,
  SharedCategoryData,
} from "./chart-types.ts";
