import { chartCapability } from "@pptxdk/capabilities/chart";
import { motionCapability } from "@pptxdk/capabilities/motion";
import { transitionCapability } from "@pptxdk/capabilities/transition";

import { createStagePipeline } from "./compose.ts";

export const { build } = createStagePipeline([
  chartCapability,
  motionCapability,
  transitionCapability,
]);

export { formatBaseline, targets } from "@pptxdk/core";
export { Deck } from "@pptxdk/react";
export { Slide } from "./slide.ts";
export { Chart } from "@pptxdk/capabilities/chart-elements";
export { chartCapability } from "@pptxdk/capabilities/chart";
export { Step } from "@pptxdk/capabilities/motion-elements";
export {
  fade,
  fly,
  floatDown,
  floatUp,
  motionCapability,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "@pptxdk/capabilities/motion";
export {
  transitionCapability,
  transitions,
} from "@pptxdk/capabilities/transition";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  Diagnostic,
  ObjectObservation,
  SlideObservation,
  Target,
} from "@pptxdk/core";
export type {
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartAxisReferences,
  ChartCategoryAxis,
  ChartCategoryAxisReferences,
  ChartDataLabelOverride,
  ChartCustomErrorBars,
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
} from "@pptxdk/capabilities/chart";
export type { SlideProps } from "./slide.ts";
