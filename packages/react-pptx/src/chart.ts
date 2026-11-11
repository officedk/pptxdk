import { chartCapability } from "@react-pptx/capabilities/chart";
import { motionCapability } from "@react-pptx/capabilities/motion";
import { transitionCapability } from "@react-pptx/capabilities/transition";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([
  chartCapability,
  motionCapability,
  transitionCapability,
]);

export { Deck, formatBaseline, targets } from "@react-pptx/core";
export { Slide } from "./slide.ts";
export { Chart } from "@react-pptx/capabilities/chart-elements";
export { chartCapability } from "@react-pptx/capabilities/chart";
export { Step } from "@react-pptx/capabilities/motion-elements";
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
} from "@react-pptx/capabilities/motion";
export {
  transitionCapability,
  transitions,
} from "@react-pptx/capabilities/transition";
export { createBuilder } from "@react-pptx/core";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  Diagnostic,
  ObjectObservation,
  SlideObservation,
  Target,
} from "@react-pptx/core";
export type {
  ChartAxis,
  ChartAxisCrossing,
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
} from "@react-pptx/capabilities/chart";
export type { SlideProps } from "./slide.ts";
