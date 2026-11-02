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
export { Chart, chartCapability } from "@react-pptx/capabilities/chart";
export {
  Step,
  fade,
  fly,
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
  ChartData,
  ChartDataLabels,
  ChartKind,
  ChartLine,
  ChartPoint,
  ChartProps,
  ChartSeries,
  ColumnOrBarChartProps,
  LineOrAreaChartProps,
} from "@react-pptx/capabilities/chart";
export type { SlideProps } from "./slide.ts";
