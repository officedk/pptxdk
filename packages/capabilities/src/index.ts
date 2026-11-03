export {
  Shape,
  blockArc,
  linearGradient,
  path,
  preset,
  radialGradient,
  shapeCapability,
} from "#/shape.ts";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
  ShapeProps,
} from "#/shape.ts";
export { boxCapability } from "#/box.ts";
export type {
  GradientStop,
  ImageVisualStyle,
  LinearGradient,
  RadialGradient,
  Shadow,
  VisualStyle,
} from "@react-pptx/core";
export { Chart, chartCapability } from "#/chart.ts";
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
  DoughnutChartProps,
  LineOrAreaChartProps,
  PieChartProps,
} from "#/chart.ts";
export { Image, imageCapability } from "#/image.ts";
export type { ImageProps, ImageSource } from "#/image.ts";
export { Group, groupCapability } from "#/group.ts";
export type { GroupProps } from "#/group.ts";
export { Cell, Row, Table, tableCapability } from "#/table.ts";
export type {
  CellProps,
  RowProps,
  TableCellStyle,
  TableColumn,
  TableProps,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
} from "#/table.ts";
export { font, P, SlideNumber, Span, Text, textCapability } from "#/text.ts";
export { transitionCapability, transitions } from "#/transition.ts";
export { slideBackgroundCapability } from "#/slide-background.ts";
export { createMorphScope } from "#/morph-api.ts";
export type { MorphKey, MorphProps, MorphScope } from "#/morph-api.ts";
export {
  after,
  Box,
  Step,
  MotionScope,
  fade,
  fly,
  floatUp,
  motionCapability,
  spin,
  spinIn,
  stretch,
  useMotionScope,
  wipe,
  zoom,
} from "#/motion.ts";
export type {
  BoxProps,
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EntranceEffect,
  MotionDirection,
  MotionGranularity,
  MotionCompletionState,
  MotionEffectObservation,
  MotionMoment,
  MotionMomentOffset,
  MotionObservation,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeObservation,
  MotionScopeProps,
  MotionProps,
  MotionStartDependency,
  MotionTargetRange,
  MotionTiming,
  MotionStepObservation,
  StepProps,
} from "#/motion.ts";
export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  ParagraphProps,
  ParagraphStyle,
  SlideNumberProps,
  ShapeTextStyle,
  SpanProps,
  TextAlignment,
  TextList,
  TextProps,
  TextStyle,
} from "#/text.ts";
export type {
  FlipTransition,
  MorphTransition,
  StandardTransition,
  StripsDirection,
  Transition,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
} from "#/transition.ts";
