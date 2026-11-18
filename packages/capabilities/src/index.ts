export {
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
  shapeStyleKeys,
} from "#/shape.ts";
export { Shape } from "#/shape-elements.ts";
export type { ShapeProps } from "#/shape-elements.ts";
export { boxCapability } from "#/box.ts";
export { Box, type BoxProps } from "#/box-elements.ts";
export {
  arrowTypes,
  hasVisibleBoxStyle,
  lineEndStyleKeys,
  type ArrowType,
  type LineEndStyle,
  type ShapeVisualStyle,
} from "#/visual.ts";
export type {
  GradientStop,
  LinearGradient,
  RadialGradient,
  Shadow,
  VisualStyle,
} from "@react-pptx/core";
export { chartCapability } from "#/chart.ts";
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
  ChartNode,
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
  ParsedChart,
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
} from "#/chart.ts";
export { Chart } from "#/chart-elements.ts";
export { imageCapability } from "#/image.ts";
export type {
  ImageDeclarationProps,
  ImageSource,
  ImageVisualStyle,
} from "#/image.ts";
export { Image } from "#/image-elements.ts";
export type { ImageProps } from "#/image-elements.ts";
export { groupCapability } from "#/group.ts";
export type { GroupDeclarationProps } from "#/group.ts";
export { Group } from "#/group-elements.ts";
export type { GroupProps } from "#/group-elements.ts";
export { tableCapability } from "#/table.ts";
export type {
  CellDeclarationProps,
  CellNode,
  RowDeclarationProps,
  RowNode,
  TableCellStyle,
  TableColumn,
  TableDeclarationProps,
  TableNode,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
} from "#/table.ts";
export { Cell, Row, Table } from "#/table-elements.ts";
export type { CellProps, RowProps, TableProps } from "#/table-elements.ts";
export { font, textCapability } from "#/text.ts";
export { P, SlideNumber, Span, Text } from "#/text-elements.ts";
export { notesCapability } from "#/notes.ts";
export { Notes } from "#/notes-elements.ts";
export type { NotesProps } from "#/notes-elements.ts";
export type { NotesDeclarationProps, NotesNode } from "#/notes-declaration.ts";
export { transitionCapability, transitions } from "#/transition.ts";
export { slideBackgroundCapability } from "#/slide-background.ts";
export { createMorphScope } from "#/morph-api.ts";
export type { MorphKey, MorphProps, MorphScope } from "#/morph-api.ts";
export {
  after,
  breathe,
  fillColor,
  observeMotion,
  createMotionScopeHandle,
  fade,
  fly,
  floatDown,
  floatUp,
  motionCapability,
  repeat,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "#/motion.ts";
export { MotionScope, Step, useMotionScope } from "#/motion-elements.ts";
export type {
  EffectInstanceObservation,
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  BreatheEffect,
  BreatheOptions,
  FillColorOptions,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffectDeclaration,
  EntranceEffect,
  MotionDirection,
  MotionGranularity,
  MotionCompletionState,
  MotionCueObservation,
  MotionCueRef,
  MotionDecision,
  MotionDependencyObservation,
  MotionDifference,
  MotionEffectRef,
  MotionEffectObservation,
  MotionIntervalRef,
  MotionMembership,
  MotionMomentObservation,
  MotionMomentRef,
  MotionNativeProjection,
  MotionMoment,
  MotionMomentOffset,
  MotionObservation,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeObservation,
  MotionScopeQueryMember,
  MotionScopeQueryObservation,
  MotionScopeRef,
  MotionScopeProps,
  MotionProps,
  MotionStartDependency,
  MotionTargetRange,
  MotionTiming,
  MotionStepObservation,
  MotionStepQuery,
  MotionStepRef,
  MotionTime,
  MotionQuery,
  RepeatCount,
  RepeatedEffect,
  SpinOptions,
  StepProps,
} from "#/motion.ts";
export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  ParagraphDeclarationProps,
  ParagraphStyle,
  ShapeTextStyle,
  SlideNumberDeclarationProps,
  SpanDeclarationProps,
  TextAlignment,
  TextDeclarationProps,
  TextList,
  TextStyle,
} from "#/text.ts";
export type {
  ParagraphProps,
  SlideNumberProps,
  SpanProps,
  TextProps,
} from "#/text-elements.ts";
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
