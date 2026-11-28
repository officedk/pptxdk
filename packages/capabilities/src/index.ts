export {
  blockArc,
  linearGradient,
  path,
  preset,
  presetGeometryNames,
  presetGeometryTuple,
  radialGradient,
  shapeCapability,
} from "#/shape/shape.ts";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
  PresetGeometryName,
  shapeStyleKeys,
} from "#/shape/shape.ts";
export { Shape } from "#/shape/shape-elements.ts";
export type { ShapeProps } from "#/shape/shape-elements.ts";
export { lineCapability } from "#/line/line.ts";
export type {
  LineAdjust,
  LineDeclarationProps,
  LineEndpoint,
  LineLayoutStyle,
  LinePort,
  LineRouting,
  LineStroke,
  LineStrokeValue,
  LineStyle,
  LineVisualStyle,
} from "#/line/line.ts";
export { Line } from "#/line/line-elements.ts";
export type { LineProps } from "#/line/line-elements.ts";
export { boxCapability } from "#/box/box.ts";
export { Box, type BoxProps } from "#/box/box-elements.ts";
export {
  arrowTypes,
  hasVisibleBoxStyle,
  type ArrowType,
} from "#/shared/visual.ts";
export type {
  GradientStop,
  LinearGradient,
  RadialGradient,
  Shadow,
  VisualStyle,
} from "@pptxdk/core";
export { chartCapability } from "#/chart/chart.ts";
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
} from "#/chart/chart.ts";
export { Chart } from "#/chart/chart-elements.ts";
export { imageCapability } from "#/image/image.ts";
export type {
  ImageDeclarationProps,
  ImageSource,
  ImageVisualStyle,
} from "#/image/image.ts";
export { Image } from "#/image/image-elements.ts";
export type { ImageProps } from "#/image/image-elements.ts";
export { groupCapability } from "#/group/group.ts";
export type { GroupDeclarationProps } from "#/group/group.ts";
export { Group } from "#/group/group-elements.ts";
export type { GroupProps } from "#/group/group-elements.ts";
export { tableCapability } from "#/table/table.ts";
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
} from "#/table/table.ts";
export { Cell, Row, Table } from "#/table/table-elements.ts";
export type {
  CellProps,
  RowProps,
  TableProps,
} from "#/table/table-elements.ts";
export { font, fromFontCollection, textCapability } from "#/text/text.ts";
export { P, SlideNumber, Span, Text } from "#/text/text-elements.ts";
export { notesCapability } from "#/notes/notes.ts";
export { Notes } from "#/notes/notes-elements.ts";
export type { NotesProps } from "#/notes/notes-elements.ts";
export type {
  NotesDeclarationProps,
  NotesNode,
} from "#/notes/notes-declaration.ts";
export { transitionCapability, transitions } from "#/transition/transition.ts";
export { slideBackgroundCapability } from "#/slide-background/slide-background.ts";
export { createMorphScope } from "#/shared/morph-api.ts";
export type { MorphKey, MorphProps, MorphScope } from "#/shared/morph-api.ts";
export {
  after,
  appear,
  breathe,
  collapse,
  disappear,
  fillColor,
  motionSchedule,
  createMotionScopeHandle,
  fade,
  fadeOut,
  fly,
  flyOut,
  floatDown,
  floatOut,
  floatUp,
  motionCapability,
  onStep,
  repeat,
  spin,
  spinIn,
  spinOut,
  stretch,
  wipe,
  wipeOut,
  zoom,
  zoomOut,
} from "#/motion/motion.ts";
export { MotionScope, Step, useMotionScope } from "#/motion/motion-elements.ts";
export type {
  EffectInstance,
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
  EntrancePreset,
  ExitEffect,
  ExitEffectDeclaration,
  ExitPreset,
  InstantEffectOptions,
  MotionDirection,
  MotionGranularity,
  MotionCompletionState,
  MotionCue,
  MotionCueRef,
  MotionDecision,
  MotionDependency,
  MotionDifference,
  MotionEffectRef,
  MotionIntervalRef,
  MotionMembership,
  MotionMomentRef,
  MotionNativeProjection,
  MotionMoment,
  MotionMomentOffset,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeQueryMember,
  MotionScopeRef,
  MotionScopeSchedule,
  MotionScopeProps,
  MotionProps,
  MotionStartDependency,
  MotionTargetRange,
  MotionTiming,
  MotionStepQuery,
  MotionStepRef,
  MotionTime,
  MotionSchedule,
  RepeatCount,
  RepeatedEffect,
  ResolvedMotionEffect,
  ResolvedMotionMoment,
  ResolvedMotionScope,
  ResolvedMotionStep,
  SpinOptions,
  StepProps,
  SteppedEffect,
} from "#/motion/motion.ts";
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
} from "#/text/text.ts";
export type {
  ParagraphProps,
  SlideNumberProps,
  SpanProps,
  TextProps,
} from "#/text/text-elements.ts";
export type {
  FlipTransition,
  MorphTransition,
  StandardTransition,
  StripsDirection,
  Transition,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
} from "#/transition/transition.ts";
