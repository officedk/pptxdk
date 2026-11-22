import { documentCapabilities } from "@pptxdk/capabilities/all";
import { slideBackgroundCapability } from "@pptxdk/capabilities/slide-background";
import { emit } from "@pptxdk/emit";

import { createStagePipeline } from "./compose.ts";

const pipeline = createStagePipeline(documentCapabilities);

/**
 * Synchronously resolves an editable CommonNode document into an immutable Resolution snapshot.
 */
export const resolve = pipeline.resolve;

/**
 * Synchronously emits a PPTX archive and BuildReport from an existing Resolution.
 */
export { emit };

/**
 * Default one-call build supporting both JSX elements and plain CommonNode documents.
 */
export const build = pipeline.build;

/** Default capability list shared across the SDK author entrypoints. */
export { documentCapabilities, slideBackgroundCapability };

export {
  authorIdSchema,
  AuthorQuery,
  AuthorQueryError,
  BuildError,
  BuildQueryError,
  CapabilityError,
  CommonContainer,
  CommonNode,
  darker,
  deckPropsSchema,
  defaultTheme,
  defineTheme,
  formatBaseline,
  isNativeColorSlot,
  layoutLengthSchema,
  layoutSpacingSchema,
  layoutStyleSchema,
  lighter,
  linkTargetSchema,
  nativeColorSlots,
  paddingStyleSchema,
  RichTextContainer,
  sectionPropsSchema,
  slideBackgroundSchema,
  slidePropsSchema,
  targets,
  textStyleSchema,
  themeInputSchema,
  themeSchema,
  themeScopeSchema,
} from "@pptxdk/core";
export { createMorphScope, isMorphKey } from "@pptxdk/core/capability";
export type {
  DeclarationChild,
  DeclarationContainer,
  DeclarationNode,
  DeclarationText,
} from "@pptxdk/core/build";
export type {
  AuthorId,
  AuthorNodeGuard,
  BuildOptions,
  BuildQuery,
  BuildReport,
  BuildResult,
  CapabilityModule,
  ColorValue,
  DeckProps,
  Diagnostic,
  DocumentProps,
  DocumentStyle,
  DocumentUpdate,
  EmitOptions,
  EmitResult,
  Extension,
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  GeometryBox,
  GradientStop,
  LayoutChildReference,
  LayoutLength,
  LayoutObservation,
  LayoutSpacing,
  LayoutStyle,
  LinearGradient,
  LinkTarget,
  Locator,
  LocatorObservation,
  LocatorResourceObservation,
  MorphKey,
  MorphScope,
  NativeColorSlot,
  NativeTableObservation,
  NativeTablePartObservation,
  ObjectObservation,
  PaddingStyle,
  PreviewTransform,
  PreviewWarning,
  RadialGradient,
  Resolution,
  ResolutionReport,
  ResolveOptions,
  ResourceObservation,
  RichTextItem,
  SectionProps,
  Shadow,
  SlideBackground,
  SlideObservation,
  SlideProps,
  SlideQuery,
  StyleUpdate,
  Target,
  Theme,
  ThemeColor,
  ThemeInput,
  ThemeScope,
  ThemeShade,
  VisualStyle,
} from "@pptxdk/core";
export { boxCapability } from "@pptxdk/capabilities/box";
export { chartCapability } from "@pptxdk/capabilities/chart";
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
export { groupCapability } from "@pptxdk/capabilities/group";
export { imageCapability } from "@pptxdk/capabilities/image";
export type { ImageSource, ImageVisualStyle } from "@pptxdk/capabilities/image";
export {
  after,
  breathe,
  createMotionScopeHandle,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  motionCapability,
  observeMotion,
  repeat,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "@pptxdk/capabilities/motion";
export type {
  BreatheEffect,
  BreatheOptions,
  Easing,
  Effect,
  EffectInstanceObservation,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  FillColorOptions,
  MotionCompletionState,
  MotionCueObservation,
  MotionCueRef,
  MotionDecision,
  MotionDependencyObservation,
  MotionDifference,
  MotionDirection,
  MotionEffectObservation,
  MotionEffectRef,
  MotionGranularity,
  MotionIntervalRef,
  MotionMembership,
  MotionMoment,
  MotionMomentObservation,
  MotionMomentOffset,
  MotionMomentRef,
  MotionNativeProjection,
  MotionObservation,
  MotionQuery,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeObservation,
  MotionScopeQueryMember,
  MotionScopeQueryObservation,
  MotionScopeRef,
  MotionStartDependency,
  MotionStepObservation,
  MotionStepQuery,
  MotionStepRef,
  MotionTargetRange,
  MotionTime,
  MotionTiming,
  RepeatCount,
  RepeatedEffect,
  SpinOptions,
} from "@pptxdk/capabilities/motion";
export { notesCapability } from "@pptxdk/capabilities/notes";
export type { NotesDeclarationProps } from "@pptxdk/capabilities/notes-declaration";
export {
  blockArc,
  linearGradient,
  path,
  preset,
  radialGradient,
  shapeCapability,
} from "@pptxdk/capabilities/shape";
export type {
  ArrowType,
  CustomGeometry,
  Geometry,
  GeometryPath,
  LineEndStyle,
  PathCommand,
  PathOptions,
  PresetGeometry,
  PresetGeometryName,
  ShapeVisualStyle,
} from "@pptxdk/capabilities/shape";
export { tableCapability } from "@pptxdk/capabilities/table";
export type {
  TableCellStyle,
  TableColumn,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
} from "@pptxdk/capabilities/table";
export { font, textCapability } from "@pptxdk/capabilities/text";
export type {
  ParagraphStyle,
  ShapeTextStyle,
  TextAlignment,
  TextList,
  TextStyle,
} from "@pptxdk/capabilities/text";
export {
  transitionCapability,
  transitions,
} from "@pptxdk/capabilities/transition";
export type {
  FlipTransition,
  MorphTransition,
  StandardTransition,
  StripsDirection,
  Transition,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
} from "@pptxdk/capabilities/transition";
