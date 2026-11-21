import { documentCapabilities } from "@pptxdk/capabilities/document";
import {
  createResolver,
  type BuildOptions,
  type BuildResult,
  type Resolution,
  type ResolveOptions,
} from "@pptxdk/core/build";
import type { CommonNode } from "@pptxdk/core/document";
import { emit, emittedBuildQuery, validateOptions } from "@pptxdk/emit";

const resolveDocument = createResolver(documentCapabilities);

/**
 * Synchronously resolves an editable CommonNode document without loading React.
 */
export function resolve(
  document: CommonNode,
  options: ResolveOptions = {},
): Resolution {
  return resolveDocument(document, options);
}

export { emit };

/**
 * Default formal build for plain authoring.
 *
 * Capability modules are registered explicitly here; author factories create declarations only and
 * never enable a capability that the caller chose to trim. Resolution runs synchronously before
 * control returns so caller edits after the call affect only later builds.
 */
export async function build(
  document: CommonNode,
  options: BuildOptions = {},
): Promise<BuildResult> {
  validateOptions(options);
  const resolveOptions: ResolveOptions =
    options.signal === undefined ? {} : { signal: options.signal };
  const resolution = resolve(document, resolveOptions);
  const emitted = emit(resolution, options);
  return Object.freeze({
    pptx: emitted.pptx,
    report: emitted.report,
    query: emittedBuildQuery(emitted.report, resolution.query),
  });
}

export {
  box,
  cell,
  chart,
  createMotionScopeHandle,
  deck,
  documentCapabilities,
  font,
  group,
  image,
  motionScope,
  notes,
  notesCapability,
  paragraph,
  row,
  section,
  shape,
  slide,
  slideNumber,
  span,
  step,
  table,
  text,
} from "@pptxdk/capabilities/document";
export type {
  BoxAuthorProps,
  BoxNode,
  CellAuthorProps,
  CellNode,
  ChartAuthorProps,
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
  DeckAuthorProps,
  DeckNode,
  DoughnutChartGroup,
  IndependentCategoryData,
  IndependentCategorySeries,
  GroupAuthorProps,
  GroupNode,
  ImageAuthorProps,
  ImageNode,
  LineOrAreaChartGroup,
  MotionScopeAuthorProps,
  MotionScopeHandle,
  MotionScopeNode,
  MotionStepAuthorProps,
  MotionStepNode,
  NotesAuthorProps,
  NotesNode,
  ParagraphAuthorProps,
  ParagraphContent,
  ParagraphNode,
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
  RowAuthorProps,
  RowNode,
  SectionAuthorProps,
  SectionNode,
  ShapeAuthorProps,
  ShapeNode,
  SlideAuthorProps,
  SlideNode,
  SlideNumberAuthorProps,
  SlideNumberNode,
  SpanAuthorProps,
  SpanContent,
  SpanNode,
  TableAuthorProps,
  TableCellStyle,
  TableColumn,
  TableNode,
  TableRowStyle,
  TableSpacing,
  TableStyle,
  TableTextStyle,
  TableVerticalAlign,
  TextAuthorProps,
  TextContent,
  TextNode,
  Transition,
} from "@pptxdk/capabilities/document";
export {
  after,
  breathe,
  fillColor,
  fade,
  fly,
  floatDown,
  floatUp,
  observeMotion,
  repeat,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "@pptxdk/capabilities/motion";
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
  EntranceEffect,
  EntranceEffectDeclaration,
  MotionCompletionState,
  MotionCueObservation,
  MotionCueRef,
  MotionDecision,
  MotionDependencyObservation,
  MotionDifference,
  MotionDirection,
  MotionEffectRef,
  MotionEffectObservation,
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
export {
  BuildError,
  CapabilityError,
  formatBaseline,
  targets,
} from "@pptxdk/core/build";
export { AuthorQuery, AuthorQueryError } from "@pptxdk/core/document";
export {
  blockArc,
  linearGradient,
  path,
  preset,
  radialGradient,
} from "@pptxdk/capabilities/shape";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
  PresetGeometryName,
} from "@pptxdk/capabilities/shape";
export type { ImageSource, ImageVisualStyle } from "@pptxdk/capabilities/image";
export { transitions } from "@pptxdk/capabilities/transition";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  CapabilityModule,
  DeclarationChild,
  DeclarationContainer,
  DeclarationNode,
  DeclarationText,
  Diagnostic,
  EmitOptions,
  EmitResult,
  Resolution,
  ResolutionReport,
  ResolveOptions,
  Target,
} from "@pptxdk/core/build";
export type {
  AuthorNodeGuard,
  CommonContainer,
  CommonNode,
  DocumentProps,
  DocumentStyle,
  DocumentUpdate,
  RichTextContainer,
  RichTextItem,
  StyleUpdate,
} from "@pptxdk/core/document";
export { createMorphScope, isMorphKey } from "@pptxdk/core/capability";
export type { MorphKey, MorphScope } from "@pptxdk/core/capability";
export {
  darker,
  defaultTheme,
  defineTheme,
  isNativeColorSlot,
  lighter,
  nativeColorSlots,
} from "@pptxdk/core/theme";
export type {
  ColorValue,
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  GradientStop,
  LinearGradient,
  NativeColorSlot,
  RadialGradient,
  Shadow,
  TextAlignment,
  TextStyle,
  Theme,
  ThemeColor,
  ThemeInput,
  ThemeScope,
  ThemeShade,
  VisualStyle,
} from "@pptxdk/core/theme";
export type {
  LayoutLength,
  LayoutSpacing,
  LayoutStyle,
  LinkTarget,
  SlideBackground,
} from "@pptxdk/core/capability";
