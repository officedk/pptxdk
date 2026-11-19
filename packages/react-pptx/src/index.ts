import { documentCapabilities } from "@react-pptx/capabilities/document";
import {
  CommonNode,
  expandDocument,
  type PreviewOptions,
  type PreviewResult,
} from "@react-pptx/core";
import { emit } from "@react-pptx/emit";
import type { ReactNode } from "react";

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

/**
 * Default one-call static preview supporting both JSX elements and plain CommonNode documents.
 * Lazily loads `@react-pptx/core/preview` only when called so normal build paths never load `@resvg/resvg-js`.
 */
export async function preview(
  input: ReactNode | CommonNode,
  options: PreviewOptions = {},
): Promise<PreviewResult> {
  if (options.pages?.some((page) => !Number.isInteger(page) || page < 1)) {
    throw new RangeError("Preview page numbers must be positive integers");
  }
  if (options.signal?.aborted) throw options.signal.reason;
  const signalOption =
    options.signal === undefined ? {} : { signal: options.signal };
  if (input instanceof CommonNode) {
    const resolution = resolve(input, signalOption);
    const { paint } = await import("@react-pptx/core/preview");
    const { pages } = paint(resolution, options);
    return Object.freeze({
      report: resolution.report,
      query: resolution.query,
      pages,
    });
  }
  const [{ paint }, document] = await Promise.all([
    import("@react-pptx/core/preview"),
    expandDocument(input, signalOption),
  ]);
  const resolution = resolve(document, signalOption);
  const { pages } = paint(resolution, options);
  return Object.freeze({
    report: resolution.report,
    query: resolution.query,
    pages,
  });
}

/** Default capability list shared by the React one-call entry and the plain document entry. */
export { documentCapabilities };

export {
  AuthorQuery,
  AuthorQueryError,
  BuildError,
  BuildQueryError,
  Deck,
  Section,
  darker,
  defaultTheme,
  defineTheme,
  formatBaseline,
  isNativeColorSlot,
  lighter,
  nativeColorSlots,
  targets,
} from "@react-pptx/core";
export { Slide } from "./slide.ts";
export {
  after,
  Box,
  Chart,
  chartCapability,
  Image,
  Group,
  createMorphScope,
  createMotionScopeHandle,
  Shape,
  blockArc,
  linearGradient,
  path,
  preset,
  radialGradient,
  Table,
  tableCapability,
  Row,
  Cell,
  groupCapability,
  imageCapability,
  motionCapability,
  shapeCapability,
  slideBackgroundCapability,
  boxCapability,
  font,
  P,
  Span,
  SlideNumber,
  Text,
  Notes,
  notesCapability,
  textCapability,
  transitionCapability,
  transitions,
  Step,
  MotionScope,
  observeMotion,
  breathe,
  fillColor,
  fade,
  fly,
  floatDown,
  floatUp,
  repeat,
  spin,
  spinIn,
  stretch,
  useMotionScope,
  wipe,
  zoom,
} from "@react-pptx/capabilities";
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
  BoxProps,
  ArrowType,
  LineEndStyle,
  ShapeVisualStyle,
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
} from "@react-pptx/capabilities";
export { appendJSX, expandDocument } from "@react-pptx/core";
export {
  CommonContainer,
  CommonNode,
  RichTextContainer,
} from "@react-pptx/core";
export type {
  AppendJSXOptions,
  DocumentProps,
  DocumentStyle,
  DocumentUpdate,
  ExpandDocumentOptions,
  ExpandedDocument,
  RichTextItem,
  StyleUpdate,
} from "@react-pptx/core";
export type {
  AuthorNodeGuard,
  BuildQuery,
  BuildOptions,
  BuildReport,
  BuildResult,
  ColorValue,
  DeckProps,
  Diagnostic,
  EmitOptions,
  EmitResult,
  GeometryBox,
  LayoutChildReference,
  Locator,
  LocatorObservation,
  LayoutLength,
  LayoutObservation,
  LayoutSpacing,
  LayoutStyle,
  LocatorResourceObservation,
  GradientStop,
  LinearGradient,
  RadialGradient,
  Shadow,
  VisualStyle,
  NativeTableObservation,
  NativeTablePartObservation,
  ObjectObservation,
  PaintOptions,
  PaintResult,
  PreviewDiagnostic,
  PreviewNodeObservation,
  PreviewOptions,
  PreviewPageResult,
  PreviewResult,
  PreviewSeverity,
  PreviewTransform,
  PreviewWarning,
  Resolution,
  ResolutionReport,
  ResolveOptions,
  ResourceObservation,
  ScreenGeometry,
  ScreenPoint,
  SlideObservation,
  SlideQuery,
  SlideBackground,
  LinkTarget,
  SectionProps,
  Target,
  Extension,
  NativeColorSlot,
  Theme,
  ThemeColor,
  ThemeInput,
  ThemeScope,
  ThemeShade,
} from "@react-pptx/core";
export type { SlideProps } from "./slide.ts";
export type {
  GroupProps,
  NotesDeclarationProps,
  NotesNode,
  NotesProps,
  MorphKey,
  MorphProps,
  MorphScope,
  ImageProps,
  ImageSource,
  ImageVisualStyle,
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  ParagraphProps,
  ParagraphStyle,
  SlideNumberProps,
  ShapeProps,
  ShapeTextStyle,
  SpanProps,
  TextAlignment,
  TextList,
  TextProps,
  TextStyle,
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
  EffectInstanceObservation,
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
} from "@react-pptx/capabilities";
export type {
  FlipTransition,
  MorphTransition,
  StandardTransition,
  StripsDirection,
  Transition,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
} from "@react-pptx/capabilities";
