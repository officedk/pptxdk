import { documentCapabilities } from "@pptxdk/capabilities/all";
import { slideBackgroundCapability } from "@pptxdk/capabilities/slide-background";
import {
  CodegenError,
  codegenDiagnosticCodeSchema,
  codegenDiagnosticSchema,
  codegenOptionsSchema,
  codegenTargetSchema,
  createCodegen,
  type CodegenDiagnostic,
  type CodegenDiagnosticCode,
  type CodegenEntryName,
  type CodegenFileContent,
  type CodegenOptions,
  type CodegenResult,
  type CodegenTarget,
} from "@pptxdk/codegen";
import type { CommonNode as DocumentCommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import {
  ImportError,
  cascadedOmissionSchema,
  createImporter,
  editabilityLossSchema,
  fatalImportIssueSchema,
  importDiagnosticSchema,
  importLocationSchema,
  importOptionsSchema,
  importReportSchema,
  type CascadedOmission,
  type EditabilityLoss,
  type FatalImportIssue,
  type ImportDiagnostic,
  type ImportLocation,
  type ImportOptions,
  type ImportReport,
  type ImportResult,
} from "@pptxdk/import";
import type { PreviewOptions, PreviewResult } from "@pptxdk/preview";
import type { ReactNode } from "react";

import { createStagePipeline } from "./compose.ts";

const pipeline = createStagePipeline(documentCapabilities);
const importDocument = createImporter(documentCapabilities);
const generateSource = createCodegen(documentCapabilities);

/**
 * Synchronously generates rebuildable vanilla TypeScript or React TSX source files from an
 * editable CommonNode document.
 */
export function codegen<T extends CodegenTarget>(
  document: DocumentCommonNode,
  options: CodegenOptions<T>,
): CodegenResult<T> {
  return generateSource(document, options);
}

export {
  CodegenError,
  codegenDiagnosticCodeSchema,
  codegenDiagnosticSchema,
  codegenOptionsSchema,
  codegenTargetSchema,
};
export type {
  CodegenDiagnostic,
  CodegenDiagnosticCode,
  CodegenEntryName,
  CodegenFileContent,
  CodegenOptions,
  CodegenResult,
  CodegenTarget,
};

/**
 * Synchronously imports a Transitional PPTX archive into an editable CommonNode document
 * and a structured `ImportReport`, detaching completely from caller-owned input bytes.
 */
export function importPptx(
  input: Uint8Array,
  options?: ImportOptions,
): ImportResult {
  return importDocument(input, options);
}

export {
  ImportError,
  cascadedOmissionSchema,
  editabilityLossSchema,
  fatalImportIssueSchema,
  importDiagnosticSchema,
  importLocationSchema,
  importOptionsSchema,
  importReportSchema,
};
export type {
  CascadedOmission,
  EditabilityLoss,
  FatalImportIssue,
  ImportDiagnostic,
  ImportLocation,
  ImportOptions,
  ImportReport,
  ImportResult,
};

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
 * Default static preview supporting both JSX elements and plain CommonNode documents.
 *
 * Defers loading the rasterizer and React reconciler until `preview()` is actually called.
 */
export async function preview(
  input: ReactNode | DocumentCommonNode,
  options: PreviewOptions = {},
): Promise<PreviewResult> {
  const { preview: runPreview } = await import("./preview.ts");
  return runPreview(input, options);
}

export type {
  PreviewNode,
  PreviewOptions,
  PreviewPageResult,
  PreviewResult,
} from "@pptxdk/preview";
export { documentCapabilities, slideBackgroundCapability };

export {
  authorIdSchema,
  backgroundRefSchema,
  backgroundValueSchema,
  borderObjectSchema,
  borderStyleSchema,
  borderValueSchema,
  boxShadowValueSchema,
  externalImageFillSourceSchema,
  footerPlaceholderSpecSchema,
  footerPlaceholdersSchema,
  headerFooterSchema,
  imageBackgroundFitSchema,
  imageBackgroundPositionKeywordSchema,
  imageBackgroundPositionSchema,
  imageBackgroundSchema,
  imageFillSourceSchema,
  placeholderSizeSchema,
  AuthorQuery,
  AuthorQueryError,
  BuildError,
  BuildQueryError,
  PreviewPageRangeError,
  buildDiagnosticFontUnavailableReasonSchema,
  buildDiagnosticSchema,
  buildOptionsSchema,
  CapabilityError,
  colorMapOverrideSchema,
  colorMapSchema,
  CommonContainer,
  CommonNode,
  DeckContainer,
  darker,
  deckPropsSchema,
  defaultColorMap,
  defaultTheme,
  defineTheme,
  emitOptionsSchema,
  extensionSchema,
  fontEmbeddingModeSchema,
  formatBaseline,
  formatBaselineSchema,
  formatSchemeSchema,
  isNativeColorSlot,
  isSchemeColorName,
  isSchemeColorSlot,
  isSlideLayoutDefinition,
  isTheme,
  layoutLengthSchema,
  layoutSpacingSchema,
  layoutStyleSchema,
  lighter,
  linkTargetSchema,
  nativeColorSlots,
  paddingStyleSchema,
  previewErrorSchema,
  previewPlaceholderSchema,
  resolvedMasterSchema,
  resolvedPlaceholderSchema,
  resolvedSlideLayoutSchema,
  resolveOptionsSchema,
  RichTextContainer,
  schemeColorNames,
  schemeColorSlots,
  sectionPropsSchema,
  slideBackgroundSchema,
  slideFooterValueSchema,
  slideNumberValueSchema,
  SlideContainer,
  SlideLayoutContainer,
  slideLayoutDataPropsSchema,
  slideLayoutTypeSchema,
  slideLayoutTypeValues,
  SlideMasterContainer,
  slideMasterDataPropsSchema,
  slidePropsSchema,
  targetNameSchema,
  targetSchema,
  targets,
  textStyleSchema,
  themeDataSchema,
  themeInputSchema,
  themeOverrideSchema,
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
  BackgroundRef,
  BackgroundValue,
  BorderObject,
  BorderStyle,
  BorderValue,
  BoxShadowValue,
  BuildDiagnostic,
  BuildDiagnosticFontUnavailableReason,
  BuildOptions,
  BuildReport,
  BuildResult,
  CapabilityModule,
  ColorMap,
  ColorMapOverride,
  ColorValue,
  ContentPlaceholderRole,
  DeckProps,
  DocumentProps,
  DocumentStyle,
  DocumentUpdate,
  EmitOptions,
  EmitResult,
  EmptyPlaceholderMap,
  Extension,
  ExternalImageFillSource,
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  FontStyleDescriptor,
  FontUsage,
  FontEmbeddingMode,
  FooterPlaceholderSpec,
  FooterPlaceholders,
  FormatScheme,
  GradientStop,
  HeaderFooter,
  ImageBackground,
  ImageBackgroundFit,
  ImageBackgroundPosition,
  ImageFillSource,
  LayoutChildReference,
  LayoutLength,
  LayoutPlaceholderSpec,
  LayoutSpacing,
  LayoutStyle,
  LinearGradient,
  LinkTarget,
  Locator,
  LocatorResource,
  MasterPlaceholderMap,
  MasterPlaceholderSpec,
  MasterTextStyles,
  MorphKey,
  MorphScope,
  NativeColorSlot,
  NativePlaceholderRef,
  ObjectTransform,
  PaddingStyle,
  PlaceholderLevelStyle,
  PlaceholderListStyles,
  PlaceholderMap,
  PlaceholderOrient,
  PlaceholderSize,
  PlaceholderStyle,
  PreviewError,
  PreviewPlaceholder,
  PreviewTransform,
  RadialGradient,
  Rect,
  Resolution,
  ResolutionDimensionsSummary,
  ResolutionReport,
  ResolvedFooter,
  ResolvedFooterSlot,
  ResolvedImage,
  ResolvedLayout,
  ResolvedMaster,
  ResolvedObject,
  ResolvedPlaceholder,
  ResolvedResource,
  ResolvedSlide,
  ResolvedSlideLayout,
  ResolvedTable,
  ResolvedTableCell,
  ResolvedTableRow,
  ResolveOptions,
  RichTextItem,
  SchemeColorName,
  SchemeColorSlot,
  ScreenPoint,
  SectionProps,
  Shadow,
  SlideBackground,
  SlideFooterValue,
  SlideNumberValue,
  SlideLayoutDefinition,
  SlideLayoutProps,
  SlideLayoutType,
  SlideLocator,
  SlideMasterProps,
  SlideProps,
  StyleUpdate,
  Target,
  TargetName,
  TextMetrics,
  TextPlaceholderRole,
  Theme,
  ThemeColor,
  ThemeData,
  ThemeInput,
  ThemeOverride,
  ThemeScope,
  ThemeShade,
  UnmeasuredFontIssue,
  VisualStyle,
} from "@pptxdk/core";
export { defineSlideLayout } from "@pptxdk/capabilities/document";
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
export { footerPlaceholderCapability } from "@pptxdk/capabilities/footer-placeholder";
export { groupCapability } from "@pptxdk/capabilities/group";
export { imageCapability } from "@pptxdk/capabilities/image";
export type { ImageSource, ImageVisualStyle } from "@pptxdk/capabilities/image";
export { mastersCapability } from "@pptxdk/capabilities/masters";
export { placeholderCapability } from "@pptxdk/capabilities/placeholder";
export type {
  PlaceholderDeclarationProps,
  PlaceholderNode,
} from "@pptxdk/capabilities/placeholder";
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
  motionSchedule,
  repeat,
  spin,
  spinIn,
  spinOptionsSchema,
  stretch,
  wipe,
  zoom,
} from "@pptxdk/capabilities/motion";
export type {
  BreatheEffect,
  BreatheOptions,
  Easing,
  Effect,
  EffectInstance,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EmphasisEffectDeclaration,
  EntranceEffect,
  EntranceEffectDeclaration,
  FillColorOptions,
  MotionCompletionState,
  MotionCue,
  MotionCueRef,
  MotionDecision,
  MotionDependency,
  MotionDifference,
  MotionDirection,
  MotionEffectRef,
  MotionGranularity,
  MotionIntervalRef,
  MotionMembership,
  MotionMoment,
  MotionMomentOffset,
  MotionMomentRef,
  MotionNativeProjection,
  MotionSchedule,
  MotionScopeHandle,
  MotionScopeMode,
  MotionScopeQueryMember,
  MotionScopeRef,
  MotionScopeSchedule,
  MotionStartDependency,
  MotionStepQuery,
  MotionStepRef,
  MotionTargetRange,
  MotionTime,
  MotionTiming,
  RepeatCount,
  RepeatedEffect,
  ResolvedMotionEffect,
  ResolvedMotionMoment,
  ResolvedMotionScope,
  ResolvedMotionStep,
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
