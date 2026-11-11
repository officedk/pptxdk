import { documentCapabilities } from "@react-pptx/capabilities/document";
import { createDocumentBuild } from "@react-pptx/core/document";

/**
 * Default formal build for plain authoring.
 *
 * Capability modules are registered explicitly here; author factories create declarations only and
 * never enable a capability that the caller chose to trim. The build consumes the declaration before
 * returning control and shares the same formal implementation as the React entry.
 */
export const build = createDocumentBuild(documentCapabilities);

export {
  box,
  cell,
  chart,
  deck,
  documentCapabilities,
  font,
  group,
  image,
  paragraph,
  row,
  section,
  shape,
  slide,
  slideNumber,
  span,
  table,
  text,
} from "@react-pptx/capabilities/document";
export type {
  BoxAuthorProps,
  BoxNode,
  CellAuthorProps,
  CellNode,
  ChartAuthorProps,
  ChartAxis,
  ChartData,
  ChartDataLabels,
  ChartKind,
  ChartLine,
  ChartNode,
  ChartPoint,
  ChartProps,
  ChartSeries,
  ColumnOrBarChartProps,
  DeckAuthorProps,
  DeckNode,
  DoughnutChartProps,
  GroupAuthorProps,
  GroupNode,
  ImageAuthorProps,
  ImageNode,
  LineOrAreaChartProps,
  ParagraphAuthorProps,
  ParagraphContent,
  ParagraphNode,
  PieChartProps,
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
} from "@react-pptx/capabilities/document";
export {
  BuildError,
  createFormalBuild,
  formatBaseline,
  targets,
} from "@react-pptx/core/build";
export { createDocumentBuild } from "@react-pptx/core/document";
export {
  blockArc,
  linearGradient,
  path,
  preset,
  radialGradient,
} from "@react-pptx/capabilities/shape";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
} from "@react-pptx/capabilities/shape";
export type {
  ImageSource,
  ImageVisualStyle,
} from "@react-pptx/capabilities/image";
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
  Target,
} from "@react-pptx/core/build";
export type {
  CommonContainer,
  CommonNode,
  DocumentProps,
  DocumentStyle,
  DocumentUpdate,
  RichTextContainer,
  RichTextItem,
  StyleUpdate,
} from "@react-pptx/core/document";
export { createMorphScope, isMorphKey } from "@react-pptx/core/capability";
export type { MorphKey, MorphScope } from "@react-pptx/core/capability";
export {
  darker,
  defaultTheme,
  defineTheme,
  isNativeColorSlot,
  lighter,
  nativeColorSlots,
} from "@react-pptx/core/theme";
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
} from "@react-pptx/core/theme";
export type {
  LayoutLength,
  LayoutSpacing,
  LayoutStyle,
  LinkTarget,
  SlideBackground,
} from "@react-pptx/core/capability";
