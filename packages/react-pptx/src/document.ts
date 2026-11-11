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
  deck,
  documentCapabilities,
  font,
  paragraph,
  section,
  slide,
  slideNumber,
  span,
  text,
} from "@react-pptx/capabilities/document";
export type {
  BoxAuthorProps,
  DeckAuthorProps,
  DeckNode,
  BoxNode,
  ParagraphAuthorProps,
  ParagraphContent,
  ParagraphNode,
  SectionAuthorProps,
  SectionNode,
  SlideAuthorProps,
  SlideNode,
  SlideNumberAuthorProps,
  SlideNumberNode,
  SpanAuthorProps,
  SpanContent,
  SpanNode,
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
} from "@react-pptx/core/capability";
