import { shapeCapability, textCapability } from "@react-pptx/capabilities";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([shapeCapability, textCapability]);

export {
  BuildError,
  Deck,
  formatBaseline,
  Slide,
  targets,
} from "@react-pptx/core";
export {
  font,
  P,
  Shape,
  shapeCapability,
  Span,
  Text,
  textCapability,
} from "@react-pptx/capabilities";
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
  Font,
  FontFace,
  FontFaceStyle,
  FontReference,
  FontSize,
  ParagraphProps,
  ParagraphStyle,
  ShapeTextStyle,
  SpanProps,
  TextAlignment,
  TextList,
  TextProps,
  TextStyle,
} from "@react-pptx/capabilities";
export type { ShapeProps } from "@react-pptx/capabilities";
