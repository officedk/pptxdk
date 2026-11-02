import {
  groupCapability,
  imageCapability,
  shapeCapability,
  textCapability,
} from "@react-pptx/capabilities";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([
  shapeCapability,
  textCapability,
  imageCapability,
  groupCapability,
]);

export {
  BuildError,
  Deck,
  formatBaseline,
  Slide,
  targets,
} from "@react-pptx/core";
export {
  Image,
  Group,
  Shape,
  groupCapability,
  imageCapability,
  shapeCapability,
  font,
  P,
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
  GroupProps,
  ImageProps,
  ImageSource,
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  ParagraphProps,
  ParagraphStyle,
  ShapeProps,
  ShapeTextStyle,
  SpanProps,
  TextAlignment,
  TextList,
  TextProps,
  TextStyle,
} from "@react-pptx/capabilities";
