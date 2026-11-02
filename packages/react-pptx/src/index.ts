import {
  groupCapability,
  imageCapability,
  shapeCapability,
  textCapability,
  transitionCapability,
} from "@react-pptx/capabilities";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([
  groupCapability,
  imageCapability,
  shapeCapability,
  textCapability,
  transitionCapability,
]);

export {
  Box,
  BuildError,
  Deck,
  formatBaseline,
  targets,
} from "@react-pptx/core";
export { Slide } from "./slide.ts";
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
  transitionCapability,
  transitions,
} from "@react-pptx/capabilities";
export { createBuilder } from "@react-pptx/core";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  BoxProps,
  Diagnostic,
  GeometryBox,
  LayoutChildReference,
  LayoutLength,
  LayoutObservation,
  LayoutSpacing,
  LayoutStyle,
  ObjectObservation,
  SlideObservation,
  Target,
} from "@react-pptx/core";
export type { SlideProps } from "./slide.ts";
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
