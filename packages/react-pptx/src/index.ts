import {
  shapeCapability,
  textCapability,
  transitionCapability,
} from "@react-pptx/capabilities";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([
  shapeCapability,
  textCapability,
  transitionCapability,
]);

export { BuildError, Deck, formatBaseline, targets } from "@react-pptx/core";
export { Slide } from "./slide.ts";
export {
  font,
  P,
  Shape,
  shapeCapability,
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
  Diagnostic,
  ObjectObservation,
  SlideObservation,
  Target,
} from "@react-pptx/core";
export type { SlideProps } from "./slide.ts";
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
export type { ShapeProps } from "@react-pptx/capabilities";
