import { textCapability } from "@react-pptx/capabilities/text";
import { transitionCapability } from "@react-pptx/capabilities/transition";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([textCapability, transitionCapability]);

export { Box, Deck, formatBaseline, targets } from "@react-pptx/core";
export { Slide } from "./slide.ts";
export {
  font,
  P,
  Span,
  Text,
  textCapability,
} from "@react-pptx/capabilities/text";
export { createBuilder } from "@react-pptx/core";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  BoxProps,
  Diagnostic,
  LayoutStyle,
  ObjectObservation,
  SlideObservation,
  Target,
} from "@react-pptx/core";
export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  ParagraphProps,
  ParagraphStyle,
  TextAlignment,
  TextList,
  TextProps,
  TextStyle,
} from "@react-pptx/capabilities/text";
export type { SlideProps } from "./slide.ts";
export {
  transitionCapability,
  transitions,
} from "@react-pptx/capabilities/transition";
export type {
  FlipTransition,
  MorphTransition,
  StandardTransition,
  StripsDirection,
  Transition,
  TransitionDirection,
  TransitionDuration,
  TransitionSpeed,
} from "@react-pptx/capabilities/transition";
