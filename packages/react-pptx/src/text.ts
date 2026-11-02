import { motionCapability } from "@react-pptx/capabilities/motion";
import { textCapability } from "@react-pptx/capabilities/text";
import { transitionCapability } from "@react-pptx/capabilities/transition";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([
  textCapability,
  transitionCapability,
  motionCapability,
]);

export { Deck, formatBaseline, targets } from "@react-pptx/core";
export { Slide } from "./slide.ts";
export {
  font,
  P,
  Span,
  Text,
  textCapability,
} from "@react-pptx/capabilities/text";
export { createBuilder } from "@react-pptx/core";
export {
  Step,
  fade,
  fly,
  floatUp,
  motionCapability,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "@react-pptx/capabilities/motion";
export type {
  Easing,
  Effect,
  EffectOptions,
  EffectPreset,
  EmphasisEffect,
  EntranceEffect,
  MotionDirection,
  MotionEffectObservation,
  MotionObservation,
  MotionProps,
  MotionStepObservation,
  StepProps,
} from "@react-pptx/capabilities/motion";
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
