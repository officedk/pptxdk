import { motionCapability } from "@pptxdk/capabilities/motion";
import { notesCapability } from "@pptxdk/capabilities/notes";
import { textCapability } from "@pptxdk/capabilities/text";
import { transitionCapability } from "@pptxdk/capabilities/transition";

import { createStagePipeline } from "./compose.ts";

export const { build } = createStagePipeline([
  textCapability,
  notesCapability,
  transitionCapability,
  motionCapability,
]);

export { formatBaseline, targets } from "@pptxdk/core";
export { Deck } from "@pptxdk/react";
export { Box, type BoxProps } from "@pptxdk/capabilities/box-elements";
export { Slide } from "./slide.ts";
export { font, textCapability } from "@pptxdk/capabilities/text";
export { P, Span, Text } from "@pptxdk/capabilities/text-elements";
export { Notes } from "@pptxdk/capabilities/notes-elements";
export { notesCapability } from "@pptxdk/capabilities/notes";
export { Step } from "@pptxdk/capabilities/motion-elements";
export {
  fade,
  fly,
  floatDown,
  floatUp,
  motionCapability,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "@pptxdk/capabilities/motion";
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
  SpinOptions,
  StepProps,
} from "@pptxdk/capabilities/motion";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  Diagnostic,
  LayoutStyle,
  ObjectObservation,
  SlideObservation,
  Target,
} from "@pptxdk/core";
export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  ParagraphStyle,
  TextAlignment,
  TextList,
  TextStyle,
} from "@pptxdk/capabilities/text";
export type {
  ParagraphProps,
  TextProps,
} from "@pptxdk/capabilities/text-elements";
export type {
  NotesDeclarationProps,
  NotesNode,
} from "@pptxdk/capabilities/notes-declaration";
export type { NotesProps } from "@pptxdk/capabilities/notes-elements";
export type { SlideProps } from "./slide.ts";
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
