import { boxCapability } from "@react-pptx/capabilities/box";
import { motionCapability } from "@react-pptx/capabilities/motion";
import { shapeCapability } from "@react-pptx/capabilities/shape";
import { transitionCapability } from "@react-pptx/capabilities/transition";

import { createStagePipeline } from "./compose.ts";

export const { build } = createStagePipeline([
  boxCapability,
  shapeCapability,
  transitionCapability,
  motionCapability,
]);

export { Deck, formatBaseline, targets } from "@react-pptx/core";
export { Box, type BoxProps } from "@react-pptx/capabilities/box-elements";
export { Slide } from "./slide.ts";
export { Shape } from "@react-pptx/capabilities/shape-elements";
export {
  blockArc,
  linearGradient,
  path,
  preset,
  radialGradient,
  shapeCapability,
} from "@react-pptx/capabilities/shape";
export type {
  GradientStop,
  LayoutStyle,
  LinearGradient,
  RadialGradient,
  Shadow,
  VisualStyle,
} from "@react-pptx/core";
export type { ImageVisualStyle } from "@react-pptx/capabilities";
export { Step } from "@react-pptx/capabilities/motion-elements";
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
  SpinOptions,
  StepProps,
} from "@react-pptx/capabilities/motion";
export type { ShapeProps } from "@react-pptx/capabilities/shape-elements";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
} from "@react-pptx/capabilities/shape";
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
