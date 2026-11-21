import { boxCapability } from "@pptxdk/capabilities/box";
import { motionCapability } from "@pptxdk/capabilities/motion";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import { transitionCapability } from "@pptxdk/capabilities/transition";

import { createStagePipeline } from "./compose.ts";

export const { build } = createStagePipeline([
  boxCapability,
  shapeCapability,
  transitionCapability,
  motionCapability,
]);

export { formatBaseline, targets } from "@pptxdk/core";
export { Deck } from "@pptxdk/react";
export { Box, type BoxProps } from "@pptxdk/capabilities/box-elements";
export { Slide } from "./slide.ts";
export { Shape } from "@pptxdk/capabilities/shape-elements";
export {
  blockArc,
  linearGradient,
  path,
  preset,
  radialGradient,
  shapeCapability,
} from "@pptxdk/capabilities/shape";
export type {
  GradientStop,
  LayoutStyle,
  LinearGradient,
  RadialGradient,
  Shadow,
  VisualStyle,
} from "@pptxdk/core";
export type { ImageVisualStyle } from "@pptxdk/capabilities";
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
export type { ShapeProps } from "@pptxdk/capabilities/shape-elements";
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PathOptions,
  PresetGeometry,
  PresetGeometryName,
} from "@pptxdk/capabilities/shape";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  Diagnostic,
  ObjectObservation,
  SlideObservation,
  Target,
} from "@pptxdk/core";
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
