import { shapeCapability } from "@react-pptx/capabilities/shape";
import { transitionCapability } from "@react-pptx/capabilities/transition";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([shapeCapability, transitionCapability]);

export { Box, Deck, formatBaseline, targets } from "@react-pptx/core";
export { Slide } from "./slide.ts";
export { Shape, shapeCapability } from "@react-pptx/capabilities/shape";
export type { BoxProps, LayoutStyle } from "@react-pptx/core";
export type { ShapeProps } from "@react-pptx/capabilities/shape";
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
