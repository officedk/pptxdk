import { shapeCapability } from "@react-pptx/capabilities/shape";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([shapeCapability]);

export { Deck, formatBaseline, Slide, targets } from "@react-pptx/core";
export { Shape, shapeCapability } from "@react-pptx/capabilities/shape";
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
