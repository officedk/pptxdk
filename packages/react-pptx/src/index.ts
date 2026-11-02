import { shapeCapability } from "@react-pptx/capabilities";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([shapeCapability]);

export {
  BuildError,
  Deck,
  formatBaseline,
  Slide,
  targets,
} from "@react-pptx/core";
export { Shape, shapeCapability } from "@react-pptx/capabilities";
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
export type { ShapeProps } from "@react-pptx/capabilities";
