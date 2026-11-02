import {
  groupCapability,
  imageCapability,
  shapeCapability,
} from "@react-pptx/capabilities";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([
  shapeCapability,
  imageCapability,
  groupCapability,
]);

export {
  BuildError,
  Deck,
  formatBaseline,
  Slide,
  targets,
} from "@react-pptx/core";
export {
  Image,
  Group,
  Shape,
  groupCapability,
  imageCapability,
  shapeCapability,
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
export type {
  GroupProps,
  ImageProps,
  ImageSource,
  ShapeProps,
} from "@react-pptx/capabilities";
