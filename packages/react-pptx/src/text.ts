import { textCapability } from "@react-pptx/capabilities/text";
import { createBuilder } from "@react-pptx/core";

export const build = createBuilder([textCapability]);

export { Deck, formatBaseline, Slide, targets } from "@react-pptx/core";
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
