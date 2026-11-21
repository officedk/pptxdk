export { appendJSX, Deck, expandDocument, Section } from "@pptxdk/react";
export type {
  AppendJSXOptions,
  DeckProps,
  ElementProps,
  ExpandDocumentOptions,
  ExpandedDocument,
  SectionProps,
} from "@pptxdk/react";
export { Slide, type SlideProps } from "./slide.ts";
export { Box, type BoxProps } from "@pptxdk/capabilities/box-elements";
export { Chart, type ChartProps } from "@pptxdk/capabilities/chart-elements";
export { Group, type GroupProps } from "@pptxdk/capabilities/group-elements";
export { Image, type ImageProps } from "@pptxdk/capabilities/image-elements";
export {
  MotionScope,
  Step,
  useMotionScope,
  type MotionProps,
  type MotionScopeProps,
  type StepProps,
} from "@pptxdk/capabilities/motion-elements";
export { Notes, type NotesProps } from "@pptxdk/capabilities/notes-elements";
export { Shape, type ShapeProps } from "@pptxdk/capabilities/shape-elements";
export {
  Cell,
  Row,
  Table,
  type CellProps,
  type RowProps,
  type TableProps,
} from "@pptxdk/capabilities/table-elements";
export {
  P,
  SlideNumber,
  Span,
  Text,
  type ParagraphProps,
  type SlideNumberProps,
  type SpanProps,
  type TextProps,
} from "@pptxdk/capabilities/text-elements";
export {
  font,
  transitions,
  type FlipTransition,
  type MorphProps,
  type MorphTransition,
  type StandardTransition,
  type StripsDirection,
  type Transition,
  type TransitionDirection,
  type TransitionDuration,
  type TransitionSpeed,
} from "@pptxdk/capabilities/document";
