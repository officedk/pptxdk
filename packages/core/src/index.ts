export { createBuilder } from "#/builder.ts";
export { BuildError, CapabilityError } from "#/diagnostics.ts";
export { Box, Deck, Slide } from "#/elements.ts";
export { pxToEmu, xmlEscape } from "#/ooxml.ts";
export { formatBaseline, targets } from "#/types.ts";
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  BoxProps,
  CapabilityLayoutMeasurement,
  CapabilityLayoutRequest,
  CapabilityChild,
  CapabilityChildren,
  CapabilityContext,
  CapabilityInput,
  CapabilityModule,
  CapabilityOutput,
  CapabilityResource,
  DeckProps,
  Diagnostic,
  ElementProps,
  Extension,
  GeometryBox,
  LayoutChildReference,
  LayoutLength,
  LayoutObservation,
  LayoutSpacing,
  LayoutStyle,
  ObjectCapabilityModule,
  ObjectObservation,
  SlideCapabilityContext,
  SlideCapabilityModule,
  SlideCapabilityOutput,
  SlideObservation,
  SlideProps,
  Target,
} from "#/types.ts";
