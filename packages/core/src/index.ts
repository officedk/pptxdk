export { createBuilder } from "#/builder.ts";
export { BuildError, CapabilityError } from "#/diagnostics.ts";
export { Deck, Slide } from "#/elements.ts";
export { pxToEmu, xmlEscape } from "#/ooxml.ts";
export { formatBaseline, targets } from "#/types.ts";
export type {
  Box,
  BuildOptions,
  BuildReport,
  BuildResult,
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
  ObjectObservation,
  SlideObservation,
  SlideProps,
  Target,
} from "#/types.ts";
