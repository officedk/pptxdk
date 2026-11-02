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
  CapabilityAnnotation,
  CapabilityContribution,
  CapabilityContext,
  CapabilityInput,
  CapabilityPackagePart,
  CapabilityPresentationContribution,
  CapabilityModule,
  CapabilityOutput,
  CapabilityResource,
  DeckProps,
  Diagnostic,
  ElementProps,
  Extension,
  ObjectCapabilityModule,
  ObjectObservation,
  NativeTableObservation,
  SlideCapabilityContext,
  SlideCapabilityModule,
  SlideCapabilityOutput,
  SlideCapabilitySlot,
  SlideObservation,
  SlideProps,
  Target,
} from "#/types.ts";
