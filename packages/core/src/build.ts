export { createResolver } from "#/resolve.ts";
export { BuildError, CapabilityError } from "#/diagnostics.ts";
export {
  buildOptionsSchema,
  diagnosticSchema,
  diagnosticSeveritySchema,
  emitOptionsSchema,
  extensionSchema,
  fontEmbeddingModeSchema,
  formatBaseline,
  formatBaselineSchema,
  resolveOptionsSchema,
  targetNameSchema,
  targetSchema,
  targets,
} from "#/stage-options.ts";
export type {
  BuildOptions,
  Diagnostic,
  DiagnosticSeverity,
  EmitOptions,
  Extension,
  FontEmbeddingMode,
  ResolveOptions,
  Target,
  TargetName,
} from "#/stage-options.ts";
export type {
  BuildQuery,
  BuildReport,
  BuildResult,
  CapabilityModule,
  EmitResult,
  Resolution,
  ResolutionReport,
  ResourceObservation,
} from "#/types.ts";
export type {
  DeclarationChild,
  DeclarationContainer,
  DeclarationNode,
  DeclarationText,
} from "#/declaration.ts";
