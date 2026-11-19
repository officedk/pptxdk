export {
  createResolver,
  emit,
  validateEmitOptions,
  validateOptions,
} from "#/formal-build.ts";
export { BuildError, CapabilityError } from "#/diagnostics.ts";
export { formatBaseline, targets } from "#/types.ts";
export type {
  BuildOptions,
  BuildQuery,
  BuildReport,
  BuildResult,
  CapabilityModule,
  Diagnostic,
  EmitOptions,
  EmitResult,
  Extension,
  Resolution,
  ResolutionReport,
  ResolveOptions,
  Target,
} from "#/types.ts";
export type {
  DeclarationChild,
  DeclarationContainer,
  DeclarationNode,
  DeclarationText,
} from "#/declaration.ts";
