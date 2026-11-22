export { emit, emittedBuildQuery, validateOptions } from "#/emit.ts";
export {
  assertXml10Characters,
  pptxParts,
  pxToEmu,
  resolvedColorXml,
  themeXml,
  xmlDocument,
  xmlEscape,
} from "#/ooxml.ts";
export type {
  PptxPackagePart,
  PptxPackageRelationshipSet,
  PptxPackageResource,
  PptxRelationship,
} from "#/ooxml.ts";
export {
  buildOptionsSchema,
  diagnosticSchema,
  diagnosticSeveritySchema,
  emitOptionsSchema,
  extensionSchema,
  fontEmbeddingModeSchema,
  formatBaseline,
  formatBaselineSchema,
  targetNameSchema,
  targetSchema,
  targets,
} from "@pptxdk/core/build";
export type {
  BuildOptions,
  BuildQuery,
  BuildReport,
  BuildResult,
  Diagnostic,
  DiagnosticSeverity,
  EmitOptions,
  EmitResult,
  Extension,
  FontEmbeddingMode,
  Resolution,
  Target,
  TargetName,
} from "@pptxdk/core/build";
