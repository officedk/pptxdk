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
export type {
  BuildOptions,
  BuildQuery,
  BuildReport,
  BuildResult,
  Diagnostic,
  EmitOptions,
  EmitResult,
  Extension,
  Resolution,
  Target,
} from "@react-pptx/core/build";
