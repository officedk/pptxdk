export { createCodegen } from "./codegen.ts";
export {
  CodegenError,
  codegenTargetSchema,
  defineObjectCodegen,
  defineSlideCodegen,
  expr,
} from "./types.ts";
export type {
  CodegenContext,
  CodegenContribution,
  CodegenDiagnostic,
  CodegenDiagnosticCode,
  CodegenEntryName,
  CodegenExpr,
  CodegenFileContent,
  CodegenOptions,
  CodegenPropSerializer,
  CodegenResult,
  CodegenTarget,
  CodegenValueSerializer,
  DeclarationCodegenHandler,
  ObjectCodegenContribution,
  ProductionDeclarationKind,
  ProductionSlideSlotProp,
  SlideCodegenContribution,
} from "./types.ts";
