import type { DeclarationChild } from "@pptxdk/core/build";
import type { RegisteredCapabilityModule } from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";
import { z } from "zod";

/** Supported author source generation syntax targets. */
export const codegenTargetSchema = z.enum(["vanilla", "react"]);
export type CodegenTarget = z.infer<typeof codegenTargetSchema>;

export type CodegenEntryName<T extends CodegenTarget = CodegenTarget> =
  T extends "vanilla"
    ? "deck.ts"
    : T extends "react"
      ? "deck.tsx"
      : "deck.ts" | "deck.tsx";

export type CodegenOptions<T extends CodegenTarget = CodegenTarget> = {
  readonly target: T;
};

export type CodegenFileContent = string | Uint8Array;

export type CodegenResult<T extends CodegenTarget = CodegenTarget> = {
  readonly target: T;
  readonly entry: CodegenEntryName<T>;
  readonly files: Readonly<Record<string, CodegenFileContent>>;
};

export type CodegenDiagnosticCode =
  | "invalid-document"
  | "invalid-options"
  | "unsupported-declaration"
  | "unsupported-slide-property"
  | "unexpressible-value";

export type CodegenDiagnostic = {
  readonly severity: "error";
  readonly path: string;
  readonly code: CodegenDiagnosticCode;
  readonly message: string;
};

/** Structured error thrown when an editable document cannot be expressed in generated source. */
export class CodegenError extends Error {
  readonly diagnostics: readonly CodegenDiagnostic[];

  constructor(diagnostics: readonly CodegenDiagnostic[]) {
    const first = diagnostics[0];
    super(
      first === undefined
        ? "Codegen failed"
        : `${first.path}: ${first.message} (${first.code})`,
    );
    this.name = "CodegenError";
    this.diagnostics = Object.freeze([...diagnostics]);
  }
}

const codegenExprBrand: unique symbol = Symbol("CodegenExpr");

export type CodegenExpr = {
  readonly [codegenExprBrand]: true;
  readonly code: string;
};

/** Marks a raw TypeScript/JSX expression string produced by a capability serializer. */
export function expr(code: string): CodegenExpr {
  return Object.freeze({
    [codegenExprBrand]: true as const,
    code,
  });
}

export function isCodegenExpr(value: unknown): value is CodegenExpr {
  return (
    typeof value === "object" &&
    value !== null &&
    codegenExprBrand in value &&
    typeof (value as { readonly code?: unknown }).code === "string"
  );
}

/**
 * Capability-facing context exposed to `ObjectCodegenContribution` and `SlideCodegenContribution`
 * handlers while emitting a single deck file.
 */
export type CodegenContext = {
  readonly target: CodegenTarget;
  readonly path: string;
  useImport(symbol: string, from?: "sdk" | "node:fs"): string;
  registerImageAsset(bytes: Uint8Array, ext: string): string;
  registerFontAsset(
    bytes: Uint8Array,
    family: string,
    ext?: "ttf" | "otf",
  ): string;
  allocateMotionScope(identity: string): string;
  emitValueExpr(value: unknown, propPath?: string): string;
  emitChildNode(node: CommonNode, index: number): string;
  emitRichTextChildren(
    children: readonly DeclarationChild[],
  ): readonly string[];
  emitObjectExtraProps(
    props: Readonly<Record<string, unknown>>,
  ): Readonly<Record<string, CodegenExpr>>;
  formatProps(entries: readonly (readonly [string, unknown])[]): string;
  formatJsxElement(
    tag: string,
    entries: readonly (readonly [string, unknown])[],
    children?: readonly string[],
  ): string;
  formatVanillaCall(
    callee: string,
    args: readonly string[],
    children?: readonly string[],
  ): string;
  fail(code: CodegenDiagnosticCode, message: string, subPath?: string): never;
};

export type DeclarationCodegenHandler = {
  readonly emitVanilla: (node: CommonNode, ctx: CodegenContext) => string;
  readonly emitReact: (node: CommonNode, ctx: CodegenContext) => string;
};

export type CodegenPropSerializer = (
  value: unknown,
  ctx: CodegenContext,
) => CodegenExpr;

export type CodegenValueSerializer = (
  value: unknown,
  ctx: CodegenContext,
) => CodegenExpr | undefined;

export type ObjectCodegenContribution = {
  readonly declarations: Readonly<Record<string, DeclarationCodegenHandler>>;
  readonly objectProps?: Readonly<Record<string, CodegenPropSerializer>>;
  readonly ownerProps?: Readonly<Record<string, CodegenPropSerializer>>;
  readonly valueSerializers?: readonly CodegenValueSerializer[];
};

export type SlideCodegenContribution = {
  readonly slideProps: Readonly<Record<string, CodegenPropSerializer>>;
  readonly valueSerializers?: readonly CodegenValueSerializer[];
};

export type CodegenContribution =
  | ObjectCodegenContribution
  | SlideCodegenContribution;

/** Typed helper for authoring object capability `codegen` contributions. */
export function defineObjectCodegen(
  contribution: ObjectCodegenContribution,
): ObjectCodegenContribution {
  return Object.freeze(contribution);
}

/** Typed helper for authoring slide capability `codegen` contributions. */
export function defineSlideCodegen(
  contribution: SlideCodegenContribution,
): SlideCodegenContribution {
  return Object.freeze(contribution);
}

/** Complete set of production declaration kinds owned by capabilities. */
export type ProductionDeclarationKind =
  | "placeholder"
  | "text"
  | "paragraph"
  | "span"
  | "slide-number"
  | "notes"
  | "image"
  | "shape"
  | "box"
  | "group"
  | "table"
  | "row"
  | "cell"
  | "chart"
  | "motion-step"
  | "motion-scope";

/** Complete set of production slide/master/layout slot properties owned by capabilities. */
export type ProductionSlideSlotProp =
  | "background"
  | "headerFooter"
  | "footerPlaceholders"
  | "transition"
  | "motion";

export type CapabilityCodegenModule = RegisteredCapabilityModule;
