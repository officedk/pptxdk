import { z } from "zod";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export const formatBaseline = "ECMA-376 5th edition Transitional" as const;

export const formatBaselineSchema: z.ZodType<typeof formatBaseline> =
  z.literal(formatBaseline);

export type Extension = "p14" | "p159" | "c15";
const rawExtensionSchema = z.enum(["p14", "p159", "c15"]);
export const extensionSchema: z.ZodType<Extension> =
  rawExtensionSchema satisfies z.ZodType<
    AssertExact<Extension, z.infer<typeof rawExtensionSchema>> extends true
      ? Extension
      : never
  >;

export type FontEmbeddingMode = "full" | "subset";
const rawFontEmbeddingModeSchema = z.enum(["full", "subset"]);
export const fontEmbeddingModeSchema: z.ZodType<FontEmbeddingMode> =
  rawFontEmbeddingModeSchema satisfies z.ZodType<
    AssertExact<
      FontEmbeddingMode,
      z.infer<typeof rawFontEmbeddingModeSchema>
    > extends true
      ? FontEmbeddingMode
      : never
  >;

export const targets = Object.freeze({
  transitional: Object.freeze({
    name: "transitional" as const,
    format: formatBaseline,
    extensions: Object.freeze(["p14", "p159", "c15"] as const),
  }),
  baselineOnly: Object.freeze({
    name: "baseline-only" as const,
    format: formatBaseline,
    extensions: Object.freeze([] as const),
  }),
});

export type TargetName = "transitional" | "baseline-only";
const targetNames = [
  targets.transitional.name,
  targets.baselineOnly.name,
] as const;

const rawTargetNameSchema = z.enum(targetNames);
export const targetNameSchema: z.ZodType<TargetName> =
  rawTargetNameSchema satisfies z.ZodType<
    AssertExact<TargetName, z.infer<typeof rawTargetNameSchema>> extends true
      ? TargetName
      : never
  >;

export type Target = {
  readonly name: string;
  readonly format: typeof formatBaseline;
  readonly extensions: readonly Extension[];
};

const rawTargetSchema = z
  .strictObject({
    name: z.string().refine((value) => value.trim().length > 0, {
      message: "Target requires a non-empty name",
    }),
    format: z.literal(formatBaseline),
    extensions: z
      .array(rawExtensionSchema)
      .superRefine((values, ctx) => {
        const seen = new Set<Extension>();
        for (const [index, extension] of values.entries()) {
          if (seen.has(extension)) {
            ctx.addIssue({
              code: "custom",
              path: [index],
              message: `Duplicate target extension declaration ${extension}`,
              params: { kind: "duplicate-extension", extension },
            });
          }
          seen.add(extension);
        }
      })
      .readonly(),
  })
  .readonly();
export const targetSchema: z.ZodType<Target> =
  rawTargetSchema satisfies z.ZodType<
    AssertExact<Target, z.infer<typeof rawTargetSchema>> extends true
      ? Target
      : never
  >;

export type ResolveOptions = {
  readonly signal?: AbortSignal | undefined;
};

const resolveOptionsObjectSchema = z.strictObject({
  signal: z.instanceof(AbortSignal).optional(),
});
export const resolveOptionsShape = resolveOptionsObjectSchema.shape;
const rawResolveOptionsSchema = resolveOptionsObjectSchema.readonly();
export const resolveOptionsSchema: z.ZodType<ResolveOptions> =
  rawResolveOptionsSchema satisfies z.ZodType<
    AssertExact<
      ResolveOptions,
      z.infer<typeof rawResolveOptionsSchema>
    > extends true
      ? ResolveOptions
      : never
  >;

export type BuildOptions = {
  /** Explicitly request embedding full files or character subsets of fonts actually used by the deck; no embedding by default. */
  readonly embedFonts?: FontEmbeddingMode | undefined;
  readonly target?: Target | undefined;
  /**
   * Cooperative cancellation signal for a single build; no default timeout when unspecified.
   * Completion, unhandled errors, and cancellation follow the first established terminal state; cancelling after completion does not alter the result.
   */
  readonly signal?: AbortSignal | undefined;
};
export type EmitOptions = BuildOptions;

const buildOptionsObjectSchema = z.strictObject({
  /** Explicitly request embedding full files or character subsets of fonts actually used by the deck; no embedding by default. */
  embedFonts: rawFontEmbeddingModeSchema.optional(),
  target: rawTargetSchema.optional(),
  /**
   * Cooperative cancellation signal for a single build; no default timeout when unspecified.
   * Completion, unhandled errors, and cancellation follow the first established terminal state; cancelling after completion does not alter the result.
   */
  signal: z.instanceof(AbortSignal).optional(),
});
export const buildOptionsShape = buildOptionsObjectSchema.shape;
const rawBuildOptionsSchema = buildOptionsObjectSchema.readonly();
export const buildOptionsSchema: z.ZodType<BuildOptions> =
  rawBuildOptionsSchema satisfies z.ZodType<
    AssertExact<
      BuildOptions,
      z.infer<typeof rawBuildOptionsSchema>
    > extends true
      ? BuildOptions
      : never
  >;

export const emitOptionsSchema: z.ZodType<EmitOptions> = buildOptionsSchema;

const diagnosticFontStyleSchema = z.enum(["normal", "italic"]);

export type BuildDiagnosticFontUnavailableReason =
  | "missing-resource"
  | "missing-face"
  | "family-mismatch"
  | "style-mismatch"
  | "substitute-required"
  | "unsupported-format"
  | "invalid-metrics"
  | "not-embeddable"
  | "license-restricted"
  | "subsetting-forbidden"
  | "subsetting-failed";

const rawBuildDiagnosticFontUnavailableReasonSchema = z.enum([
  "missing-resource",
  "missing-face",
  "family-mismatch",
  "style-mismatch",
  "substitute-required",
  "unsupported-format",
  "invalid-metrics",
  "not-embeddable",
  "license-restricted",
  "subsetting-forbidden",
  "subsetting-failed",
]);
export const buildDiagnosticFontUnavailableReasonSchema: z.ZodType<BuildDiagnosticFontUnavailableReason> =
  rawBuildDiagnosticFontUnavailableReasonSchema satisfies z.ZodType<
    AssertExact<
      BuildDiagnosticFontUnavailableReason,
      z.infer<typeof rawBuildDiagnosticFontUnavailableReasonSchema>
    > extends true
      ? BuildDiagnosticFontUnavailableReason
      : never
  >;

export type BuildDiagnostic =
  | {
      code: "invalid-options";
      path: string;
      message: string;
      field?: string | undefined;
      expected?: string | undefined;
    }
  | {
      code: "unknown-prop";
      path: string;
      message: string;
      field?: string | undefined;
      prop: string;
      scope?: "props" | "style" | "options" | undefined;
    }
  | {
      code: "invalid-prop";
      path: string;
      message: string;
      field?: string | undefined;
      expected?: string | undefined;
    }
  | {
      code: "invalid-style";
      path: string;
      message: string;
      field?: string | undefined;
      expected?: string | undefined;
    }
  | {
      code: "invalid-structure";
      path: string;
      message: string;
      field?: string | undefined;
      expected?: string | undefined;
      conflictPath?: string | undefined;
    }
  | {
      code: "duplicate-id";
      path: string;
      message: string;
      field?: string | undefined;
      identifier?: string | undefined;
      conflictPath?: string | undefined;
    }
  | {
      code: "broken-reference";
      path: string;
      message: string;
      field?: string | undefined;
      reference?: string | undefined;
      targetKind?:
        | "slide"
        | "layout"
        | "placeholder"
        | "object"
        | "cue"
        | "scope"
        | "part"
        | "resource"
        | "axis"
        | undefined;
    }
  | {
      code: "invalid-continuity";
      path: string;
      message: string;
      field?: string | undefined;
      reason?:
        | "inside-group"
        | "on-master"
        | "needs-object"
        | "unsupported-object"
        | "kind-mismatch"
        | "name-conflict"
        | undefined;
      conflictPath?: string | undefined;
    }
  | {
      code: "missing-capability";
      path: string;
      message: string;
      field?: string | undefined;
      capability: string;
    }
  | {
      code: "font-unavailable";
      path: string;
      message: string;
      field?: string | undefined;
      font: {
        family: string;
        weight?: number | undefined;
        style?: "normal" | "italic" | undefined;
      };
      reason: BuildDiagnosticFontUnavailableReason;
    }
  | {
      code: "font-missing-glyph";
      path: string;
      message: string;
      field?: string | undefined;
      font: {
        family: string;
        weight: number;
        style: "normal" | "italic";
      };
      characters: readonly string[];
    }
  | {
      code: "invalid-resource";
      path: string;
      message: string;
      field?: string | undefined;
      resourceKind: "image" | "part" | "capability";
    }
  | {
      code: "target-violation";
      path: string;
      message: string;
      field?: string | undefined;
      target: string;
      extension?: string | undefined;
      mechanism?: string | undefined;
    }
  | {
      code: "render-failed";
      path: string;
      message: string;
      field?: string | undefined;
      phase: "expand" | "append";
    };

const rawBuildDiagnosticSchema = z.discriminatedUnion("code", [
  z.strictObject({
    code: z.literal("invalid-options"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    expected: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("unknown-prop"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    prop: z.string(),
    scope: z.enum(["props", "style", "options"]).optional(),
  }),
  z.strictObject({
    code: z.literal("invalid-prop"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    expected: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("invalid-style"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    expected: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("invalid-structure"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    expected: z.string().optional(),
    conflictPath: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("duplicate-id"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    identifier: z.string().optional(),
    conflictPath: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("broken-reference"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    reference: z.string().optional(),
    targetKind: z
      .enum([
        "slide",
        "layout",
        "placeholder",
        "object",
        "cue",
        "scope",
        "part",
        "resource",
        "axis",
      ])
      .optional(),
  }),
  z.strictObject({
    code: z.literal("invalid-continuity"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    reason: z
      .enum([
        "inside-group",
        "on-master",
        "needs-object",
        "unsupported-object",
        "kind-mismatch",
        "name-conflict",
      ])
      .optional(),
    conflictPath: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("missing-capability"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    capability: z.string(),
  }),
  z.strictObject({
    code: z.literal("font-unavailable"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    font: z.strictObject({
      family: z.string(),
      weight: z.number().optional(),
      style: diagnosticFontStyleSchema.optional(),
    }),
    reason: rawBuildDiagnosticFontUnavailableReasonSchema,
  }),
  z.strictObject({
    code: z.literal("font-missing-glyph"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    font: z.strictObject({
      family: z.string(),
      weight: z.number(),
      style: diagnosticFontStyleSchema,
    }),
    characters: z.array(z.string()).readonly(),
  }),
  z.strictObject({
    code: z.literal("invalid-resource"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    resourceKind: z.enum(["image", "part", "capability"]),
  }),
  z.strictObject({
    code: z.literal("target-violation"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    target: z.string(),
    extension: z.string().optional(),
    mechanism: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("render-failed"),
    path: z.string(),
    message: z.string(),
    field: z.string().optional(),
    phase: z.enum(["expand", "append"]),
  }),
]);
/** Typed resolve/emit/expand failure diagnostic discriminated by `code`. */
export const buildDiagnosticSchema: z.ZodType<BuildDiagnostic> =
  rawBuildDiagnosticSchema satisfies z.ZodType<
    AssertExact<
      BuildDiagnostic,
      z.infer<typeof rawBuildDiagnosticSchema>
    > extends true
      ? BuildDiagnostic
      : never
  >;

const previewFontStyleSchema = diagnosticFontStyleSchema;

export type PreviewError =
  | {
      code: "preview-missing-font";
      path: string;
      field?: string | undefined;
      message: string;
      font: {
        family: string;
        weight: number;
        style: "normal" | "italic";
      };
    }
  | {
      code: "preview-unsupported-font";
      path: string;
      field?: string | undefined;
      message: string;
      font: {
        family: string;
        weight?: number | undefined;
        style?: "normal" | "italic" | undefined;
      };
      reason?: string | undefined;
    }
  | {
      code: "preview-missing-glyph";
      path: string;
      field?: string | undefined;
      message: string;
      font: {
        family: string;
        weight: number;
        style: "normal" | "italic";
      };
      characters: readonly string[];
    }
  | {
      code: "preview-corrupt-image";
      path: string;
      field?: string | undefined;
      message: string;
      format?: "png" | "jpeg" | "gif" | undefined;
    }
  | {
      code: "preview-unsupported-capability";
      path: string;
      field?: string | undefined;
      message: string;
      capability: string;
    };

export const previewErrorOptionSchemas = [
  z.strictObject({
    code: z.literal("preview-missing-font"),
    path: z.string(),
    field: z.string().optional(),
    message: z.string(),
    font: z.strictObject({
      family: z.string(),
      weight: z.number(),
      style: previewFontStyleSchema,
    }),
  }),
  z.strictObject({
    code: z.literal("preview-unsupported-font"),
    path: z.string(),
    field: z.string().optional(),
    message: z.string(),
    font: z.strictObject({
      family: z.string(),
      weight: z.number().optional(),
      style: previewFontStyleSchema.optional(),
    }),
    reason: z.string().optional(),
  }),
  z.strictObject({
    code: z.literal("preview-missing-glyph"),
    path: z.string(),
    field: z.string().optional(),
    message: z.string(),
    font: z.strictObject({
      family: z.string(),
      weight: z.number(),
      style: previewFontStyleSchema,
    }),
    characters: z.array(z.string()).readonly(),
  }),
  z.strictObject({
    code: z.literal("preview-corrupt-image"),
    path: z.string(),
    field: z.string().optional(),
    message: z.string(),
    format: z.enum(["png", "jpeg", "gif"]).optional(),
  }),
  z.strictObject({
    code: z.literal("preview-unsupported-capability"),
    path: z.string(),
    field: z.string().optional(),
    message: z.string(),
    capability: z.string(),
  }),
] as const;

const rawPreviewErrorSchema = z.discriminatedUnion(
  "code",
  previewErrorOptionSchemas,
);
/** Typed per-page static preview failure discriminated by `code`. */
export const previewErrorSchema: z.ZodType<PreviewError> =
  rawPreviewErrorSchema satisfies z.ZodType<
    AssertExact<
      PreviewError,
      z.infer<typeof rawPreviewErrorSchema>
    > extends true
      ? PreviewError
      : never
  >;

export type PreviewPlaceholder = {
  id?: string | undefined;
  path: string;
  url: string;
};

const rawPreviewPlaceholderSchema = z.strictObject({
  id: z.string().optional(),
  path: z.string(),
  url: z.string(),
});
/** Offline external image placeholder rendered deterministically during static preview. */
export const previewPlaceholderSchema: z.ZodType<PreviewPlaceholder> =
  rawPreviewPlaceholderSchema satisfies z.ZodType<
    AssertExact<
      PreviewPlaceholder,
      z.infer<typeof rawPreviewPlaceholderSchema>
    > extends true
      ? PreviewPlaceholder
      : never
  >;

/** Error thrown when requested preview page numbers exceed the resolved deck slide count. */
export class PreviewPageRangeError extends RangeError {
  readonly requestedPages: readonly number[];
  readonly totalSlides: number;

  constructor(requestedPages: readonly number[], totalSlides: number) {
    super("Preview page number is outside the document");
    this.name = "PreviewPageRangeError";
    this.requestedPages = Object.freeze([...requestedPages]);
    this.totalSlides = totalSlides;
  }
}
