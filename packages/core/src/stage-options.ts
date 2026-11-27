import { z } from "zod";

export const formatBaseline = "ECMA-376 5th edition Transitional" as const;

export const formatBaselineSchema = z.literal(formatBaseline);

export const extensionSchema = z.enum(["p14", "p159", "c15"]);
export type Extension = z.infer<typeof extensionSchema>;

export const fontEmbeddingModeSchema = z.enum(["full", "subset"]);
export type FontEmbeddingMode = z.infer<typeof fontEmbeddingModeSchema>;

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

const targetNames = [
  targets.transitional.name,
  targets.baselineOnly.name,
] as const;

export const targetNameSchema = z.enum(targetNames);
export type TargetName = z.infer<typeof targetNameSchema>;

export const targetSchema = z
  .strictObject({
    name: z.string().refine((value) => value.trim().length > 0, {
      message: "Target requires a non-empty name",
    }),
    format: formatBaselineSchema,
    extensions: z
      .array(extensionSchema)
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
export type Target = z.infer<typeof targetSchema>;

export const resolveOptionsSchema = z
  .strictObject({
    signal: z.instanceof(AbortSignal).optional(),
  })
  .readonly();
export type ResolveOptions = z.infer<typeof resolveOptionsSchema>;

export const buildOptionsSchema = z
  .strictObject({
    /** Explicitly request embedding full files or character subsets of fonts actually used by the deck; no embedding by default. */
    embedFonts: fontEmbeddingModeSchema.optional(),
    target: targetSchema.optional(),
    /**
     * Cooperative cancellation signal for a single build; no default timeout when unspecified.
     * Completion, unhandled errors, and cancellation follow the first established terminal state; cancelling after completion does not alter the result.
     */
    signal: z.instanceof(AbortSignal).optional(),
  })
  .readonly();
export type BuildOptions = z.infer<typeof buildOptionsSchema>;

export const emitOptionsSchema = buildOptionsSchema;
export type EmitOptions = z.infer<typeof emitOptionsSchema>;

const diagnosticFontStyleSchema = z.enum(["normal", "italic"]);

export const buildDiagnosticFontUnavailableReasonSchema = z.enum([
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
export type BuildDiagnosticFontUnavailableReason = z.infer<
  typeof buildDiagnosticFontUnavailableReasonSchema
>;

/** Typed resolve/emit/expand failure diagnostic discriminated by `code`. */
export const buildDiagnosticSchema = z.discriminatedUnion("code", [
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
    reason: buildDiagnosticFontUnavailableReasonSchema,
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
export type BuildDiagnostic = z.infer<typeof buildDiagnosticSchema>;

const previewFontStyleSchema = diagnosticFontStyleSchema;

/** Typed per-page static preview failure discriminated by `code`. */
export const previewErrorSchema = z.discriminatedUnion("code", [
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
]);
export type PreviewError = z.infer<typeof previewErrorSchema>;

/** Offline external image placeholder rendered deterministically during static preview. */
export const previewPlaceholderSchema = z.strictObject({
  id: z.string().optional(),
  path: z.string(),
  url: z.string(),
});
export type PreviewPlaceholder = z.infer<typeof previewPlaceholderSchema>;

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
