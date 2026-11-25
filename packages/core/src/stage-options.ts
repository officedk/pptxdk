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

export const diagnosticSeveritySchema = z.enum(["error", "warning"]);
export type DiagnosticSeverity = z.infer<typeof diagnosticSeveritySchema>;

export const diagnosticSchema = z
  .object({
    severity: diagnosticSeveritySchema,
    path: z.string(),
    code: z.string(),
    message: z.string(),
  })
  .readonly();
export type Diagnostic = z.infer<typeof diagnosticSchema>;
