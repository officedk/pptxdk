import {
  previewErrorSchema,
  previewPlaceholderSchema,
  type PreviewError,
  type PreviewPlaceholder,
} from "@pptxdk/core/build";
import { z } from "zod";

export { previewErrorSchema, previewPlaceholderSchema };
export type { PreviewError, PreviewPlaceholder };

export const previewPageNumberSchema = z.int().positive();

export const previewPagesSchema = z.array(previewPageNumberSchema).readonly();

export const paintOptionsSchema = z
  .strictObject({
    /** 1-based page numbers to paint. Omit to draw all pages; an empty list draws none. */
    pages: previewPagesSchema.optional(),
    signal: z.instanceof(AbortSignal).optional(),
  })
  .readonly();
export type PaintOptions = z.infer<typeof paintOptionsSchema>;

export const previewOptionsSchema = paintOptionsSchema;
export type PreviewOptions = z.infer<typeof previewOptionsSchema>;

export const previewSeveritySchema = z.enum(["error", "warning"]);
export type PreviewSeverity = z.infer<typeof previewSeveritySchema>;

export const previewDiagnosticSchema = z
  .object({
    severity: previewSeveritySchema,
    page: previewPageNumberSchema,
    path: z.string(),
    message: z.string(),
  })
  .readonly();
export type PreviewDiagnostic = z.infer<typeof previewDiagnosticSchema>;

export const screenPointSchema = z
  .object({
    x: z.number(),
    y: z.number(),
  })
  .readonly();
export type ScreenPoint = z.infer<typeof screenPointSchema>;

export const screenGeometrySchema = z
  .object({
    corners: z.array(screenPointSchema).readonly(),
    screenBox: z
      .object({
        x: z.number(),
        y: z.number(),
        width: z.number(),
        height: z.number(),
      })
      .readonly(),
  })
  .readonly();
export type ScreenGeometry = z.infer<typeof screenGeometrySchema>;

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

/** Validates paint/preview options (`pages`, `signal`) via shared pure-data schemas. */
export function validatePaintOptions(options: unknown = {}): PaintOptions {
  const parsed = paintOptionsSchema.safeParse(options);
  if (parsed.success) return parsed.data;

  const issue = parsed.error.issues[0];
  if (issue === undefined) {
    throw new TypeError("Preview options must be an object");
  }
  if (issue.path[0] === "pages") {
    throw new RangeError("Preview page numbers must be positive integers");
  }
  if (issue.path[0] === "signal") {
    throw new TypeError("Preview options signal must be an AbortSignal");
  }
  if (issue.path.length === 0 && issue.code === "unrecognized_keys") {
    throw new TypeError(
      `Preview options contain unknown parameter ${issue.keys[0]}`,
    );
  }
  throw new TypeError("Preview options must be an object");
}
