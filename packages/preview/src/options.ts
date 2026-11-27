import {
  PreviewPageRangeError,
  previewErrorSchema,
  previewPlaceholderSchema,
  type PreviewError,
  type PreviewPlaceholder,
} from "@pptxdk/core/build";
import { z } from "zod";

export { PreviewPageRangeError, previewErrorSchema, previewPlaceholderSchema };
export type { PreviewError, PreviewPlaceholder };

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type PaintOptions = {
  /** 1-based page numbers to paint. Omit to draw all pages; an empty list draws none. */
  readonly pages?: readonly number[] | undefined;
  readonly signal?: AbortSignal | undefined;
};

export type PreviewOptions = PaintOptions;

export type ScreenPoint = {
  readonly x: number;
  readonly y: number;
};

export type ScreenGeometry = {
  readonly corners: readonly ScreenPoint[];
  readonly screenBox: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
};

const rawPreviewPageNumberSchema = z.int().positive();
export const previewPageNumberSchema: z.ZodType<number> =
  rawPreviewPageNumberSchema;

const rawPreviewPagesSchema = z.array(rawPreviewPageNumberSchema).readonly();
export const previewPagesSchema: z.ZodType<readonly number[]> =
  rawPreviewPagesSchema;

const rawPaintOptionsSchema = z
  .strictObject({
    /** 1-based page numbers to paint. Omit to draw all pages; an empty list draws none. */
    pages: rawPreviewPagesSchema.optional(),
    signal: z.instanceof(AbortSignal).optional(),
  })
  .readonly();
export const paintOptionsSchema: z.ZodType<PaintOptions> =
  rawPaintOptionsSchema satisfies z.ZodType<
    AssertExact<
      PaintOptions,
      z.infer<typeof rawPaintOptionsSchema>
    > extends true
      ? PaintOptions
      : never
  >;

export const previewOptionsSchema: z.ZodType<PreviewOptions> =
  paintOptionsSchema;

const rawScreenPointSchema = z
  .object({
    x: z.number(),
    y: z.number(),
  })
  .readonly();
export const screenPointSchema: z.ZodType<ScreenPoint> =
  rawScreenPointSchema satisfies z.ZodType<
    AssertExact<ScreenPoint, z.infer<typeof rawScreenPointSchema>> extends true
      ? ScreenPoint
      : never
  >;

const rawScreenGeometrySchema = z
  .object({
    corners: z.array(rawScreenPointSchema).readonly(),
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
export const screenGeometrySchema: z.ZodType<ScreenGeometry> =
  rawScreenGeometrySchema satisfies z.ZodType<
    AssertExact<
      ScreenGeometry,
      z.infer<typeof rawScreenGeometrySchema>
    > extends true
      ? ScreenGeometry
      : never
  >;

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
