import { z } from "zod";

import { BuildError } from "#/diagnostics.ts";

const percentPattern = /^(-?(?:\d+(?:\.\d*)?|\.\d+))%$/u;
const nonNegativePercentPattern = /^(\d+(?:\.\d*)?|\.\d+)%$/u;

const percentLengthSchema = z
  .string()
  .regex(
    percentPattern,
    "must be a pixel number, percentage, or auto",
  ) as unknown as z.ZodType<`${number}%`>;

const nonNegativePercentLengthSchema = z
  .string()
  .regex(
    nonNegativePercentPattern,
    "must be a non-negative pixel number, percentage, or auto",
  ) as unknown as z.ZodType<`${number}%`>;

/** Fixed canvas length; numeric values are interpreted in pixels, percentages are evaluated against the parent layout region. */
export const layoutLengthSchema = z.union([
  z.number().finite(),
  percentLengthSchema,
  z.literal("auto"),
]);
export type LayoutLength = z.infer<typeof layoutLengthSchema>;

export const nonNegativeLayoutLengthSchema = z.union([
  z.number().finite().nonnegative(),
  nonNegativePercentLengthSchema,
  z.literal("auto"),
]);

export const layoutSpacingSchema = z.union([
  z.number().finite(),
  z.tuple([z.number().finite(), z.number().finite()]).readonly(),
  z
    .tuple([
      z.number().finite(),
      z.number().finite(),
      z.number().finite(),
      z.number().finite(),
    ])
    .readonly(),
]);
export type LayoutSpacing = z.infer<typeof layoutSpacingSchema>;

export const nonNegativeLayoutSpacingSchema = z.union([
  z.number().finite().nonnegative(),
  z
    .tuple([
      z.number().finite().nonnegative(),
      z.number().finite().nonnegative(),
    ])
    .readonly(),
  z
    .tuple([
      z.number().finite().nonnegative(),
      z.number().finite().nonnegative(),
      z.number().finite().nonnegative(),
      z.number().finite().nonnegative(),
    ])
    .readonly(),
]);

export const paddingStyleSchema = z
  .strictObject({
    padding: nonNegativeLayoutSpacingSchema.optional(),
    paddingTop: z.number().finite().nonnegative().optional(),
    paddingRight: z.number().finite().nonnegative().optional(),
    paddingBottom: z.number().finite().nonnegative().optional(),
    paddingLeft: z.number().finite().nonnegative().optional(),
  })
  .readonly();
export type PaddingStyle = z.infer<typeof paddingStyleSchema>;

export const layoutDisplaySchema = z.enum(["block", "flex", "grid", "none"]);
export const layoutPositionSchema = z.enum(["static", "relative", "absolute"]);
export const flexDirectionSchema = z.enum([
  "row",
  "column",
  "row-reverse",
  "column-reverse",
]);
export const flexWrapSchema = z.enum(["nowrap", "wrap"]);
export const justifyContentSchema = z.enum([
  "start",
  "end",
  "center",
  "flex-start",
  "flex-end",
  "space-between",
  "space-around",
  "space-evenly",
]);
export const alignItemsSchema = z.enum([
  "start",
  "end",
  "center",
  "stretch",
  "flex-start",
  "flex-end",
  "baseline",
]);
export const alignSelfSchema = z.enum([
  "auto",
  "start",
  "end",
  "center",
  "stretch",
  "flex-start",
  "flex-end",
  "baseline",
]);
export const alignContentSchema = z.enum([
  "start",
  "end",
  "center",
  "flex-start",
  "flex-end",
  "space-between",
  "space-around",
  "space-evenly",
  "stretch",
]);
export const justifyItemsSchema = z.enum([
  "start",
  "end",
  "center",
  "stretch",
  "flex-start",
  "flex-end",
]);
export const justifySelfSchema = z.enum([
  "auto",
  "start",
  "end",
  "center",
  "stretch",
  "flex-start",
  "flex-end",
]);

/** Fixed canvas CSS layout subset; numeric lengths are interpreted in pixels. */
const layoutStyleObjectSchema = z.strictObject({
  display: layoutDisplaySchema.optional(),
  position: layoutPositionSchema.optional(),
  inset: layoutLengthSchema.optional(),
  left: layoutLengthSchema.optional(),
  top: layoutLengthSchema.optional(),
  right: layoutLengthSchema.optional(),
  bottom: layoutLengthSchema.optional(),
  width: nonNegativeLayoutLengthSchema.optional(),
  height: nonNegativeLayoutLengthSchema.optional(),
  minWidth: nonNegativeLayoutLengthSchema.optional(),
  minHeight: nonNegativeLayoutLengthSchema.optional(),
  maxWidth: nonNegativeLayoutLengthSchema.optional(),
  maxHeight: nonNegativeLayoutLengthSchema.optional(),
  margin: layoutSpacingSchema.optional(),
  marginTop: z.number().finite().optional(),
  marginRight: z.number().finite().optional(),
  marginBottom: z.number().finite().optional(),
  marginLeft: z.number().finite().optional(),
  padding: nonNegativeLayoutSpacingSchema.optional(),
  paddingTop: z.number().finite().nonnegative().optional(),
  paddingRight: z.number().finite().nonnegative().optional(),
  paddingBottom: z.number().finite().nonnegative().optional(),
  paddingLeft: z.number().finite().nonnegative().optional(),
  gap: z.number().finite().nonnegative().optional(),
  rowGap: z.number().finite().nonnegative().optional(),
  columnGap: z.number().finite().nonnegative().optional(),
  flexDirection: flexDirectionSchema.optional(),
  flexWrap: flexWrapSchema.optional(),
  justifyContent: justifyContentSchema.optional(),
  alignItems: alignItemsSchema.optional(),
  alignSelf: alignSelfSchema.optional(),
  alignContent: alignContentSchema.optional(),
  flex: z.number().finite().nonnegative().optional(),
  flexGrow: z.number().finite().nonnegative().optional(),
  flexShrink: z.number().finite().nonnegative().optional(),
  flexBasis: nonNegativeLayoutLengthSchema.optional(),
  aspectRatio: z.number().finite().positive().optional(),
  gridTemplateColumns: z.string().optional(),
  gridTemplateRows: z.string().optional(),
  /** Track sizes for implicit Grid rows, cycling in order like CSS `grid-auto-rows`; defaults to `auto`. */
  gridAutoRows: z.string().optional(),
  gridColumn: z.union([z.number(), z.string()]).optional(),
  gridRow: z.union([z.number(), z.string()]).optional(),
  justifyItems: justifyItemsSchema.optional(),
  justifySelf: justifySelfSchema.optional(),
  zIndex: z.number().finite().optional(),
});

export const layoutStyleSchema = layoutStyleObjectSchema.readonly();
export type LayoutStyle = z.infer<typeof layoutStyleSchema>;

const slideLayoutStyleObjectSchema = layoutStyleObjectSchema.omit({
  zIndex: true,
});
export const slideLayoutStyleShape = slideLayoutStyleObjectSchema.shape;
export const slideLayoutStyleSchema = slideLayoutStyleObjectSchema.readonly();

const enumLayoutStyleKeys = new Set<string>([
  "display",
  "position",
  "flexDirection",
  "flexWrap",
  "justifyContent",
  "alignItems",
  "alignSelf",
  "alignContent",
  "justifyItems",
  "justifySelf",
]);

const layoutSchemaKeys = new Set<string>(
  Object.keys(layoutStyleObjectSchema.shape),
);

/**
 * Validates pure value shapes of layout properties present in `style`, preserving existing
 * BuildError diagnostic codes (`unsupported-layout-value` for enum mismatches, `invalid-layout-style`
 * for invalid lengths, spacing, numbers, or grid track/placement types).
 */
export function validateLayoutStyleValues(
  style: Readonly<Record<string, unknown>>,
  path: string,
): void {
  const layoutEntries: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(style)) {
    if (layoutSchemaKeys.has(key) && value !== undefined) {
      layoutEntries[key] = value;
    }
  }
  const parsed = layoutStyleSchema.safeParse(layoutEntries);
  if (parsed.success) return;

  const issue = parsed.error.issues[0];
  const key = typeof issue?.path[0] === "string" ? issue.path[0] : undefined;
  const code =
    key !== undefined && enumLayoutStyleKeys.has(key)
      ? "unsupported-layout-value"
      : "invalid-layout-style";
  const message =
    key === undefined
      ? (issue?.message ?? "Invalid layout style")
      : `${key}: ${issue?.message ?? "invalid value"}`;
  throw new BuildError([{ severity: "error", path, code, message }]);
}
