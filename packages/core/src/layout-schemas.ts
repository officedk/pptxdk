import { z } from "zod";

import { BuildError } from "#/diagnostics.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type LayoutLength = number | `${number}%` | "auto";

export type LayoutSpacing =
  | number
  | readonly [number, number]
  | readonly [number, number, number, number];

export type NonNegativeLayoutSpacing =
  | number
  | readonly [number, number]
  | readonly [number, number, number, number];

export type PaddingStyle = {
  readonly padding?: NonNegativeLayoutSpacing | undefined;
  readonly paddingTop?: number | undefined;
  readonly paddingRight?: number | undefined;
  readonly paddingBottom?: number | undefined;
  readonly paddingLeft?: number | undefined;
};

/** Fixed canvas CSS layout subset; numeric lengths are interpreted in pixels. */
export type LayoutStyle = {
  readonly display?: "block" | "flex" | "grid" | "none" | undefined;
  readonly position?: "static" | "relative" | "absolute" | undefined;
  readonly inset?: LayoutLength | undefined;
  readonly left?: LayoutLength | undefined;
  readonly top?: LayoutLength | undefined;
  readonly right?: LayoutLength | undefined;
  readonly bottom?: LayoutLength | undefined;
  readonly width?: LayoutLength | undefined;
  readonly height?: LayoutLength | undefined;
  readonly minWidth?: LayoutLength | undefined;
  readonly minHeight?: LayoutLength | undefined;
  readonly maxWidth?: LayoutLength | undefined;
  readonly maxHeight?: LayoutLength | undefined;
  readonly margin?: LayoutSpacing | undefined;
  readonly marginTop?: number | undefined;
  readonly marginRight?: number | undefined;
  readonly marginBottom?: number | undefined;
  readonly marginLeft?: number | undefined;
  readonly padding?: NonNegativeLayoutSpacing | undefined;
  readonly paddingTop?: number | undefined;
  readonly paddingRight?: number | undefined;
  readonly paddingBottom?: number | undefined;
  readonly paddingLeft?: number | undefined;
  readonly gap?: number | undefined;
  readonly rowGap?: number | undefined;
  readonly columnGap?: number | undefined;
  readonly flexDirection?:
    | "row"
    | "column"
    | "row-reverse"
    | "column-reverse"
    | undefined;
  readonly flexWrap?: "nowrap" | "wrap" | undefined;
  readonly justifyContent?:
    | "start"
    | "end"
    | "center"
    | "flex-start"
    | "flex-end"
    | "space-between"
    | "space-around"
    | "space-evenly"
    | undefined;
  readonly alignItems?:
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end"
    | "baseline"
    | undefined;
  readonly alignSelf?:
    | "auto"
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end"
    | "baseline"
    | undefined;
  readonly alignContent?:
    | "start"
    | "end"
    | "center"
    | "flex-start"
    | "flex-end"
    | "space-between"
    | "space-around"
    | "space-evenly"
    | "stretch"
    | undefined;
  /**
   * Non-negative flex shorthand (equivalent to `flexGrow: value, flexShrink: 1, flexBasis: 0`).
   *
   * Flex items default to automatic minimum content size (`minWidth: "auto"` / `minHeight: "auto"`).
   * For equal-width columns regardless of inner content size, also set `minWidth: 0` or use CSS Grid
   * (`gridTemplateColumns: "1fr 1fr"`).
   */
  readonly flex?: number | undefined;
  readonly flexGrow?: number | undefined;
  readonly flexShrink?: number | undefined;
  readonly flexBasis?: LayoutLength | undefined;
  readonly aspectRatio?: number | undefined;
  readonly gridTemplateColumns?: string | undefined;
  readonly gridTemplateRows?: string | undefined;
  /** Track sizes for implicit Grid rows, cycling in order like CSS `grid-auto-rows`; defaults to `auto`. */
  readonly gridAutoRows?: string | undefined;
  readonly gridColumn?: string | number | undefined;
  readonly gridRow?: string | number | undefined;
  readonly justifyItems?:
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end"
    | undefined;
  readonly justifySelf?:
    | "auto"
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end"
    | undefined;
  readonly zIndex?: number | undefined;
};

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

const rawLayoutLengthSchema = z.union([
  z.number().finite(),
  percentLengthSchema,
  z.literal("auto"),
]);
/** Fixed canvas length; numeric values are interpreted in pixels, percentages are evaluated against the parent layout region. */
export const layoutLengthSchema: z.ZodType<LayoutLength> =
  rawLayoutLengthSchema satisfies z.ZodType<
    AssertExact<
      LayoutLength,
      z.infer<typeof rawLayoutLengthSchema>
    > extends true
      ? LayoutLength
      : never
  >;

const rawNonNegativeLayoutLengthSchema = z.union([
  z.number().finite().nonnegative(),
  nonNegativePercentLengthSchema,
  z.literal("auto"),
]);
export const nonNegativeLayoutLengthSchema: z.ZodType<LayoutLength> =
  rawNonNegativeLayoutLengthSchema;

const rawLayoutSpacingSchema = z.union([
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
export const layoutSpacingSchema: z.ZodType<LayoutSpacing> =
  rawLayoutSpacingSchema satisfies z.ZodType<
    AssertExact<
      LayoutSpacing,
      z.infer<typeof rawLayoutSpacingSchema>
    > extends true
      ? LayoutSpacing
      : never
  >;

const rawNonNegativeLayoutSpacingSchema = z.union([
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
export const nonNegativeLayoutSpacingSchema: z.ZodType<NonNegativeLayoutSpacing> =
  rawNonNegativeLayoutSpacingSchema satisfies z.ZodType<
    AssertExact<
      NonNegativeLayoutSpacing,
      z.infer<typeof rawNonNegativeLayoutSpacingSchema>
    > extends true
      ? NonNegativeLayoutSpacing
      : never
  >;

const paddingStyleObjectSchema = z.strictObject({
  padding: rawNonNegativeLayoutSpacingSchema.optional(),
  paddingTop: z.number().finite().nonnegative().optional(),
  paddingRight: z.number().finite().nonnegative().optional(),
  paddingBottom: z.number().finite().nonnegative().optional(),
  paddingLeft: z.number().finite().nonnegative().optional(),
});
export const paddingStyleShape = paddingStyleObjectSchema.shape;
const rawPaddingStyleSchema = paddingStyleObjectSchema.readonly();
export const paddingStyleSchema: z.ZodType<PaddingStyle> =
  rawPaddingStyleSchema satisfies z.ZodType<
    AssertExact<
      PaddingStyle,
      z.infer<typeof rawPaddingStyleSchema>
    > extends true
      ? PaddingStyle
      : never
  >;

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
  inset: rawLayoutLengthSchema.optional(),
  left: rawLayoutLengthSchema.optional(),
  top: rawLayoutLengthSchema.optional(),
  right: rawLayoutLengthSchema.optional(),
  bottom: rawLayoutLengthSchema.optional(),
  width: rawNonNegativeLayoutLengthSchema.optional(),
  height: rawNonNegativeLayoutLengthSchema.optional(),
  minWidth: rawNonNegativeLayoutLengthSchema.optional(),
  minHeight: rawNonNegativeLayoutLengthSchema.optional(),
  maxWidth: rawNonNegativeLayoutLengthSchema.optional(),
  maxHeight: rawNonNegativeLayoutLengthSchema.optional(),
  margin: rawLayoutSpacingSchema.optional(),
  marginTop: z.number().finite().optional(),
  marginRight: z.number().finite().optional(),
  marginBottom: z.number().finite().optional(),
  marginLeft: z.number().finite().optional(),
  padding: rawNonNegativeLayoutSpacingSchema.optional(),
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
  flexBasis: rawNonNegativeLayoutLengthSchema.optional(),
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

export const layoutStyleShape = layoutStyleObjectSchema.shape;
const rawLayoutStyleSchema = layoutStyleObjectSchema.readonly();
export const layoutStyleSchema: z.ZodType<LayoutStyle> =
  rawLayoutStyleSchema satisfies z.ZodType<
    AssertExact<LayoutStyle, z.infer<typeof rawLayoutStyleSchema>> extends true
      ? LayoutStyle
      : never
  >;

const slideLayoutStyleObjectSchema = layoutStyleObjectSchema.omit({
  zIndex: true,
});
export const slideLayoutStyleShape = slideLayoutStyleObjectSchema.shape;
export const slideLayoutStyleSchema: z.ZodType<Omit<LayoutStyle, "zIndex">> =
  slideLayoutStyleObjectSchema.readonly();

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
  const message =
    key === undefined
      ? (issue?.message ?? "Invalid layout style")
      : `${key}: ${issue?.message ?? "invalid value"}`;
  throw new BuildError([
    {
      code: "invalid-style",
      path,
      message,
      ...(key === undefined ? {} : { field: `style.${key}` }),
    },
  ]);
}
