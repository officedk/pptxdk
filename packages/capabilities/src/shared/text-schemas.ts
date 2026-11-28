import {
  CapabilityError,
  authorIdSchema,
  layoutStyleShape,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  nonNegativeLayoutSpacingSchema,
  textListBulletSchema,
  textListNumberSchema,
  textListSchema,
  textStyleShape,
  validateTextStyle,
  visualStyleShape,
  type AuthorId,
  type LayoutStyle,
  type LinkTarget,
  type NativePlaceholderRef,
  type NonNegativeLayoutSpacing,
  type TextList,
  type TextStyle,
  type VisualStyle,
} from "@pptxdk/core/capability";
import { pxToEmu, xmlEscape } from "@pptxdk/emit";
import { z } from "zod";

import {
  motionPropsShape,
  type MotionProps,
} from "../motion/motion-schemas.ts";
import { optionalNameSchema } from "./identity.ts";
import { morphKeySchema, type MorphKey } from "./morph-api.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export { textListBulletSchema, textListNumberSchema, textListSchema };
export type { TextList, TextStyle };

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export type RunTextStyle = Pick<
  TextStyle,
  | "fontFamily"
  | "eastAsianFontFamily"
  | "fontSize"
  | "fontWeight"
  | "fontStyle"
  | "color"
  | "letterSpacing"
  | "textTransform"
  | "textDecoration"
  | "textShadow"
  | "lang"
>;

const runStyleObjectSchema = z.strictObject({
  fontFamily: textStyleShape.fontFamily,
  eastAsianFontFamily: textStyleShape.eastAsianFontFamily,
  fontSize: textStyleShape.fontSize,
  fontWeight: textStyleShape.fontWeight,
  fontStyle: textStyleShape.fontStyle,
  color: textStyleShape.color,
  letterSpacing: textStyleShape.letterSpacing,
  textTransform: textStyleShape.textTransform,
  textDecoration: textStyleShape.textDecoration,
  textShadow: textStyleShape.textShadow,
  lang: textStyleShape.lang,
});

const rawRunStyleSchema = runStyleObjectSchema.readonly();
/** Run-level style schema shared by `Span` and `SlideNumber`. */
export const runStyleSchema: z.ZodType<RunTextStyle> =
  rawRunStyleSchema satisfies z.ZodType<
    AssertExact<RunTextStyle, z.infer<typeof rawRunStyleSchema>> extends true
      ? RunTextStyle
      : never
  >;

export type ParagraphStyle = RunTextStyle &
  Pick<TextStyle, "lineHeight" | "align" | "spacingAfter"> & {
    readonly list?: TextList | undefined;
  };

const paragraphStyleObjectSchema = runStyleObjectSchema.extend({
  lineHeight: textStyleShape.lineHeight,
  align: textStyleShape.align,
  spacingAfter: textStyleShape.spacingAfter,
  list: textListSchema.optional(),
});

const rawParagraphStyleSchema = paragraphStyleObjectSchema.readonly();
/** Paragraph-level style schema shared by `P` across Text, Shape, Table Cell, and Notes. */
export const paragraphStyleSchema: z.ZodType<ParagraphStyle> =
  rawParagraphStyleSchema satisfies z.ZodType<
    AssertExact<
      ParagraphStyle,
      z.infer<typeof rawParagraphStyleSchema>
    > extends true
      ? ParagraphStyle
      : never
  >;

export const runStyleKeys = new Set(Object.keys(runStyleObjectSchema.shape));
export const paragraphStyleKeys = new Set(
  Object.keys(paragraphStyleObjectSchema.shape),
);

/** Coordinate bounds for text box `x` and `y` in logical pixels. */
export const textCoordinateSchema = z
  .number()
  .finite()
  .min(-1_000_000)
  .max(1_000_000);

/** Positive pixel dimension bounds convertible to DrawingML EMU. */
export const textPositiveSizeSchema = z
  .number()
  .finite()
  .positive()
  .refine((value) => {
    const extent = pxToEmu(value);
    return extent >= 1 && extent <= 2_147_483_647;
  }, "Converted value exceeds shape dimension range");

const nonNegativeInsetSchema = z
  .number()
  .finite()
  .nonnegative()
  .refine((value) => {
    const converted = pxToEmu(value);
    return Number.isSafeInteger(converted);
  }, "Inset must be within DrawingML EMU range");

const spacingValuesWithinEmuRange = (
  spacing: NonNegativeLayoutSpacing,
): boolean =>
  typeof spacing === "number"
    ? Number.isSafeInteger(pxToEmu(spacing))
    : spacing.every((entry) => Number.isSafeInteger(pxToEmu(entry)));

export const textPaddingValueSchema = nonNegativeLayoutSpacingSchema.refine(
  spacingValuesWithinEmuRange,
  "Inset must be within DrawingML EMU range",
);

export const textInsetsStyleSchema = z
  .object({
    padding: textPaddingValueSchema.optional(),
    paddingTop: nonNegativeInsetSchema.optional(),
    paddingRight: nonNegativeInsetSchema.optional(),
    paddingBottom: nonNegativeInsetSchema.optional(),
    paddingLeft: nonNegativeInsetSchema.optional(),
  })
  .readonly();

export type TextBoxStyle = LayoutStyle & TextStyle & VisualStyle;

const textBoxStyleObjectSchema = z.strictObject({
  ...layoutStyleShape,
  ...textStyleShape,
  ...visualStyleShape,
});
export const textBoxStyleKeys: readonly string[] = Object.keys(
  textBoxStyleObjectSchema.shape,
);
const rawTextBoxStyleSchema = textBoxStyleObjectSchema.readonly();
export const textBoxStyleSchema: z.ZodType<TextBoxStyle> =
  rawTextBoxStyleSchema satisfies z.ZodType<
    AssertExact<
      TextBoxStyle,
      z.infer<typeof rawTextBoxStyleSchema>
    > extends true
      ? TextBoxStyle
      : never
  >;

export type TextDeclarationProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly name?: string | undefined;
  readonly placeholder?: string | undefined;
  readonly nativePlaceholder?: NativePlaceholderRef | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  readonly style?: TextBoxStyle | undefined;
};

const textAuthorPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  style: rawTextBoxStyleSchema.optional(),
  ...motionPropsShape,
});

export const textAuthorPropsKeys: readonly string[] = Object.keys(
  textAuthorPropsObjectSchema.shape,
);
const rawTextAuthorPropsSchema = textAuthorPropsObjectSchema.readonly();
export const textAuthorPropsSchema: z.ZodType<TextDeclarationProps> =
  rawTextAuthorPropsSchema satisfies z.ZodType<
    AssertExact<
      TextDeclarationProps,
      z.infer<typeof rawTextAuthorPropsSchema>
    > extends true
      ? TextDeclarationProps
      : never
  >;

export const textBoxAllowedParamKeys = new Set([
  ...textAuthorPropsKeys,
  "children",
]);

export type ParagraphDeclarationProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly style?: ParagraphStyle | undefined;
};

const paragraphAuthorPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  style: rawParagraphStyleSchema.optional(),
  ...motionPropsShape,
});
export const paragraphAuthorPropsKeys: readonly string[] = Object.keys(
  paragraphAuthorPropsObjectSchema.shape,
);
const rawParagraphAuthorPropsSchema =
  paragraphAuthorPropsObjectSchema.readonly();
export const paragraphAuthorPropsSchema: z.ZodType<ParagraphDeclarationProps> =
  rawParagraphAuthorPropsSchema satisfies z.ZodType<
    AssertExact<
      ParagraphDeclarationProps,
      z.infer<typeof rawParagraphAuthorPropsSchema>
    > extends true
      ? ParagraphDeclarationProps
      : never
  >;

export type SpanDeclarationProps = {
  readonly id?: AuthorId | undefined;
  readonly style?: RunTextStyle | undefined;
};

const spanAuthorPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  style: rawRunStyleSchema.optional(),
});
export const spanAuthorPropsKeys: readonly string[] = Object.keys(
  spanAuthorPropsObjectSchema.shape,
);
const rawSpanAuthorPropsSchema = spanAuthorPropsObjectSchema.readonly();
export const spanAuthorPropsSchema: z.ZodType<SpanDeclarationProps> =
  rawSpanAuthorPropsSchema satisfies z.ZodType<
    AssertExact<
      SpanDeclarationProps,
      z.infer<typeof rawSpanAuthorPropsSchema>
    > extends true
      ? SpanDeclarationProps
      : never
  >;

export type SlideNumberDeclarationProps = SpanDeclarationProps;
export const slideNumberAuthorPropsKeys = spanAuthorPropsKeys;
export const slideNumberAuthorPropsSchema: z.ZodType<SlideNumberDeclarationProps> =
  spanAuthorPropsSchema;

/** Parses and validates a paragraph list declaration, preserving DrawingML error codes. */
export function parseTextList(value: unknown): TextList {
  if (
    !isRecord(value) ||
    (value["type"] !== "bullet" && value["type"] !== "number")
  ) {
    throw new CapabilityError(
      "invalid-list",
      "list.type must be bullet or number",
    );
  }
  const parsed = textListSchema.safeParse(value);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue?.code === "unrecognized_keys") {
      throw new CapabilityError(
        "unknown-style-property",
        `list contains unimplemented style ${issue.keys[0]}`,
      );
    }
    const field = issue?.path[0];
    if (field === "level") {
      throw new CapabilityError(
        "invalid-list",
        "list.level must be an integer between 0 and 8",
      );
    }
    if (field === "character") {
      throw new CapabilityError(
        "invalid-list",
        "Bullet must be a single Unicode character",
      );
    }
    if (field === "style") {
      throw new CapabilityError(
        "invalid-list",
        "Numbering style currently only supports arabicPeriod",
      );
    }
    if (field === "startAt") {
      throw new CapabilityError(
        "invalid-list",
        "Numbering start value must be an integer between 1 and 32767",
      );
    }
    throw new CapabilityError("invalid-list", issue?.message ?? "Invalid list");
  }
  if (parsed.data.type === "bullet") {
    const character = parsed.data.character ?? "•";
    try {
      xmlEscape(character);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "Bullet contains characters not allowed in XML 1.0",
      );
    }
  }
  return parsed.data;
}

function checkedStyleObject(
  value: unknown,
  allowedKeys: ReadonlySet<string>,
  label: string,
): Record<string, unknown> {
  if (value === undefined) return {};
  if (!isRecord(value)) {
    throw new CapabilityError("invalid-style", `${label} must be an object`);
  }
  const unknown = Object.keys(value).find((key) => !allowedKeys.has(key));
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-style-property",
      `${label} contains unimplemented style ${unknown}`,
    );
  }
  return value;
}

/** Validates run-level style declarations using shared core text schemas. */
export function parseRunStyle(
  value: unknown,
  label = "Span.style",
): RunTextStyle {
  const props = checkedStyleObject(value, runStyleKeys, label);
  return validateTextStyle(props);
}

/** Validates paragraph-level style declarations including `list`. */
export function parseParagraphStyle(
  value: unknown,
  label = "P.style",
): ParagraphStyle {
  const props = checkedStyleObject(value, paragraphStyleKeys, label);
  const { list, ...textProps } = props;
  validateTextStyle(textProps);
  const parsedList = list === undefined ? undefined : parseTextList(list);
  return {
    ...(textProps as Omit<TextStyle, "verticalAlign" | "textWrap">),
    ...(parsedList === undefined ? {} : { list: parsedList }),
  };
}

/** Validates text box padding/inset declarations via Zod and converts edges to pixel numbers. */
export function parseTextInsets(value: unknown): {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
} {
  const style = isRecord(value) ? value : {};
  const candidate = {
    ...(style["padding"] === undefined ? {} : { padding: style["padding"] }),
    ...(style["paddingTop"] === undefined
      ? {}
      : { paddingTop: style["paddingTop"] }),
    ...(style["paddingRight"] === undefined
      ? {}
      : { paddingRight: style["paddingRight"] }),
    ...(style["paddingBottom"] === undefined
      ? {}
      : { paddingBottom: style["paddingBottom"] }),
    ...(style["paddingLeft"] === undefined
      ? {}
      : { paddingLeft: style["paddingLeft"] }),
  };
  const parsed = textInsetsStyleSchema.safeParse(candidate);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue?.path[0] === "padding") {
      const rawPadding = candidate.padding;
      const isTupleOrNumber =
        typeof rawPadding === "number" ||
        (Array.isArray(rawPadding) &&
          (rawPadding.length === 2 || rawPadding.length === 4));
      throw new CapabilityError(
        "invalid-text-insets",
        isTupleOrNumber
          ? "padding inset must be a finite non-negative pixel number"
          : "padding must be a number, a 2-tuple, or a 4-tuple",
      );
    }
    const edge = String(issue?.path[0] ?? "padding").replace("padding", "");
    const label = edge === "" ? "padding" : edge.toLowerCase();
    throw new CapabilityError(
      "invalid-text-insets",
      `${label} inset must be a finite non-negative pixel number`,
    );
  }
  const raw = parsed.data.padding;
  let edges = { top: 0, right: 0, bottom: 0, left: 0 };
  if (typeof raw === "number") {
    edges = { top: raw, right: raw, bottom: raw, left: raw };
  } else if (Array.isArray(raw)) {
    edges =
      raw.length === 2
        ? { top: raw[0], right: raw[1], bottom: raw[0], left: raw[1] }
        : { top: raw[0], right: raw[1], bottom: raw[2], left: raw[3] };
  }
  return {
    top: parsed.data.paddingTop ?? edges.top,
    right: parsed.data.paddingRight ?? edges.right,
    bottom: parsed.data.paddingBottom ?? edges.bottom,
    left: parsed.data.paddingLeft ?? edges.left,
  };
}

/** Validates a text box coordinate (`x` or `y`) with Zod schema bounds. */
export function parseTextCoordinate(
  value: unknown,
  label: "x" | "y",
  path?: string,
): number {
  const parsed = textCoordinateSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} must be a finite pixel number between -1000000 and 1000000`,
      path,
    );
  }
  return parsed.data;
}

/** Validates a text box extent (`width` or `height`) with Zod schema bounds. */
export function parseTextExtent(
  value: unknown,
  label: string,
  path?: string,
): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    if (value === 0 || pxToEmu(value) === 0) {
      throw new CapabilityError(
        "invalid-text-box",
        `Computed ${label} is 0px; ${label} must be greater than 0px`,
        path,
      );
    }
    if (value < 0 || pxToEmu(value) < 1) {
      throw new CapabilityError(
        "invalid-text-box",
        `Computed ${label} is ${value}px; ${label} must be greater than 0px`,
        path,
      );
    }
  } else {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} must be a finite positive pixel number`,
      path,
    );
  }
  const parsed = textPositiveSizeSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} converted value exceeds shape dimension range`,
      path,
    );
  }
  return parsed.data;
}
