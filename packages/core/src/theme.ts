import { z } from "zod";

import { CapabilityError } from "#/diagnostics.ts";
import { snapshotValue } from "#/input-values.ts";

export const fontFaceStyleSchema = z.enum(["normal", "italic"]);
export type FontFaceStyle = z.infer<typeof fontFaceStyleSchema>;

const fontWeightNumberSchema = z
  .int()
  .min(100)
  .max(900)
  .refine(
    (value) => value % 100 === 0,
    "fontWeight must be normal, bold, or a hundreds integer between 100 and 900",
  );

const fontWeightSchema = z.union([
  z.literal("normal"),
  z.literal("bold"),
  fontWeightNumberSchema,
]);

export const uint8ArraySchema = z.instanceof(
  Uint8Array,
) as unknown as z.ZodType<Uint8Array>;

/** Font files provided by the caller can be used for measurement or embedded according to build options. */
export const fontFaceSchema = z
  .strictObject({
    weight: fontWeightNumberSchema.optional(),
    style: fontFaceStyleSchema.optional(),
    data: uint8ArraySchema
      .refine((bytes) => bytes.byteLength > 0, "Font face data cannot be empty")
      .optional(),
  })
  .readonly();
export type FontFace = z.infer<typeof fontFaceSchema>;

const nonEmptyTrimmedStringSchema = z
  .string()
  .refine((value) => value.trim() !== "", "Value cannot be empty");

/** Font family and explicit measurement resources referenceable across deck themes and components. */
export const fontSchema = z
  .strictObject({
    family: nonEmptyTrimmedStringSchema,
    faces: z.array(fontFaceSchema).readonly().optional(),
    substitute: z.boolean().optional(),
    embeddable: z.boolean().optional(),
  })
  .readonly();
export type Font = z.infer<typeof fontSchema>;

export const fontOptionsSchema = z
  .strictObject({
    substitute: z.boolean().optional(),
    embeddable: z.boolean().optional(),
  })
  .readonly();
export type FontOptions = z.infer<typeof fontOptionsSchema>;

export const fontReferenceSchema = z.union([
  nonEmptyTrimmedStringSchema,
  fontSchema,
]);
export type FontReference = z.infer<typeof fontReferenceSchema>;

export const fontTokenSchema = fontReferenceSchema;
export type FontToken = z.infer<typeof fontTokenSchema>;

const textLengthPattern = /^(\d+(?:\.\d+)?)(px|pt)$/u;
const textLengthStringSchema = z
  .string()
  .regex(textLengthPattern) as unknown as z.ZodType<
  `${number}px` | `${number}pt`
>;

function textLengthToPixels(value: string | number): number {
  if (typeof value === "number") return value;
  const match = textLengthPattern.exec(value);
  if (match === null) return Number.NaN;
  const amount = Number(match[1]);
  return match[2] === "pt" ? (amount * 4) / 3 : amount;
}

function fontSizeToPoints(value: string | number): number {
  if (typeof value === "string" && value.endsWith("pt")) {
    return Number(value.slice(0, -2));
  }
  return textLengthToPixels(value) * 0.75;
}

export const fontSizeSchema = z.union([
  z.number().finite().positive(),
  textLengthStringSchema.refine(
    (value) => textLengthToPixels(value) > 0,
    "Text length must be greater than 0",
  ),
]);
export type FontSize = z.infer<typeof fontSizeSchema>;

const validatedFontSizeSchema = fontSizeSchema.refine((value) => {
  const serialized = Math.round(fontSizeToPoints(value) * 100);
  return serialized >= 100 && serialized <= 400_000;
}, "fontSize converted value must be within DrawingML range of 1 to 4000pt");

const validatedLineHeightSchema = fontSizeSchema.refine((value) => {
  const centipoints = Math.round(textLengthToPixels(value) * 75);
  return centipoints >= 0 && centipoints <= 158_400;
}, "lineHeight converted value must be within DrawingML range of 0 to 1584pt");

export const textAlignmentSchema = z.enum([
  "left",
  "center",
  "right",
  "justify",
]);
export type TextAlignment = z.infer<typeof textAlignmentSchema>;

export const nativeColorSlots = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export const nativeColorSlotSchema = z.enum(nativeColorSlots);
export type NativeColorSlot = z.infer<typeof nativeColorSlotSchema>;

/** Native theme color represented with DrawingML luminance modulation. */
export const themeShadeSchema = z
  .strictObject({
    slot: nativeColorSlotSchema,
    lumMod: z.number().finite().min(0).max(1),
    lumOff: z.number().finite().min(0).max(1).optional(),
  })
  .readonly();
export type ThemeShade = z.infer<typeof themeShadeSchema>;

export const themeColorSchema = z.union([z.string(), themeShadeSchema]);
export type ThemeColor = z.infer<typeof themeColorSchema>;

/** Color entry can directly provide color, theme slot, or theme token, optionally with alpha. */
export const colorValueSchema = z.union([
  z.string(),
  themeShadeSchema,
  z
    .strictObject({
      color: themeColorSchema,
      alpha: z.number().finite().min(0).max(1).optional(),
    })
    .readonly(),
]);
export type ColorValue = z.infer<typeof colorValueSchema>;

/** Final theme value applied to a color entry. */
export const resolvedColorSchema = z
  .strictObject({
    hex: z.string(),
    alpha: z.number().finite().min(0).max(1).optional(),
    slot: nativeColorSlotSchema.optional(),
    lumMod: z.number().finite().min(0).max(1).optional(),
    lumOff: z.number().finite().min(0).max(1).optional(),
  })
  .readonly();
export type ResolvedColor = z.infer<typeof resolvedColorSchema>;

export const gradientStopSchema = z
  .strictObject({
    color: colorValueSchema,
    position: z.number().finite(),
  })
  .readonly();
export type GradientStop = z.infer<typeof gradientStopSchema>;

export const linearGradientSchema = z
  .strictObject({
    type: z.literal("linear-gradient"),
    angle: z.number().finite(),
    stops: z.array(gradientStopSchema).readonly(),
  })
  .readonly();
export type LinearGradient = z.infer<typeof linearGradientSchema>;

export const radialGradientSchema = z
  .strictObject({
    type: z.literal("radial-gradient"),
    stops: z.array(gradientStopSchema).readonly(),
  })
  .readonly();
export type RadialGradient = z.infer<typeof radialGradientSchema>;

export const shadowSchema = z
  .strictObject({
    color: colorValueSchema.optional(),
    blur: z.number().finite().nonnegative().optional(),
    distance: z.number().finite().nonnegative().optional(),
    angle: z.number().finite().optional(),
  })
  .readonly();
export type Shadow = z.infer<typeof shadowSchema>;

/** Visual fields shared across native DrawingML shapes, images, text boxes, and visible Boxes. */
export const visualStyleSchema = z
  .strictObject({
    backgroundColor: colorValueSchema.optional(),
    backgroundImage: z
      .union([
        linearGradientSchema,
        radialGradientSchema,
        z
          .strictObject({
            image: z.union([z.string(), z.instanceof(URL), uint8ArraySchema]),
          })
          .readonly(),
      ])
      .optional(),
    borderColor: colorValueSchema.optional(),
    borderWidth: z.number().finite().optional(),
    borderStyle: z.enum(["solid", "dashed", "dotted"]).optional(),
    borderRadius: z.number().finite().optional(),
    opacity: z.number().finite().optional(),
    boxShadow: shadowSchema.optional(),
    softEdge: z.number().finite().optional(),
    rotate: z.number().finite().optional(),
  })
  .readonly();
export type VisualStyle = z.infer<typeof visualStyleSchema>;

export const bcp47LangSchema = z
  .string()
  .refine(
    (value) =>
      value.trim() !== "" && /^[A-Za-z0-9]+(-[A-Za-z0-9]+)*$/.test(value),
    "Text language must be a valid BCP 47 language tag",
  )
  .refine((value) => {
    try {
      const canonical = Intl.getCanonicalLocales(value);
      return Boolean(canonical && canonical.length === 1);
    } catch {
      return false;
    }
  }, "Text language must be a valid BCP 47 language tag");

const textStyleObjectSchema = z.strictObject({
  fontFamily: fontReferenceSchema.optional(),
  fontSize: validatedFontSizeSchema.optional(),
  fontWeight: fontWeightSchema.optional(),
  fontStyle: fontFaceStyleSchema.optional(),
  color: colorValueSchema.optional(),
  lineHeight: validatedLineHeightSchema.optional(),
  verticalAlign: z.enum(["top", "middle", "bottom"]).optional(),
  /** Paragraph alignment when not explicitly overridden. */
  align: textAlignmentSchema.optional(),
  /** Paragraph space-after when not explicitly overridden, in pixels. */
  spacingAfter: z
    .number()
    .finite()
    .nonnegative()
    .refine((value) => {
      const centipoints = Math.round(value * 75);
      return centipoints >= 0 && centipoints <= 158_400;
    }, "spacingAfter converted value must be within DrawingML range of 0 to 1584pt")
    .optional(),
  /** Character spacing in pixels. */
  letterSpacing: z
    .number()
    .finite()
    .refine((value) => {
      const points = Math.round(value * 75);
      return points >= -400_000 && points <= 400_000;
    }, "letterSpacing converted value must be within DrawingML range of -4000 to 4000pt")
    .optional(),
  textTransform: z.enum(["none", "uppercase"]).optional(),
  textDecoration: z.enum(["none", "underline", "line-through"]).optional(),
  textShadow: z.union([shadowSchema, z.literal("none")]).optional(),
  /** Whether the text box wraps normally or keeps each line on a single line. */
  textWrap: z.enum(["normal", "nowrap"]).optional(),
  /** Natural language (BCP 47 language tag, e.g. zh-CN, en-US). */
  lang: bcp47LangSchema.optional(),
});

export const textStyleShape = textStyleObjectSchema.shape;

/** Inheritable text style fields shared across text objects and containers. */
export const textStyleSchema = textStyleObjectSchema.readonly();
export type TextStyle = z.infer<typeof textStyleSchema>;

/** Deck-level theme input; major/minor fonts are simultaneously written to native PPTX theme. */
export const themeInputSchema = z
  .strictObject({
    name: z.string().optional(),
    colors: z
      .object({
        dk1: z.string().optional(),
        lt1: z.string().optional(),
        dk2: z.string().optional(),
        lt2: z.string().optional(),
        accent1: z.string().optional(),
        accent2: z.string().optional(),
        accent3: z.string().optional(),
        accent4: z.string().optional(),
        accent5: z.string().optional(),
        accent6: z.string().optional(),
        hlink: z.string().optional(),
        folHlink: z.string().optional(),
      })
      .catchall(themeColorSchema)
      .readonly()
      .optional(),
    fonts: z
      .object({
        major: fontTokenSchema.optional(),
        minor: fontTokenSchema.optional(),
      })
      .catchall(fontTokenSchema)
      .readonly()
      .optional(),
  })
  .readonly();
export type ThemeInput = z.infer<typeof themeInputSchema>;

export const themeSchema = z
  .strictObject({
    name: z.string(),
    colors: z
      .object({
        dk1: z.string(),
        lt1: z.string(),
        dk2: z.string(),
        lt2: z.string(),
        accent1: z.string(),
        accent2: z.string(),
        accent3: z.string(),
        accent4: z.string(),
        accent5: z.string(),
        accent6: z.string(),
        hlink: z.string(),
        folHlink: z.string(),
      })
      .catchall(themeColorSchema)
      .readonly(),
    fonts: z
      .object({
        major: fontTokenSchema,
        minor: fontTokenSchema,
      })
      .catchall(fontTokenSchema)
      .readonly(),
  })
  .readonly();
export type Theme = z.infer<typeof themeSchema>;

/** Local theme scopes on slides or containers can only override color and font tokens already present in the deck. */
export const themeScopeSchema = z
  .strictObject({
    colors: z.record(z.string(), themeColorSchema).readonly().optional(),
    fonts: z.record(z.string(), fontTokenSchema).readonly().optional(),
  })
  .readonly();
export type ThemeScope = z.infer<typeof themeScopeSchema>;

export const resolvedThemeFontSchema = z
  .strictObject({
    family: z.string(),
    font: fontSchema.optional(),
    themeFont: z.enum(["major", "minor"]).optional(),
  })
  .readonly();
export type ResolvedThemeFont = z.infer<typeof resolvedThemeFontSchema>;

/** Theme, local scopes, and accumulated container text defaults in a single build. */
export type ThemeContext = {
  readonly theme: Theme;
  readonly textStyle: TextStyle;
  with(scope?: ThemeScope, style?: TextStyle): ThemeContext;
  /** Replaces the accumulated container text defaults while retaining the same theme scope chain. */
  withTextStyle(style: TextStyle): ThemeContext;
  resolveColor(value: ColorValue): ResolvedColor;
  resolveFont(value: FontReference): ResolvedThemeFont;
};

const officeColors: Record<NativeColorSlot, string> = {
  dk1: "#000000",
  lt1: "#FFFFFF",
  dk2: "#1F497D",
  lt2: "#EEECE1",
  accent1: "#4F81BD",
  accent2: "#C0504D",
  accent3: "#9BBB59",
  accent4: "#8064A2",
  accent5: "#4BACC6",
  accent6: "#F79646",
  hlink: "#0000FF",
  folHlink: "#800080",
};

const contextRecords = new WeakMap<
  ThemeContext,
  {
    readonly parent?: ThemeContext;
    readonly scope?: ThemeScope;
  }
>();

export const textStyleKeys = new Set(Object.keys(textStyleObjectSchema.shape));

function textStyleErrorForIssue(
  issue: z.core.$ZodIssue | undefined,
  rawStyle: Readonly<Record<string, unknown>>,
): CapabilityError {
  const field = issue?.path[0];
  if (field === "fontFamily") {
    if (issue?.path[1] === "faces") {
      return new CapabilityError(
        "invalid-font-face",
        "Font.faces contains invalid weight, style, or resource",
      );
    }
    return new CapabilityError(
      "invalid-font-family",
      issue?.message ??
        "fontFamily must be a font value or non-empty family string",
    );
  }
  if (field === "fontSize") {
    const raw = rawStyle["fontSize"];
    const validLengthSyntax =
      typeof raw === "number" ||
      (typeof raw === "string" && textLengthPattern.test(raw));
    return new CapabilityError(
      validLengthSyntax ? "invalid-font-size" : "invalid-text-length",
      validLengthSyntax
        ? "fontSize must be greater than 0 and within DrawingML range of 1 to 4000pt"
        : "fontSize must be a positive pixel value or px/pt length",
    );
  }
  if (field === "fontWeight") {
    return new CapabilityError(
      "invalid-font-weight",
      "fontWeight must be normal, bold, or a hundreds integer between 100 and 900",
    );
  }
  if (field === "fontStyle") {
    return new CapabilityError(
      "invalid-font-style",
      "fontStyle must be normal or italic",
    );
  }
  if (field === "color") {
    return new CapabilityError("invalid-color", "Invalid color value format");
  }
  if (field === "lineHeight") {
    return new CapabilityError(
      "invalid-text-length",
      "lineHeight must be a positive pixel value or px/pt length within 0 to 1584pt",
    );
  }
  if (field === "verticalAlign") {
    return new CapabilityError(
      "invalid-vertical-align",
      "verticalAlign must be top, middle, or bottom",
    );
  }
  if (field === "align") {
    return new CapabilityError(
      "invalid-text-alignment",
      "align must be left, center, right, or justify",
    );
  }
  if (field === "spacingAfter") {
    return new CapabilityError(
      "invalid-paragraph-spacing",
      "spacingAfter must be a finite non-negative pixel number within DrawingML range of 0 to 1584pt",
    );
  }
  if (field === "letterSpacing") {
    return new CapabilityError(
      "invalid-letter-spacing",
      "letterSpacing converted value must be within DrawingML range of -4000 to 4000pt",
    );
  }
  if (field === "textTransform") {
    return new CapabilityError(
      "invalid-text-transform",
      "textTransform must be none or uppercase",
    );
  }
  if (field === "textDecoration") {
    return new CapabilityError(
      "invalid-text-decoration",
      "textDecoration must be none, underline, or line-through",
    );
  }
  if (field === "textWrap") {
    return new CapabilityError(
      "invalid-text-wrap",
      "textWrap must be normal or nowrap",
    );
  }
  if (field === "textShadow") {
    return new CapabilityError(
      "invalid-shadow",
      issue?.message ?? "textShadow must be a shadow object or 'none'",
    );
  }
  if (field === "lang") {
    return new CapabilityError(
      "invalid-text-lang",
      `Text language must be a valid BCP 47 language tag: ${String(rawStyle["lang"])}`,
    );
  }
  return new CapabilityError(
    "invalid-style",
    issue?.message ?? "Invalid text style",
  );
}

/**
 * Validates inheritable text style fields shared across text objects and containers.
 */
export function validateTextStyle(
  style: Readonly<Record<string, unknown>>,
  theme?: ThemeContext,
): TextStyle {
  const candidate: Record<string, unknown> = {};
  for (const [key, item] of Object.entries(style)) {
    if (textStyleKeys.has(key) && item !== undefined) {
      candidate[key] = item;
    }
  }
  const parsed = textStyleObjectSchema.safeParse(candidate);
  if (!parsed.success) {
    throw textStyleErrorForIssue(parsed.error.issues[0], candidate);
  }
  const validated = parsed.data;
  if (theme !== undefined) {
    if (typeof validated.fontFamily === "string") {
      theme.resolveFont(validated.fontFamily);
    }
    if (validated.color !== undefined) {
      theme.resolveColor(validated.color);
    }
    if (
      validated.textShadow !== undefined &&
      validated.textShadow !== "none" &&
      validated.textShadow.color !== undefined
    ) {
      theme.resolveColor(validated.textShadow.color);
    }
  }
  return style;
}

/**
 * Asserts that a value is a valid BCP 47 language tag.
 * Rejects non-string values, empty or whitespace strings, and malformed tags.
 */
export function assertValidLang(value: unknown): string {
  const parsed = bcp47LangSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-text-lang",
      `Text language must be a valid BCP 47 language tag: ${String(value)}`,
    );
  }
  return parsed.data;
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Extract inheritable text fields from container style containing layout fields. */
export function textStyleFrom(value: unknown, theme?: ThemeContext): TextStyle {
  if (value === undefined) return {};
  if (!record(value))
    throw new CapabilityError("invalid-style", "style must be an object");
  const extracted: Record<string, unknown> = {};
  for (const [key, item] of Object.entries(value)) {
    if (textStyleKeys.has(key) && item !== undefined) {
      extracted[key] = item;
    }
  }
  return validateTextStyle(extracted, theme);
}

function nativeSlot(value: string): value is NativeColorSlot {
  return (nativeColorSlots as readonly string[]).includes(value);
}

export function isFont(value: unknown): value is Font {
  return record(value) && typeof value["family"] === "string";
}

function normalizeTheme(input: ThemeInput = {}): Theme {
  if (!record(input)) throw new TypeError("Theme must be an object");
  const parsed = themeInputSchema.safeParse(input);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    throw new CapabilityError(
      "invalid-theme",
      issue?.message ?? "Invalid theme declaration",
    );
  }
  const validInput = parsed.data;
  // Theme values are content inputs: isolate nested shades, font declarations, and font bytes.
  const colors = snapshotValue({
    ...officeColors,
    ...validInput.colors,
  }) as Theme["colors"];
  const fonts = snapshotValue({
    ...validInput.fonts,
    major: validInput.fonts?.major ?? "Aptos Display",
    minor: validInput.fonts?.minor ?? "Aptos",
  }) as Theme["fonts"];
  return Object.freeze({
    name: validInput.name ?? "Office Theme",
    colors,
    fonts,
  });
}

/** Fill unsupplied theme items using initial Office palette and fonts. */
export function defineTheme(input: ThemeInput = {}): Theme {
  return normalizeTheme(input);
}

export const defaultTheme = defineTheme();

/** Generate darker variant in PowerPoint native palette. */
export function darker(slot: NativeColorSlot, amount: number): ThemeShade {
  return { slot, lumMod: 1 - amount };
}

/** Generate lighter variant in PowerPoint native palette. */
export function lighter(slot: NativeColorSlot, amount: number): ThemeShade {
  return { slot, lumMod: 1 - amount, lumOff: amount };
}

export function isNativeColorSlot(value: string): value is NativeColorSlot {
  return nativeSlot(value);
}

function scopeValue(
  context: ThemeContext,
  name: string,
): { readonly value: ThemeColor; readonly parent?: ThemeContext } | undefined {
  let current: ThemeContext | undefined = context;
  while (current !== undefined) {
    const scope = contextRecords.get(current)?.scope;
    if (scope?.colors !== undefined && Object.hasOwn(scope.colors, name)) {
      return {
        value: scope.colors[name]!,
        ...(contextRecords.get(current)?.parent === undefined
          ? {}
          : { parent: contextRecords.get(current)!.parent! }),
      };
    }
    current = contextRecords.get(current)?.parent;
  }
  return undefined;
}

function scopeFontValue(
  context: ThemeContext,
  name: string,
): { readonly value: FontToken; readonly parent?: ThemeContext } | undefined {
  let current: ThemeContext | undefined = context;
  while (current !== undefined) {
    const scope = contextRecords.get(current)?.scope;
    if (scope?.fonts !== undefined && Object.hasOwn(scope.fonts, name)) {
      return {
        value: scope.fonts[name]!,
        ...(contextRecords.get(current)?.parent === undefined
          ? {}
          : { parent: contextRecords.get(current)!.parent! }),
      };
    }
    current = contextRecords.get(current)?.parent;
  }
  return undefined;
}

function shadeHex(hex: string, lumMod: number, lumOff: number): string {
  const [r, g, b] = [0, 2, 4].map(
    (start) => Number.parseInt(hex.slice(start, start + 2), 16) / 255,
  ) as [number, number, number];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  const delta = max - min;
  const saturation =
    delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  const hue =
    delta === 0
      ? 0
      : max === r
        ? ((g - b) / delta + 6) % 6
        : max === g
          ? (b - r) / delta + 2
          : (r - g) / delta + 4;
  const shaded = Math.min(1, Math.max(0, lightness * lumMod + lumOff));
  const chroma = (1 - Math.abs(2 * shaded - 1)) * saturation;
  const x = chroma * (1 - Math.abs((hue % 2) - 1));
  const m = shaded - chroma / 2;
  const [r1, g1, b1] =
    hue < 1
      ? [chroma, x, 0]
      : hue < 2
        ? [x, chroma, 0]
        : hue < 3
          ? [0, chroma, x]
          : hue < 4
            ? [0, x, chroma]
            : hue < 5
              ? [x, 0, chroma]
              : [chroma, 0, x];
  return [r1, g1, b1]
    .map((channel) =>
      Math.round((channel + m) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

function checkedFraction(value: unknown, label: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < 0 ||
    value > 1
  ) {
    throw new CapabilityError(
      "invalid-color",
      `${label} must be a number between 0 and 1`,
    );
  }
  return value;
}

function resolveColorName(
  raw: string,
  context: ThemeContext,
  depth: number,
): ResolvedColor {
  if (depth > 16) {
    throw new CapabilityError(
      "invalid-color",
      `Circular color token reference: ${raw}`,
    );
  }
  const value = raw.trim();
  if (value.startsWith("#")) {
    const digits = value.slice(1);
    const expanded =
      digits.length === 3 || digits.length === 4
        ? digits
            .split("")
            .map((digit) => digit + digit)
            .join("")
        : digits;
    if (!/^[0-9a-f]{6}([0-9a-f]{2})?$/iu.test(expanded)) {
      throw new CapabilityError("invalid-color", `Invalid color value: ${raw}`);
    }
    const alpha =
      expanded.length === 8
        ? Number.parseInt(expanded.slice(6), 16) / 255
        : undefined;
    return {
      hex: expanded.slice(0, 6).toUpperCase(),
      ...(alpha === undefined || alpha === 1 ? {} : { alpha }),
    };
  }
  if (nativeSlot(value)) {
    const configured = context.theme.colors[value];
    const resolved = resolveColorName(
      configured,
      rootContext(context),
      depth + 1,
    );
    return { ...resolved, slot: value };
  }
  const local = scopeValue(context, value);
  if (local !== undefined) {
    return resolveColorValue(
      local.value,
      local.parent ?? rootContext(context),
      depth + 1,
    );
  }
  const configured = context.theme.colors[value];
  if (configured === undefined) {
    throw new CapabilityError(
      "invalid-color",
      `Unknown color or theme token: ${raw}`,
    );
  }
  // Deck token aliases are resolved at the point of use so they see overrides from the nearest scope.
  return resolveColorValue(configured, context, depth + 1);
}

function rootContext(context: ThemeContext): ThemeContext {
  let current = context;
  while (contextRecords.get(current)?.parent !== undefined) {
    current = contextRecords.get(current)!.parent!;
  }
  return current;
}

function resolveColorValue(
  value: ThemeColor,
  context: ThemeContext,
  depth: number,
): ResolvedColor {
  if (typeof value === "string") return resolveColorName(value, context, depth);
  if (
    !record(value) ||
    !nativeSlot(String(value["slot"])) ||
    typeof value["lumMod"] !== "number" ||
    !Number.isFinite(value["lumMod"])
  ) {
    throw new CapabilityError("invalid-color", "Invalid theme shade value");
  }
  const lumMod = checkedFraction(value["lumMod"], "lumMod");
  const lumOff =
    value["lumOff"] === undefined
      ? undefined
      : checkedFraction(value["lumOff"], "lumOff");
  const slot = value["slot"];
  const base = resolveColorName(slot, rootContext(context), depth + 1);
  return {
    hex: shadeHex(base.hex, lumMod, lumOff ?? 0),
    slot,
    lumMod,
    ...(lumOff === undefined ? {} : { lumOff }),
  };
}

function createContext(
  theme: Theme,
  parent: ThemeContext | undefined,
  scope: ThemeScope | undefined,
  textStyle: TextStyle,
): ThemeContext {
  let context: ThemeContext;
  context = Object.freeze({
    theme,
    textStyle: Object.freeze(textStyle),
    with(nextScope?: ThemeScope, style?: TextStyle) {
      let checkedScope: ThemeScope | undefined;
      if (nextScope !== undefined) {
        const parsedScope = themeScopeSchema.safeParse(nextScope);
        if (!parsedScope.success) {
          throw new CapabilityError(
            "invalid-theme-scope",
            parsedScope.error.issues[0]?.message ??
              "Local theme scope is invalid",
          );
        }
        const colors = parsedScope.data.colors;
        for (const name of Object.keys(colors ?? {})) {
          if (nativeSlot(name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `Local theme scope cannot redefine native color slot ${name}`,
            );
          }
          if (!Object.hasOwn(theme.colors, name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `Theme has no color token ${name}`,
            );
          }
        }
        const fonts = parsedScope.data.fonts;
        for (const name of Object.keys(fonts ?? {})) {
          if (!Object.hasOwn(theme.fonts, name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `Theme has no font token ${name}`,
            );
          }
        }
        checkedScope = Object.freeze({
          ...(colors === undefined
            ? {}
            : { colors: Object.freeze({ ...colors }) }),
          ...(fonts === undefined
            ? {}
            : { fonts: Object.freeze({ ...fonts }) }),
        });
      }
      const candidateContext = createContext(
        theme,
        context,
        checkedScope,
        context.textStyle,
      );
      const validated =
        style === undefined ? {} : textStyleFrom(style, candidateContext);
      const nextStyle = { ...context.textStyle, ...validated };
      return createContext(theme, context, checkedScope, nextStyle);
    },
    withTextStyle(style) {
      return createContext(
        theme,
        context,
        undefined,
        textStyleFrom(style, context),
      );
    },
    resolveColor(value: ColorValue) {
      let raw: string | ThemeShade;
      let alpha: number | undefined;
      if (typeof value === "string" || (record(value) && "slot" in value)) {
        raw = value;
      } else if (record(value) && "color" in value) {
        raw = value["color"];
        alpha =
          value["alpha"] === undefined
            ? undefined
            : checkedFraction(value["alpha"], "color alpha");
      } else {
        throw new CapabilityError(
          "invalid-color",
          "Invalid color value format",
        );
      }
      const resolved =
        typeof raw === "string"
          ? resolveColorName(raw, context, 0)
          : resolveColorValue(raw, context, 0);
      if (alpha === undefined) return resolved;
      return {
        ...resolved,
        alpha: (resolved.alpha ?? 1) * alpha,
      };
    },
    resolveFont(value: FontReference) {
      if (isFont(value)) return { family: value.family, font: value };
      if (typeof value !== "string" || value.trim() === "") {
        throw new CapabilityError(
          "invalid-font-family",
          "fontFamily cannot be empty",
        );
      }
      const local = scopeFontValue(context, value);
      const token =
        local?.value ??
        (Object.hasOwn(theme.fonts, value) ? theme.fonts[value]! : value);
      const themeFontName: ResolvedThemeFont["themeFont"] =
        value === "major" || value === "minor" ? value : undefined;
      const globalThemeFont =
        themeFontName === undefined ? undefined : theme.fonts[themeFontName];
      const globalFamily = isFont(globalThemeFont)
        ? globalThemeFont.family
        : globalThemeFont;
      if (isFont(token)) {
        return {
          family: token.family,
          font: token,
          ...(themeFontName === undefined || token.family !== globalFamily
            ? {}
            : { themeFont: themeFontName }),
        };
      }
      return {
        family: token,
        ...(themeFontName === undefined || token !== globalFamily
          ? {}
          : { themeFont: themeFontName }),
      };
    },
  });
  contextRecords.set(context, {
    ...(parent === undefined ? {} : { parent }),
    ...(scope === undefined ? {} : { scope }),
  });
  return context;
}

/** Create a theme chain unique to a single build; local scopes do not mutate the shared theme object. */
export function createThemeContext(
  input?: ThemeInput | Theme,
  style?: TextStyle,
): ThemeContext {
  if (style !== undefined) {
    if (!record(style)) {
      throw new CapabilityError("invalid-style", "style must be an object");
    }
    const unknown = Object.keys(style).find((key) => !textStyleKeys.has(key));
    if (unknown !== undefined) {
      throw new CapabilityError(
        "invalid-style",
        `Deck.style contains unsupported text property ${unknown}`,
      );
    }
  }
  const theme = normalizeTheme(input);
  const initialContext = createContext(theme, undefined, undefined, {});
  return createContext(
    theme,
    undefined,
    undefined,
    textStyleFrom(style, initialContext),
  );
}
