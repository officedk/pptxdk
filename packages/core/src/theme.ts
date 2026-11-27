import { z } from "zod";

import { CapabilityError } from "#/diagnostics.ts";
import { snapshotValue } from "#/input-values.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type FontFaceStyle = "normal" | "italic";
const rawFontFaceStyleSchema = z.enum(["normal", "italic"]);
export const fontFaceStyleSchema: z.ZodType<FontFaceStyle> =
  rawFontFaceStyleSchema satisfies z.ZodType<
    AssertExact<
      FontFaceStyle,
      z.infer<typeof rawFontFaceStyleSchema>
    > extends true
      ? FontFaceStyle
      : never
  >;

const fontWeightNumberSchema = z
  .number()
  .int("fontWeight must be normal, bold, or an integer between 1 and 1000")
  .min(1, "fontWeight must be normal, bold, or an integer between 1 and 1000")
  .max(
    1000,
    "fontWeight must be normal, bold, or an integer between 1 and 1000",
  );

const fontWeightSchema = z.union([
  z.literal("normal"),
  z.literal("bold"),
  fontWeightNumberSchema,
]);

export const uint8ArraySchema = z.instanceof(
  Uint8Array,
) as unknown as z.ZodType<Uint8Array>;

/**
 * Single font face resource declaration paired with a {@link Font}.
 *
 * When `data` is a static single-font binary and `weight` or `style` is omitted,
 * the SDK infers the effective value from `OS/2.usWeightClass` and italic metadata
 * (`OS/2.fsSelection.italic`, `OS/2.fsSelection.oblique`, or non-zero `post.italicAngle`).
 * Explicit values that disagree with file metadata fail validation during build.
 * When `embedFonts` is enabled, the owning `Font.family` must match the file's
 * native Unicode legacy family name (`Name ID 1`).
 */
export type FontFace = {
  /** OpenType numeric weight (`1..1000`); defaults to `400` when `data` is omitted. */
  readonly weight?: number | undefined;
  /** Glyph style (`"normal"` or `"italic"`); defaults to `"normal"` when `data` is omitted. */
  readonly style?: FontFaceStyle | undefined;
  /** Static single-font bytes (`TTF`, `OTF`, `WOFF`, or `WOFF2`); variable fonts (`fvar`) are rejected. */
  readonly data?: Uint8Array | undefined;
};

const rawFontFaceSchema = z
  .strictObject({
    weight: fontWeightNumberSchema.optional(),
    style: rawFontFaceStyleSchema.optional(),
    data: uint8ArraySchema
      .refine((bytes) => bytes.byteLength > 0, "Font face data cannot be empty")
      .optional(),
  })
  .readonly();
/** Font files provided by the caller can be used for measurement or embedded according to build options. */
export const fontFaceSchema: z.ZodType<FontFace> =
  rawFontFaceSchema satisfies z.ZodType<
    AssertExact<FontFace, z.infer<typeof rawFontFaceSchema>> extends true
      ? FontFace
      : never
  >;

const nonEmptyTrimmedStringSchema = z
  .string()
  .refine((value) => value.trim() !== "", "Value cannot be empty");

/** Font family and explicit measurement resources referenceable across deck themes and components. */
export type Font = {
  readonly family: string;
  readonly faces?: readonly FontFace[] | undefined;
  readonly substitute?: boolean | undefined;
  readonly embeddable?: boolean | undefined;
};

const rawFontSchema = z
  .strictObject({
    family: nonEmptyTrimmedStringSchema,
    faces: z.array(rawFontFaceSchema).readonly().optional(),
    substitute: z.boolean().optional(),
    embeddable: z.boolean().optional(),
  })
  .readonly();
/** Font family and explicit measurement resources referenceable across deck themes and components. */
export const fontSchema: z.ZodType<Font> = rawFontSchema satisfies z.ZodType<
  AssertExact<Font, z.infer<typeof rawFontSchema>> extends true ? Font : never
>;

export type FontOptions = {
  readonly substitute?: boolean | undefined;
  readonly embeddable?: boolean | undefined;
};

const rawFontOptionsSchema = z
  .strictObject({
    substitute: z.boolean().optional(),
    embeddable: z.boolean().optional(),
  })
  .readonly();
export const fontOptionsSchema: z.ZodType<FontOptions> =
  rawFontOptionsSchema satisfies z.ZodType<
    AssertExact<FontOptions, z.infer<typeof rawFontOptionsSchema>> extends true
      ? FontOptions
      : never
  >;

export type FontReference = string | Font;
export type FontToken = FontReference;

const rawFontReferenceSchema = z.union([
  nonEmptyTrimmedStringSchema,
  rawFontSchema,
]);
export const fontReferenceSchema: z.ZodType<FontReference> =
  rawFontReferenceSchema satisfies z.ZodType<
    AssertExact<
      FontReference,
      z.infer<typeof rawFontReferenceSchema>
    > extends true
      ? FontReference
      : never
  >;

export const fontTokenSchema: z.ZodType<FontToken> = fontReferenceSchema;

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

/**
 * Font size for text runs and paragraphs.
 *
 * Positive numbers default to pixels (`px`) converted at `96 dpi` (`1pt = 4/3 px`, so `16` or
 * `"16px"` equals `"12pt"`). Explicit `"pt"` strings such as `"12pt"` are also accepted.
 * Converted sizes must fall within `1pt` to `4000pt`.
 */
export type FontSize = number | `${number}px` | `${number}pt`;

const rawFontSizeSchema = z.union([
  z.number().finite().positive(),
  textLengthStringSchema.refine(
    (value) => textLengthToPixels(value) > 0,
    "Text length must be greater than 0",
  ),
]);
export const fontSizeSchema: z.ZodType<FontSize> =
  rawFontSizeSchema satisfies z.ZodType<
    AssertExact<FontSize, z.infer<typeof rawFontSizeSchema>> extends true
      ? FontSize
      : never
  >;

const validatedFontSizeSchema = rawFontSizeSchema.refine((value) => {
  const serialized = Math.round(fontSizeToPoints(value) * 100);
  return serialized >= 100 && serialized <= 400_000;
}, "fontSize converted value must be within DrawingML range of 1 to 4000pt");

const validatedLineHeightSchema = rawFontSizeSchema.refine((value) => {
  const centipoints = Math.round(textLengthToPixels(value) * 75);
  return centipoints >= 0 && centipoints <= 158_400;
}, "lineHeight converted value must be within DrawingML range of 0 to 1584pt");

export type TextAlignment = "left" | "center" | "right" | "justify";
const rawTextAlignmentSchema = z.enum(["left", "center", "right", "justify"]);
export const textAlignmentSchema: z.ZodType<TextAlignment> =
  rawTextAlignmentSchema satisfies z.ZodType<
    AssertExact<
      TextAlignment,
      z.infer<typeof rawTextAlignmentSchema>
    > extends true
      ? TextAlignment
      : never
  >;

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

export type NativeColorSlot = (typeof nativeColorSlots)[number];
const rawNativeColorSlotSchema = z.enum(nativeColorSlots);
export const nativeColorSlotSchema: z.ZodType<NativeColorSlot> =
  rawNativeColorSlotSchema satisfies z.ZodType<
    AssertExact<
      NativeColorSlot,
      z.infer<typeof rawNativeColorSlotSchema>
    > extends true
      ? NativeColorSlot
      : never
  >;

export const schemeColorSlots = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export type SchemeColorSlot = (typeof schemeColorSlots)[number];
const rawSchemeColorSlotSchema = z.enum(schemeColorSlots);
export const schemeColorSlotSchema: z.ZodType<SchemeColorSlot> =
  rawSchemeColorSlotSchema satisfies z.ZodType<
    AssertExact<
      SchemeColorSlot,
      z.infer<typeof rawSchemeColorSlotSchema>
    > extends true
      ? SchemeColorSlot
      : never
  >;

export const schemeColorNames = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export type SchemeColorName = (typeof schemeColorNames)[number];
const rawSchemeColorNameSchema = z.enum(schemeColorNames);
export const schemeColorNameSchema: z.ZodType<SchemeColorName> =
  rawSchemeColorNameSchema satisfies z.ZodType<
    AssertExact<
      SchemeColorName,
      z.infer<typeof rawSchemeColorNameSchema>
    > extends true
      ? SchemeColorName
      : never
  >;

/** Native `<p:clrMap>` / `<a:overrideClrMapping>` mapping the 12 scheme slots to 12 theme palette slots. */
export type ColorMap = {
  readonly bg1: NativeColorSlot;
  readonly tx1: NativeColorSlot;
  readonly bg2: NativeColorSlot;
  readonly tx2: NativeColorSlot;
  readonly accent1: NativeColorSlot;
  readonly accent2: NativeColorSlot;
  readonly accent3: NativeColorSlot;
  readonly accent4: NativeColorSlot;
  readonly accent5: NativeColorSlot;
  readonly accent6: NativeColorSlot;
  readonly hlink: NativeColorSlot;
  readonly folHlink: NativeColorSlot;
};

const rawColorMapSchema = z
  .strictObject({
    bg1: rawNativeColorSlotSchema,
    tx1: rawNativeColorSlotSchema,
    bg2: rawNativeColorSlotSchema,
    tx2: rawNativeColorSlotSchema,
    accent1: rawNativeColorSlotSchema,
    accent2: rawNativeColorSlotSchema,
    accent3: rawNativeColorSlotSchema,
    accent4: rawNativeColorSlotSchema,
    accent5: rawNativeColorSlotSchema,
    accent6: rawNativeColorSlotSchema,
    hlink: rawNativeColorSlotSchema,
    folHlink: rawNativeColorSlotSchema,
  })
  .readonly();
/** Native `<p:clrMap>` / `<a:overrideClrMapping>` mapping the 12 scheme slots to 12 theme palette slots. */
export const colorMapSchema: z.ZodType<ColorMap> =
  rawColorMapSchema satisfies z.ZodType<
    AssertExact<ColorMap, z.infer<typeof rawColorMapSchema>> extends true
      ? ColorMap
      : never
  >;

export const defaultColorMap: ColorMap = Object.freeze({
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
});

export type ColorMapOverride = "master" | ColorMap;
const rawColorMapOverrideSchema = z.union([
  z.literal("master"),
  rawColorMapSchema,
]);
export const colorMapOverrideSchema: z.ZodType<ColorMapOverride> =
  rawColorMapOverrideSchema satisfies z.ZodType<
    AssertExact<
      ColorMapOverride,
      z.infer<typeof rawColorMapOverrideSchema>
    > extends true
      ? ColorMapOverride
      : never
  >;

/** Native theme color represented with DrawingML luminance modulation. */
export type ThemeShade = {
  readonly slot: SchemeColorName;
  readonly lumMod: number;
  readonly lumOff?: number | undefined;
};

const rawThemeShadeSchema = z
  .strictObject({
    slot: rawSchemeColorNameSchema,
    lumMod: z.number().finite().min(0).max(1),
    lumOff: z.number().finite().min(0).max(1).optional(),
  })
  .readonly();
/** Native theme color represented with DrawingML luminance modulation. */
export const themeShadeSchema: z.ZodType<ThemeShade> =
  rawThemeShadeSchema satisfies z.ZodType<
    AssertExact<ThemeShade, z.infer<typeof rawThemeShadeSchema>> extends true
      ? ThemeShade
      : never
  >;

export type ThemeColor = string | ThemeShade;
const rawThemeColorSchema = z.union([z.string(), rawThemeShadeSchema]);
export const themeColorSchema: z.ZodType<ThemeColor> =
  rawThemeColorSchema satisfies z.ZodType<
    AssertExact<ThemeColor, z.infer<typeof rawThemeColorSchema>> extends true
      ? ThemeColor
      : never
  >;

/** Color entry can directly provide color, theme slot, or theme token, optionally with alpha. */
export type ColorValue =
  | string
  | ThemeShade
  | {
      readonly color: ThemeColor;
      readonly alpha?: number | undefined;
    };

const rawColorValueSchema = z.union([
  z.string(),
  rawThemeShadeSchema,
  z
    .strictObject({
      color: rawThemeColorSchema,
      alpha: z.number().finite().min(0).max(1).optional(),
    })
    .readonly(),
]);
/** Color entry can directly provide color, theme slot, or theme token, optionally with alpha. */
export const colorValueSchema: z.ZodType<ColorValue> =
  rawColorValueSchema satisfies z.ZodType<
    AssertExact<ColorValue, z.infer<typeof rawColorValueSchema>> extends true
      ? ColorValue
      : never
  >;

/** Final theme value applied to a color entry. */
export type ResolvedColor = {
  readonly hex: string;
  readonly alpha?: number | undefined;
  readonly slot?: NativeColorSlot | undefined;
  readonly scheme?: SchemeColorName | undefined;
  readonly lumMod?: number | undefined;
  readonly lumOff?: number | undefined;
};

const rawResolvedColorSchema = z
  .strictObject({
    hex: z.string(),
    alpha: z.number().finite().min(0).max(1).optional(),
    slot: rawNativeColorSlotSchema.optional(),
    scheme: rawSchemeColorNameSchema.optional(),
    lumMod: z.number().finite().min(0).max(1).optional(),
    lumOff: z.number().finite().min(0).max(1).optional(),
  })
  .readonly();
/** Final theme value applied to a color entry. */
export const resolvedColorSchema: z.ZodType<ResolvedColor> =
  rawResolvedColorSchema satisfies z.ZodType<
    AssertExact<
      ResolvedColor,
      z.infer<typeof rawResolvedColorSchema>
    > extends true
      ? ResolvedColor
      : never
  >;

export type GradientStop = {
  readonly color: ColorValue;
  /** Normalized stop position in the closed interval `[0, 1]` (`0` = `0%` start, `1` = `100%` end). */
  readonly position: number;
};

const rawGradientStopSchema = z
  .strictObject({
    color: rawColorValueSchema,
    position: z.number().finite(),
  })
  .readonly();
export const gradientStopSchema: z.ZodType<GradientStop> =
  rawGradientStopSchema satisfies z.ZodType<
    AssertExact<
      GradientStop,
      z.infer<typeof rawGradientStopSchema>
    > extends true
      ? GradientStop
      : never
  >;

export type LinearGradient = {
  readonly type: "linear-gradient";
  readonly angle: number;
  readonly stops: readonly GradientStop[];
};

const rawLinearGradientSchema = z
  .strictObject({
    type: z.literal("linear-gradient"),
    angle: z.number().finite(),
    stops: z.array(rawGradientStopSchema).readonly(),
  })
  .readonly();
export const linearGradientSchema: z.ZodType<LinearGradient> =
  rawLinearGradientSchema satisfies z.ZodType<
    AssertExact<
      LinearGradient,
      z.infer<typeof rawLinearGradientSchema>
    > extends true
      ? LinearGradient
      : never
  >;

export type RadialGradient = {
  readonly type: "radial-gradient";
  readonly stops: readonly GradientStop[];
};

const rawRadialGradientSchema = z
  .strictObject({
    type: z.literal("radial-gradient"),
    stops: z.array(rawGradientStopSchema).readonly(),
  })
  .readonly();
export const radialGradientSchema: z.ZodType<RadialGradient> =
  rawRadialGradientSchema satisfies z.ZodType<
    AssertExact<
      RadialGradient,
      z.infer<typeof rawRadialGradientSchema>
    > extends true
      ? RadialGradient
      : never
  >;

export type ImageBackgroundFit = "fill" | "cover" | "contain";
const rawImageBackgroundFitSchema = z.enum(["fill", "cover", "contain"]);
export const imageBackgroundFitSchema: z.ZodType<ImageBackgroundFit> =
  rawImageBackgroundFitSchema satisfies z.ZodType<
    AssertExact<
      ImageBackgroundFit,
      z.infer<typeof rawImageBackgroundFitSchema>
    > extends true
      ? ImageBackgroundFit
      : never
  >;

const percentPositionPattern = /^(-?(?:\d+(?:\.\d*)?|\.\d+))%$/u;
const positionPercentSchema = z
  .string()
  .regex(percentPositionPattern) as unknown as z.ZodType<`${number}%`>;
const positionCoordinateSchema = z.union([
  z.number().finite(),
  positionPercentSchema,
]);

export type ImageBackgroundPositionKeyword =
  | "center"
  | "top"
  | "bottom"
  | "left"
  | "right"
  | "top left"
  | "top right"
  | "bottom left"
  | "bottom right";

const rawImageBackgroundPositionKeywordSchema = z.enum([
  "center",
  "top",
  "bottom",
  "left",
  "right",
  "top left",
  "top right",
  "bottom left",
  "bottom right",
]);
export const imageBackgroundPositionKeywordSchema: z.ZodType<ImageBackgroundPositionKeyword> =
  rawImageBackgroundPositionKeywordSchema satisfies z.ZodType<
    AssertExact<
      ImageBackgroundPositionKeyword,
      z.infer<typeof rawImageBackgroundPositionKeywordSchema>
    > extends true
      ? ImageBackgroundPositionKeyword
      : never
  >;

export type ImageBackgroundPosition =
  | ImageBackgroundPositionKeyword
  | {
      readonly x?: number | `${number}%` | undefined;
      readonly y?: number | `${number}%` | undefined;
    };

const rawImageBackgroundPositionSchema = z.union([
  rawImageBackgroundPositionKeywordSchema,
  z
    .strictObject({
      x: positionCoordinateSchema.optional(),
      y: positionCoordinateSchema.optional(),
    })
    .readonly(),
]);
export const imageBackgroundPositionSchema: z.ZodType<ImageBackgroundPosition> =
  rawImageBackgroundPositionSchema satisfies z.ZodType<
    AssertExact<
      ImageBackgroundPosition,
      z.infer<typeof rawImageBackgroundPositionSchema>
    > extends true
      ? ImageBackgroundPosition
      : never
  >;

export type ExternalImageFillSource = {
  readonly externalUrl: string;
};

const rawExternalImageFillSourceSchema = z
  .strictObject({
    externalUrl: z.string().min(1),
  })
  .readonly();
export const externalImageFillSourceSchema: z.ZodType<ExternalImageFillSource> =
  rawExternalImageFillSourceSchema satisfies z.ZodType<
    AssertExact<
      ExternalImageFillSource,
      z.infer<typeof rawExternalImageFillSourceSchema>
    > extends true
      ? ExternalImageFillSource
      : never
  >;

export type ImageFillSource =
  | string
  | URL
  | Uint8Array
  | ExternalImageFillSource;

const rawImageFillSourceSchema = z.union([
  z.string(),
  z.instanceof(URL),
  uint8ArraySchema,
  rawExternalImageFillSourceSchema,
]);
export const imageFillSourceSchema: z.ZodType<ImageFillSource> =
  rawImageFillSourceSchema satisfies z.ZodType<
    AssertExact<
      ImageFillSource,
      z.infer<typeof rawImageFillSourceSchema>
    > extends true
      ? ImageFillSource
      : never
  >;

export type ImageBackground = {
  readonly image: ImageFillSource;
  readonly fit?: ImageBackgroundFit | undefined;
  readonly position?: ImageBackgroundPosition | undefined;
  readonly opacity?: number | undefined;
};

const rawImageBackgroundSchema = z
  .strictObject({
    image: rawImageFillSourceSchema,
    fit: rawImageBackgroundFitSchema.optional(),
    position: rawImageBackgroundPositionSchema.optional(),
    opacity: z.number().finite().min(0).max(1).optional(),
  })
  .readonly();
export const imageBackgroundSchema: z.ZodType<ImageBackground> =
  rawImageBackgroundSchema satisfies z.ZodType<
    AssertExact<
      ImageBackground,
      z.infer<typeof rawImageBackgroundSchema>
    > extends true
      ? ImageBackground
      : never
  >;

/** Fill value (`"none"`, color token/hex, gradient, or image background). */
export type BackgroundValue =
  | ColorValue
  | LinearGradient
  | RadialGradient
  | ImageBackground;

const rawBackgroundValueSchema = z.union([
  z.literal("none"),
  rawColorValueSchema,
  rawLinearGradientSchema,
  rawRadialGradientSchema,
  rawImageBackgroundSchema,
]);
export const backgroundValueSchema: z.ZodType<BackgroundValue> =
  rawBackgroundValueSchema satisfies z.ZodType<
    AssertExact<
      BackgroundValue,
      z.infer<typeof rawBackgroundValueSchema>
    > extends true
      ? BackgroundValue
      : never
  >;

export type BorderStyle = "solid" | "dashed" | "dotted";
const rawBorderStyleSchema = z.enum(["solid", "dashed", "dotted"]);
export const borderStyleSchema: z.ZodType<BorderStyle> =
  rawBorderStyleSchema satisfies z.ZodType<
    AssertExact<BorderStyle, z.infer<typeof rawBorderStyleSchema>> extends true
      ? BorderStyle
      : never
  >;

export type BorderObject = {
  readonly color: ColorValue;
  readonly width?: number | undefined;
  readonly style?: BorderStyle | undefined;
};

const borderObjectRawSchema = z.strictObject({
  color: rawColorValueSchema,
  width: z.number().finite().nonnegative().optional(),
  style: rawBorderStyleSchema.optional(),
});
export const borderObjectShape = borderObjectRawSchema.shape;
const rawBorderObjectSchema = borderObjectRawSchema.readonly();
export const borderObjectSchema: z.ZodType<BorderObject> =
  rawBorderObjectSchema satisfies z.ZodType<
    AssertExact<
      BorderObject,
      z.infer<typeof rawBorderObjectSchema>
    > extends true
      ? BorderObject
      : never
  >;

export type BorderValue = "none" | BorderObject;
const rawBorderValueSchema = z.union([
  z.literal("none"),
  rawBorderObjectSchema,
]);
export const borderValueSchema: z.ZodType<BorderValue> =
  rawBorderValueSchema satisfies z.ZodType<
    AssertExact<BorderValue, z.infer<typeof rawBorderValueSchema>> extends true
      ? BorderValue
      : never
  >;

export type Shadow = {
  readonly color?: ColorValue | undefined;
  readonly blur?: number | undefined;
  readonly distance?: number | undefined;
  readonly angle?: number | undefined;
};

const rawShadowSchema = z
  .strictObject({
    color: rawColorValueSchema.optional(),
    blur: z.number().finite().nonnegative().optional(),
    distance: z.number().finite().nonnegative().optional(),
    angle: z.number().finite().optional(),
  })
  .readonly();
export const shadowSchema: z.ZodType<Shadow> =
  rawShadowSchema satisfies z.ZodType<
    AssertExact<Shadow, z.infer<typeof rawShadowSchema>> extends true
      ? Shadow
      : never
  >;

export type BoxShadowValue = "none" | Shadow;
const rawBoxShadowValueSchema = z.union([z.literal("none"), rawShadowSchema]);
export const boxShadowValueSchema: z.ZodType<BoxShadowValue> =
  rawBoxShadowValueSchema satisfies z.ZodType<
    AssertExact<
      BoxShadowValue,
      z.infer<typeof rawBoxShadowValueSchema>
    > extends true
      ? BoxShadowValue
      : never
  >;

/** Visual fields shared across native DrawingML shapes, images, text boxes, and visible Boxes. */
export type VisualStyle = {
  readonly background?: BackgroundValue | undefined;
  readonly border?: BorderValue | undefined;
  readonly borderRadius?: number | undefined;
  readonly opacity?: number | undefined;
  readonly boxShadow?: BoxShadowValue | undefined;
  readonly softEdge?: number | undefined;
  readonly rotate?: number | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
};

const visualStyleObjectSchema = z.strictObject({
  background: rawBackgroundValueSchema.optional(),
  border: rawBorderValueSchema.optional(),
  borderRadius: z.number().finite().optional(),
  opacity: z.number().finite().optional(),
  boxShadow: rawBoxShadowValueSchema.optional(),
  softEdge: z.number().finite().optional(),
  rotate: z.number().finite().optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
});
export const visualStyleShape = visualStyleObjectSchema.shape;
const rawVisualStyleSchema = visualStyleObjectSchema.readonly();
/** Visual fields shared across native DrawingML shapes, images, text boxes, and visible Boxes. */
export const visualStyleSchema: z.ZodType<VisualStyle> =
  rawVisualStyleSchema satisfies z.ZodType<
    AssertExact<VisualStyle, z.infer<typeof rawVisualStyleSchema>> extends true
      ? VisualStyle
      : never
  >;

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

/** Inheritable text style fields shared across text objects and containers. */
export type TextStyle = {
  readonly fontFamily?: FontReference | undefined;
  readonly fontSize?: FontSize | undefined;
  readonly fontWeight?: "normal" | "bold" | number | undefined;
  readonly fontStyle?: FontFaceStyle | undefined;
  readonly color?: ColorValue | undefined;
  readonly lineHeight?: FontSize | undefined;
  readonly verticalAlign?: "top" | "middle" | "bottom" | undefined;
  /** Paragraph alignment when not explicitly overridden. */
  readonly align?: TextAlignment | undefined;
  /** Paragraph space-after when not explicitly overridden, in pixels. */
  readonly spacingAfter?: number | undefined;
  /** Character spacing in pixels. */
  readonly letterSpacing?: number | undefined;
  readonly textTransform?: "none" | "uppercase" | undefined;
  readonly textDecoration?: "none" | "underline" | "line-through" | undefined;
  readonly textShadow?: Shadow | "none" | undefined;
  /** Whether the text box wraps normally or keeps each line on a single line. */
  readonly textWrap?: "normal" | "nowrap" | undefined;
  /** Natural language (BCP 47 language tag, e.g. zh-CN, en-US). */
  readonly lang?: string | undefined;
};

const textStyleObjectSchema = z.strictObject({
  fontFamily: rawFontReferenceSchema.optional(),
  fontSize: validatedFontSizeSchema.optional(),
  fontWeight: fontWeightSchema.optional(),
  fontStyle: rawFontFaceStyleSchema.optional(),
  color: rawColorValueSchema.optional(),
  lineHeight: validatedLineHeightSchema.optional(),
  verticalAlign: z.enum(["top", "middle", "bottom"]).optional(),
  /** Paragraph alignment when not explicitly overridden. */
  align: rawTextAlignmentSchema.optional(),
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
  textShadow: z.union([rawShadowSchema, z.literal("none")]).optional(),
  /** Whether the text box wraps normally or keeps each line on a single line. */
  textWrap: z.enum(["normal", "nowrap"]).optional(),
  /** Natural language (BCP 47 language tag, e.g. zh-CN, en-US). */
  lang: bcp47LangSchema.optional(),
});

export const textStyleShape = textStyleObjectSchema.shape;

const rawTextStyleSchema = textStyleObjectSchema.readonly();
/** Inheritable text style fields shared across text objects and containers. */
export const textStyleSchema: z.ZodType<TextStyle> =
  rawTextStyleSchema satisfies z.ZodType<
    AssertExact<TextStyle, z.infer<typeof rawTextStyleSchema>> extends true
      ? TextStyle
      : never
  >;

export type FormatColorTransform = {
  readonly color: string;
  readonly alpha?: number | undefined;
  readonly lumMod?: number | undefined;
  readonly lumOff?: number | undefined;
  readonly tint?: number | undefined;
  readonly shade?: number | undefined;
  readonly satMod?: number | undefined;
};

const rawFormatColorTransformSchema = z
  .strictObject({
    color: z.union([z.literal("phClr"), rawSchemeColorNameSchema, z.string()]),
    alpha: z.number().finite().min(0).max(1).optional(),
    lumMod: z.number().finite().min(0).max(1).optional(),
    lumOff: z.number().finite().min(0).max(1).optional(),
    tint: z.number().finite().min(0).max(1).optional(),
    shade: z.number().finite().min(0).max(1).optional(),
    satMod: z.number().finite().min(0).optional(),
  })
  .readonly();
export const formatColorTransformSchema: z.ZodType<FormatColorTransform> =
  rawFormatColorTransformSchema satisfies z.ZodType<
    AssertExact<
      FormatColorTransform,
      z.infer<typeof rawFormatColorTransformSchema>
    > extends true
      ? FormatColorTransform
      : never
  >;

export type FormatGradientStop = {
  readonly position: number;
  readonly color: FormatColorTransform;
};

const rawFormatGradientStopSchema = z
  .strictObject({
    position: z.number().finite().min(0).max(1),
    color: rawFormatColorTransformSchema,
  })
  .readonly();
export const formatGradientStopSchema: z.ZodType<FormatGradientStop> =
  rawFormatGradientStopSchema satisfies z.ZodType<
    AssertExact<
      FormatGradientStop,
      z.infer<typeof rawFormatGradientStopSchema>
    > extends true
      ? FormatGradientStop
      : never
  >;

export type FormatFillStyle =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: FormatColorTransform }
  | {
      readonly kind: "gradient";
      readonly rotWithShape?: boolean | undefined;
      readonly stops: readonly FormatGradientStop[];
      readonly linear?:
        | {
            readonly angle: number;
            readonly scaled?: boolean | undefined;
          }
        | undefined;
      readonly path?:
        | {
            readonly shape: "circle" | "rect" | "shape";
            readonly fillToRect?:
              | {
                  readonly l: number;
                  readonly t: number;
                  readonly r: number;
                  readonly b: number;
                }
              | undefined;
          }
        | undefined;
    };

const rawFormatFillStyleSchema = z.union([
  z
    .strictObject({
      kind: z.literal("none"),
    })
    .readonly(),
  z
    .strictObject({
      kind: z.literal("solid"),
      color: rawFormatColorTransformSchema,
    })
    .readonly(),
  z
    .strictObject({
      kind: z.literal("gradient"),
      rotWithShape: z.boolean().optional(),
      stops: z.array(rawFormatGradientStopSchema).readonly(),
      linear: z
        .strictObject({
          angle: z.number().finite(),
          scaled: z.boolean().optional(),
        })
        .readonly()
        .optional(),
      path: z
        .strictObject({
          shape: z.enum(["circle", "rect", "shape"]),
          fillToRect: z
            .strictObject({
              l: z.number().finite(),
              t: z.number().finite(),
              r: z.number().finite(),
              b: z.number().finite(),
            })
            .readonly()
            .optional(),
        })
        .readonly()
        .optional(),
    })
    .readonly(),
]);
export const formatFillStyleSchema: z.ZodType<FormatFillStyle> =
  rawFormatFillStyleSchema satisfies z.ZodType<
    AssertExact<
      FormatFillStyle,
      z.infer<typeof rawFormatFillStyleSchema>
    > extends true
      ? FormatFillStyle
      : never
  >;

export type FormatLineStyle = {
  readonly widthEmu?: number | undefined;
  readonly cap?: "rnd" | "sq" | "flat" | undefined;
  readonly compound?:
    | "sng"
    | "dbl"
    | "thickThin"
    | "thinThick"
    | "tri"
    | undefined;
  readonly align?: "ctr" | "in" | undefined;
  readonly dash?: string | undefined;
  readonly fill?: FormatFillStyle | undefined;
};

const rawFormatLineStyleSchema = z
  .strictObject({
    widthEmu: z.number().int().nonnegative().optional(),
    cap: z.enum(["rnd", "sq", "flat"]).optional(),
    compound: z
      .enum(["sng", "dbl", "thickThin", "thinThick", "tri"])
      .optional(),
    align: z.enum(["ctr", "in"]).optional(),
    dash: z.string().optional(),
    fill: rawFormatFillStyleSchema.optional(),
  })
  .readonly();
export const formatLineStyleSchema: z.ZodType<FormatLineStyle> =
  rawFormatLineStyleSchema satisfies z.ZodType<
    AssertExact<
      FormatLineStyle,
      z.infer<typeof rawFormatLineStyleSchema>
    > extends true
      ? FormatLineStyle
      : never
  >;

export type FormatEffectStyle = {
  readonly shadow?: Shadow | undefined;
};

const rawFormatEffectStyleSchema = z
  .strictObject({
    shadow: rawShadowSchema.optional(),
  })
  .readonly();
export const formatEffectStyleSchema: z.ZodType<FormatEffectStyle> =
  rawFormatEffectStyleSchema satisfies z.ZodType<
    AssertExact<
      FormatEffectStyle,
      z.infer<typeof rawFormatEffectStyleSchema>
    > extends true
      ? FormatEffectStyle
      : never
  >;

/** DrawingML `<a:fmtScheme>` style matrix stored on a native theme or theme override. */
export type FormatScheme = {
  readonly name?: string | undefined;
  readonly fillStyles?: readonly FormatFillStyle[] | undefined;
  readonly lineStyles?: readonly FormatLineStyle[] | undefined;
  readonly effectStyles?: readonly FormatEffectStyle[] | undefined;
  readonly bgFillStyles?: readonly FormatFillStyle[] | undefined;
};

const rawFormatSchemeSchema = z
  .strictObject({
    name: z.string().optional(),
    fillStyles: z.array(rawFormatFillStyleSchema).readonly().optional(),
    lineStyles: z.array(rawFormatLineStyleSchema).readonly().optional(),
    effectStyles: z.array(rawFormatEffectStyleSchema).readonly().optional(),
    bgFillStyles: z.array(rawFormatFillStyleSchema).readonly().optional(),
  })
  .readonly();
/** DrawingML `<a:fmtScheme>` style matrix stored on a native theme or theme override. */
export const formatSchemeSchema: z.ZodType<FormatScheme> =
  rawFormatSchemeSchema satisfies z.ZodType<
    AssertExact<
      FormatScheme,
      z.infer<typeof rawFormatSchemeSchema>
    > extends true
      ? FormatScheme
      : never
  >;

/** Deck-level or master-level theme input; major/minor fonts are written to native PPTX theme. */
export type ThemeInput = {
  readonly name?: string | undefined;
  readonly colors?:
    | Readonly<
        {
          dk1?: string | undefined;
          lt1?: string | undefined;
          dk2?: string | undefined;
          lt2?: string | undefined;
          accent1?: string | undefined;
          accent2?: string | undefined;
          accent3?: string | undefined;
          accent4?: string | undefined;
          accent5?: string | undefined;
          accent6?: string | undefined;
          hlink?: string | undefined;
          folHlink?: string | undefined;
        } & Record<string, ThemeColor>
      >
    | undefined;
  readonly fonts?:
    | Readonly<
        {
          major?: FontToken | undefined;
          minor?: FontToken | undefined;
        } & Record<string, FontToken>
      >
    | undefined;
  readonly formatScheme?: FormatScheme | undefined;
};

const rawThemeInputSchema = z
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
      .catchall(rawThemeColorSchema)
      .readonly()
      .optional(),
    fonts: z
      .object({
        major: rawFontReferenceSchema.optional(),
        minor: rawFontReferenceSchema.optional(),
      })
      .catchall(rawFontReferenceSchema)
      .readonly()
      .optional(),
    formatScheme: rawFormatSchemeSchema.optional(),
  })
  .readonly();
/** Deck-level or master-level theme input; major/minor fonts are written to native PPTX theme. */
export const themeInputSchema: z.ZodType<ThemeInput> =
  rawThemeInputSchema satisfies z.ZodType<
    AssertExact<ThemeInput, z.infer<typeof rawThemeInputSchema>> extends true
      ? ThemeInput
      : never
  >;

/** Native `<a:themeOverride>` partial override for a SlideLayout or Slide. */
export type ThemeOverride = {
  readonly colors?:
    | {
        readonly dk1?: string | undefined;
        readonly lt1?: string | undefined;
        readonly dk2?: string | undefined;
        readonly lt2?: string | undefined;
        readonly accent1?: string | undefined;
        readonly accent2?: string | undefined;
        readonly accent3?: string | undefined;
        readonly accent4?: string | undefined;
        readonly accent5?: string | undefined;
        readonly accent6?: string | undefined;
        readonly hlink?: string | undefined;
        readonly folHlink?: string | undefined;
      }
    | undefined;
  readonly fonts?:
    | {
        readonly major?: FontToken | undefined;
        readonly minor?: FontToken | undefined;
      }
    | undefined;
  readonly formatScheme?: FormatScheme | undefined;
};

const rawThemeOverrideSchema = z
  .strictObject({
    colors: z
      .strictObject({
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
      .readonly()
      .optional(),
    fonts: z
      .strictObject({
        major: rawFontReferenceSchema.optional(),
        minor: rawFontReferenceSchema.optional(),
      })
      .readonly()
      .optional(),
    formatScheme: rawFormatSchemeSchema.optional(),
  })
  .readonly();
/** Native `<a:themeOverride>` partial override for a SlideLayout or Slide. */
export const themeOverrideSchema: z.ZodType<ThemeOverride> =
  rawThemeOverrideSchema satisfies z.ZodType<
    AssertExact<
      ThemeOverride,
      z.infer<typeof rawThemeOverrideSchema>
    > extends true
      ? ThemeOverride
      : never
  >;

export type ThemeData = {
  readonly name: string;
  readonly colors: Readonly<
    {
      dk1: string;
      lt1: string;
      dk2: string;
      lt2: string;
      accent1: string;
      accent2: string;
      accent3: string;
      accent4: string;
      accent5: string;
      accent6: string;
      hlink: string;
      folHlink: string;
    } & Record<string, ThemeColor>
  >;
  readonly fonts: Readonly<
    {
      major: FontToken;
      minor: FontToken;
    } & Record<string, FontToken>
  >;
  readonly formatScheme?: FormatScheme | undefined;
};

const rawThemeDataSchema = z
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
      .catchall(rawThemeColorSchema)
      .readonly(),
    fonts: z
      .object({
        major: rawFontReferenceSchema,
        minor: rawFontReferenceSchema,
      })
      .catchall(rawFontReferenceSchema)
      .readonly(),
    formatScheme: rawFormatSchemeSchema.optional(),
  })
  .readonly();
export const themeDataSchema: z.ZodType<ThemeData> =
  rawThemeDataSchema satisfies z.ZodType<
    AssertExact<ThemeData, z.infer<typeof rawThemeDataSchema>> extends true
      ? ThemeData
      : never
  >;

export type Theme = ThemeData & {
  /** Updates this theme in place while preserving instance identity across all referencing SlideMasters. */
  set(patch: ThemeInput): Theme;
};

export const themeSchema: z.ZodType<Theme> = rawThemeDataSchema.transform(
  (data) => defineTheme(data),
);

/** Local theme scopes on slides or containers can only override color and font tokens already present in the deck. */
export type ThemeScope = {
  readonly colors?: Readonly<Record<string, ThemeColor>> | undefined;
  readonly fonts?: Readonly<Record<string, FontToken>> | undefined;
};

const rawThemeScopeSchema = z
  .strictObject({
    colors: z.record(z.string(), rawThemeColorSchema).readonly().optional(),
    fonts: z.record(z.string(), rawFontReferenceSchema).readonly().optional(),
  })
  .readonly();
/** Local theme scopes on slides or containers can only override color and font tokens already present in the deck. */
export const themeScopeSchema: z.ZodType<ThemeScope> =
  rawThemeScopeSchema satisfies z.ZodType<
    AssertExact<ThemeScope, z.infer<typeof rawThemeScopeSchema>> extends true
      ? ThemeScope
      : never
  >;

export type ResolvedThemeFont = {
  readonly family: string;
  readonly font?: Font | undefined;
  readonly themeFont?: "major" | "minor" | undefined;
};

const rawResolvedThemeFontSchema = z
  .strictObject({
    family: z.string(),
    font: rawFontSchema.optional(),
    themeFont: z.enum(["major", "minor"]).optional(),
  })
  .readonly();
export const resolvedThemeFontSchema: z.ZodType<ResolvedThemeFont> =
  rawResolvedThemeFontSchema satisfies z.ZodType<
    AssertExact<
      ResolvedThemeFont,
      z.infer<typeof rawResolvedThemeFontSchema>
    > extends true
      ? ResolvedThemeFont
      : never
  >;

/** Theme, local scopes, active color map, native overrides, and accumulated container text defaults in a single build. */
export type ThemeContext = {
  readonly theme: Theme;
  readonly colorMap: ColorMap;
  readonly formatScheme: FormatScheme | undefined;
  readonly textStyle: TextStyle;
  with(scope?: ThemeScope, style?: TextStyle): ThemeContext;
  /** Replaces the accumulated container text defaults while retaining the same theme scope chain. */
  withTextStyle(style: TextStyle): ThemeContext;
  /** Derives a child context applying native `<a:themeOverride>` and/or `<p:clrMapOvr>` without mutating `theme`. */
  withNativeOverrides(
    themeOverride?: ThemeOverride,
    colorMapOverride?: ColorMapOverride,
  ): ThemeContext;
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
    readonly themeOverride?: ThemeOverride;
  }
>();

export const textStyleKeys = new Set(Object.keys(textStyleObjectSchema.shape));

function textStyleErrorForIssue(
  issue: z.core.$ZodIssue | undefined,
  rawStyle: Readonly<Record<string, unknown>>,
): CapabilityError {
  const field = issue?.path[0];
  if (field === "fontFamily") {
    const hasFacesIssue =
      issue?.path[1] === "faces" ||
      (issue?.code === "invalid_union" &&
        issue.errors.some((branch) =>
          branch.some((sub) => sub.path[0] === "faces"),
        ));
    if (hasFacesIssue) {
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
      "fontWeight must be normal, bold, or an integer between 1 and 1000",
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

function schemeSlot(value: string): value is SchemeColorSlot {
  return (schemeColorSlots as readonly string[]).includes(value);
}

function schemeColorName(value: string): value is SchemeColorName {
  return (schemeColorNames as readonly string[]).includes(value);
}

export function isSchemeColorSlot(value: string): value is SchemeColorSlot {
  return schemeSlot(value);
}

export function isSchemeColorName(value: string): value is SchemeColorName {
  return schemeColorName(value);
}

export function isFont(value: unknown): value is Font {
  return record(value) && typeof value["family"] === "string";
}

const themeBrandSymbol = Symbol.for("@pptxdk/core/theme");

/** Checks whether a runtime value is an identity-bearing `Theme` instance produced by `defineTheme`. */
export function isTheme(value: unknown): value is Theme {
  return (
    typeof value === "object" && value !== null && themeBrandSymbol in value
  );
}

class NativeThemeRecord implements Theme {
  readonly [themeBrandSymbol] = true as const;
  #name: string;
  #colors: ThemeData["colors"];
  #fonts: ThemeData["fonts"];
  #formatScheme: FormatScheme | undefined;

  constructor(initial: ThemeData) {
    this.#name = initial.name;
    this.#colors = initial.colors;
    this.#fonts = initial.fonts;
    this.#formatScheme = initial.formatScheme;
    Object.defineProperties(this, {
      name: {
        enumerable: true,
        configurable: false,
        get: () => this.#name,
      },
      colors: {
        enumerable: true,
        configurable: false,
        get: () => this.#colors,
      },
      fonts: {
        enumerable: true,
        configurable: false,
        get: () => this.#fonts,
      },
      ...(initial.formatScheme === undefined
        ? {}
        : {
            formatScheme: {
              enumerable: true,
              configurable: false,
              get: () => this.#formatScheme,
            },
          }),
    });
  }

  get name(): string {
    return this.#name;
  }

  get colors(): ThemeData["colors"] {
    return this.#colors;
  }

  get fonts(): ThemeData["fonts"] {
    return this.#fonts;
  }

  get formatScheme(): FormatScheme | undefined {
    return this.#formatScheme;
  }

  set(patch: ThemeInput): Theme {
    if (this === defaultTheme) {
      throw new TypeError("Cannot mutate defaultTheme");
    }
    if (!record(patch)) {
      throw new TypeError("Theme update must be an object");
    }
    const parsed = themeInputSchema.safeParse(patch);
    if (!parsed.success) {
      const issue = parsed.error.issues[0];
      throw new CapabilityError(
        "invalid-theme",
        issue?.message ?? "Invalid theme declaration",
      );
    }
    const validPatch = parsed.data;
    if (validPatch.name !== undefined) {
      this.#name = validPatch.name;
    }
    if (validPatch.colors !== undefined) {
      this.#colors = snapshotValue({
        ...this.#colors,
        ...validPatch.colors,
      }) as ThemeData["colors"];
    }
    if (validPatch.fonts !== undefined) {
      this.#fonts = snapshotValue({
        ...this.#fonts,
        ...validPatch.fonts,
      }) as ThemeData["fonts"];
    }
    if (validPatch.formatScheme !== undefined) {
      const hadFormatScheme = Object.hasOwn(this, "formatScheme");
      this.#formatScheme = snapshotValue(validPatch.formatScheme);
      if (!hadFormatScheme) {
        Object.defineProperty(this, "formatScheme", {
          enumerable: true,
          configurable: false,
          get: () => this.#formatScheme,
        });
      }
    }
    return this;
  }
}

function parseThemeInputData(input: ThemeInput = {}): ThemeData {
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
  }) as ThemeData["colors"];
  const fonts = snapshotValue({
    ...validInput.fonts,
    major: validInput.fonts?.major ?? "Aptos Display",
    minor: validInput.fonts?.minor ?? "Aptos",
  }) as ThemeData["fonts"];
  const formatScheme =
    validInput.formatScheme === undefined
      ? undefined
      : snapshotValue(validInput.formatScheme);
  return {
    name: validInput.name ?? "Office Theme",
    colors,
    fonts,
    ...(formatScheme === undefined ? {} : { formatScheme }),
  };
}

function normalizeTheme(input?: ThemeInput | Theme): Theme {
  if (input === undefined) return defaultTheme;
  if (isTheme(input)) return input;
  return new NativeThemeRecord(parseThemeInputData(input));
}

/** Fill unsupplied theme items using initial Office palette and fonts, returning an identity-bearing `Theme`. */
export function defineTheme(input: ThemeInput = {}): Theme {
  return new NativeThemeRecord(parseThemeInputData(input));
}

export const defaultTheme: Theme = defineTheme();

/** Generate darker variant in PowerPoint native palette. */
export function darker(slot: SchemeColorName, amount: number): ThemeShade {
  return { slot, lumMod: 1 - amount };
}

/** Generate lighter variant in PowerPoint native palette. */
export function lighter(slot: SchemeColorName, amount: number): ThemeShade {
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

function rgbChannelsFromHex(hex: string): [number, number, number] {
  return [0, 2, 4].map(
    (start) => Number.parseInt(hex.slice(start, start + 2), 16) / 255,
  ) as [number, number, number];
}

function hexFromRgbChannels(r: number, g: number, b: number): string {
  return [r, g, b]
    .map((channel) =>
      Math.round(Math.min(1, Math.max(0, channel)) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
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
  return [hue, saturation, lightness];
}

function hslToRgb(
  hue: number,
  saturation: number,
  lightness: number,
): [number, number, number] {
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const x = chroma * (1 - Math.abs((hue % 2) - 1));
  const m = lightness - chroma / 2;
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
  return [r1 + m, g1 + m, b1 + m];
}

function shadeHex(hex: string, lumMod: number, lumOff: number): string {
  const [r, g, b] = rgbChannelsFromHex(hex);
  const [hue, saturation, lightness] = rgbToHsl(r, g, b);
  const shaded = Math.min(1, Math.max(0, lightness * lumMod + lumOff));
  const [r1, g1, b1] = hslToRgb(hue, saturation, shaded);
  return hexFromRgbChannels(r1, g1, b1);
}

/**
 * Applies DrawingML color modifiers (`tint`, `shade`, `satMod`, `lumMod`, `lumOff`, `alpha`)
 * onto a resolved base color (used by theme format schemes and `<p:bgRef>` evaluation).
 */
export function applyFormatColorTransforms(
  base: ResolvedColor,
  transform: FormatColorTransform,
): ResolvedColor {
  let [r, g, b] = rgbChannelsFromHex(base.hex);
  if (transform.tint !== undefined) {
    const t = transform.tint;
    r = r * t + (1 - t);
    g = g * t + (1 - t);
    b = b * t + (1 - t);
  }
  if (transform.shade !== undefined) {
    const s = transform.shade;
    r *= s;
    g *= s;
    b *= s;
  }
  if (
    transform.satMod !== undefined ||
    transform.lumMod !== undefined ||
    transform.lumOff !== undefined
  ) {
    const [hue, sat, lum] = rgbToHsl(r, g, b);
    const nextSat = Math.min(1, Math.max(0, sat * (transform.satMod ?? 1)));
    const nextLum = Math.min(
      1,
      Math.max(0, lum * (transform.lumMod ?? 1) + (transform.lumOff ?? 0)),
    );
    [r, g, b] = hslToRgb(hue, nextSat, nextLum);
  }
  const alpha =
    transform.alpha === undefined
      ? base.alpha
      : (base.alpha ?? 1) * transform.alpha;
  return {
    hex: hexFromRgbChannels(r, g, b),
    ...(alpha === undefined || alpha === 1 ? {} : { alpha }),
    ...(base.slot === undefined ? {} : { slot: base.slot }),
    ...(base.scheme === undefined ? {} : { scheme: base.scheme }),
  };
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

function effectiveThemeColor(
  context: ThemeContext,
  slot: NativeColorSlot,
): string {
  const overrideColor =
    contextRecords.get(context)?.themeOverride?.colors?.[slot];
  return overrideColor ?? context.theme.colors[slot];
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
  if (schemeSlot(value)) {
    const mappedSlot = context.colorMap[value];
    const configured = effectiveThemeColor(context, mappedSlot);
    const resolved = resolveColorName(
      configured,
      rootContext(context),
      depth + 1,
    );
    return { ...resolved, slot: mappedSlot, scheme: value };
  }
  if (nativeSlot(value)) {
    const configured = effectiveThemeColor(context, value);
    const resolved = resolveColorName(
      configured,
      rootContext(context),
      depth + 1,
    );
    return { ...resolved, slot: value, scheme: value };
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
    !schemeColorName(String(value["slot"])) ||
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
  const scheme = value["slot"];
  const base = resolveColorName(scheme, context, depth + 1);
  return {
    hex: shadeHex(base.hex, lumMod, lumOff ?? 0),
    ...(base.slot === undefined ? {} : { slot: base.slot }),
    scheme,
    lumMod,
    ...(lumOff === undefined ? {} : { lumOff }),
  };
}

function mergeThemeOverrides(
  base: ThemeOverride | undefined,
  next: ThemeOverride | undefined,
): ThemeOverride | undefined {
  if (base === undefined) return next;
  if (next === undefined) return base;
  const colors =
    base.colors === undefined && next.colors === undefined
      ? undefined
      : { ...base.colors, ...next.colors };
  const fonts =
    base.fonts === undefined && next.fonts === undefined
      ? undefined
      : { ...base.fonts, ...next.fonts };
  const formatScheme = next.formatScheme ?? base.formatScheme;
  return Object.freeze({
    ...(colors === undefined ? {} : { colors: Object.freeze(colors) }),
    ...(fonts === undefined ? {} : { fonts: Object.freeze(fonts) }),
    ...(formatScheme === undefined ? {} : { formatScheme }),
  });
}

function createContext(
  theme: Theme,
  parent: ThemeContext | undefined,
  scope: ThemeScope | undefined,
  textStyle: TextStyle,
  colorMap: ColorMap = defaultColorMap,
  themeOverride?: ThemeOverride,
): ThemeContext {
  const formatScheme = themeOverride?.formatScheme ?? theme.formatScheme;
  let context: ThemeContext;
  context = Object.freeze({
    theme,
    colorMap,
    formatScheme,
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
          if (schemeColorName(name)) {
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
        colorMap,
        themeOverride,
      );
      const validated =
        style === undefined ? {} : textStyleFrom(style, candidateContext);
      const nextStyle = { ...context.textStyle, ...validated };
      return createContext(
        theme,
        context,
        checkedScope,
        nextStyle,
        colorMap,
        themeOverride,
      );
    },
    withTextStyle(style) {
      return createContext(
        theme,
        context,
        undefined,
        textStyleFrom(style, context),
        colorMap,
        themeOverride,
      );
    },
    withNativeOverrides(nextThemeOverride, nextColorMapOverride) {
      const mergedOverride = mergeThemeOverrides(
        themeOverride,
        nextThemeOverride,
      );
      const nextColorMap =
        nextColorMapOverride === undefined || nextColorMapOverride === "master"
          ? colorMap
          : nextColorMapOverride;
      return createContext(
        theme,
        context,
        undefined,
        context.textStyle,
        nextColorMap,
        mergedOverride,
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
      const themeFontName: ResolvedThemeFont["themeFont"] =
        value === "major" || value === "minor" ? value : undefined;
      const overrideFont =
        themeFontName === undefined
          ? undefined
          : themeOverride?.fonts?.[themeFontName];
      const token =
        local?.value ??
        overrideFont ??
        (Object.hasOwn(theme.fonts, value) ? theme.fonts[value]! : value);
      const globalThemeFont =
        themeFontName === undefined
          ? undefined
          : (overrideFont ?? theme.fonts[themeFontName]);
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
    ...(themeOverride === undefined ? {} : { themeOverride }),
  });
  return context;
}

/** Create a theme chain unique to a single build; local scopes do not mutate the shared theme object. */
export function createThemeContext(
  input?: ThemeInput | Theme,
  style?: TextStyle,
  options: {
    readonly colorMap?: ColorMap;
    readonly themeOverride?: ThemeOverride;
  } = {},
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
  const colorMap = options.colorMap ?? defaultColorMap;
  const initialContext = createContext(
    theme,
    undefined,
    undefined,
    {},
    colorMap,
    options.themeOverride,
  );
  return createContext(
    theme,
    undefined,
    undefined,
    textStyleFrom(style, initialContext),
    colorMap,
    options.themeOverride,
  );
}
