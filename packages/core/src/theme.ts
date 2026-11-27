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

export const schemeColorSlotSchema = z.enum(schemeColorSlots);
export type SchemeColorSlot = z.infer<typeof schemeColorSlotSchema>;

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

export const schemeColorNameSchema = z.enum(schemeColorNames);
export type SchemeColorName = z.infer<typeof schemeColorNameSchema>;

/** Native `<p:clrMap>` / `<a:overrideClrMapping>` mapping the 12 scheme slots to 12 theme palette slots. */
export const colorMapSchema = z
  .strictObject({
    bg1: nativeColorSlotSchema,
    tx1: nativeColorSlotSchema,
    bg2: nativeColorSlotSchema,
    tx2: nativeColorSlotSchema,
    accent1: nativeColorSlotSchema,
    accent2: nativeColorSlotSchema,
    accent3: nativeColorSlotSchema,
    accent4: nativeColorSlotSchema,
    accent5: nativeColorSlotSchema,
    accent6: nativeColorSlotSchema,
    hlink: nativeColorSlotSchema,
    folHlink: nativeColorSlotSchema,
  })
  .readonly();
export type ColorMap = z.infer<typeof colorMapSchema>;

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

export const colorMapOverrideSchema = z.union([
  z.literal("master"),
  colorMapSchema,
]);
export type ColorMapOverride = z.infer<typeof colorMapOverrideSchema>;

/** Native theme color represented with DrawingML luminance modulation. */
export const themeShadeSchema = z
  .strictObject({
    slot: schemeColorNameSchema,
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
    scheme: schemeColorNameSchema.optional(),
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

export const imageBackgroundFitSchema = z.enum(["fill", "cover", "contain"]);
export type ImageBackgroundFit = z.infer<typeof imageBackgroundFitSchema>;

const percentPositionPattern = /^(-?(?:\d+(?:\.\d*)?|\.\d+))%$/u;
const positionPercentSchema = z
  .string()
  .regex(percentPositionPattern) as unknown as z.ZodType<`${number}%`>;
const positionCoordinateSchema = z.union([
  z.number().finite(),
  positionPercentSchema,
]);

export const imageBackgroundPositionKeywordSchema = z.enum([
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

export const imageBackgroundPositionSchema = z.union([
  imageBackgroundPositionKeywordSchema,
  z
    .strictObject({
      x: positionCoordinateSchema.optional(),
      y: positionCoordinateSchema.optional(),
    })
    .readonly(),
]);
export type ImageBackgroundPosition = z.infer<
  typeof imageBackgroundPositionSchema
>;

export const externalImageFillSourceSchema = z
  .strictObject({
    externalUrl: z.string().min(1),
  })
  .readonly();
export type ExternalImageFillSource = z.infer<
  typeof externalImageFillSourceSchema
>;

export const imageFillSourceSchema = z.union([
  z.string(),
  z.instanceof(URL),
  uint8ArraySchema,
  externalImageFillSourceSchema,
]);
export type ImageFillSource = z.infer<typeof imageFillSourceSchema>;

export const imageBackgroundSchema = z
  .strictObject({
    image: imageFillSourceSchema,
    fit: imageBackgroundFitSchema.optional(),
    position: imageBackgroundPositionSchema.optional(),
    opacity: z.number().finite().min(0).max(1).optional(),
  })
  .readonly();
export type ImageBackground = z.infer<typeof imageBackgroundSchema>;

export const backgroundValueSchema = z.union([
  z.literal("none"),
  colorValueSchema,
  linearGradientSchema,
  radialGradientSchema,
  imageBackgroundSchema,
]);
export type BackgroundValue = z.infer<typeof backgroundValueSchema>;

export const borderStyleSchema = z.enum(["solid", "dashed", "dotted"]);
export type BorderStyle = z.infer<typeof borderStyleSchema>;

export const borderObjectSchema = z
  .strictObject({
    color: colorValueSchema,
    width: z.number().finite().nonnegative().optional(),
    style: borderStyleSchema.optional(),
  })
  .readonly();
export type BorderObject = z.infer<typeof borderObjectSchema>;

export const borderValueSchema = z.union([
  z.literal("none"),
  borderObjectSchema,
]);
export type BorderValue = z.infer<typeof borderValueSchema>;

export const shadowSchema = z
  .strictObject({
    color: colorValueSchema.optional(),
    blur: z.number().finite().nonnegative().optional(),
    distance: z.number().finite().nonnegative().optional(),
    angle: z.number().finite().optional(),
  })
  .readonly();
export type Shadow = z.infer<typeof shadowSchema>;

export const boxShadowValueSchema = z.union([z.literal("none"), shadowSchema]);
export type BoxShadowValue = z.infer<typeof boxShadowValueSchema>;

/** Visual fields shared across native DrawingML shapes, images, text boxes, and visible Boxes. */
export const visualStyleSchema = z
  .strictObject({
    background: backgroundValueSchema.optional(),
    border: borderValueSchema.optional(),
    borderRadius: z.number().finite().optional(),
    opacity: z.number().finite().optional(),
    boxShadow: boxShadowValueSchema.optional(),
    softEdge: z.number().finite().optional(),
    rotate: z.number().finite().optional(),
    flipH: z.boolean().optional(),
    flipV: z.boolean().optional(),
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

export const formatColorTransformSchema = z
  .strictObject({
    color: z.union([z.literal("phClr"), schemeColorNameSchema, z.string()]),
    alpha: z.number().finite().min(0).max(1).optional(),
    lumMod: z.number().finite().min(0).max(1).optional(),
    lumOff: z.number().finite().min(0).max(1).optional(),
    tint: z.number().finite().min(0).max(1).optional(),
    shade: z.number().finite().min(0).max(1).optional(),
    satMod: z.number().finite().min(0).optional(),
  })
  .readonly();
export type FormatColorTransform = z.infer<typeof formatColorTransformSchema>;

export const formatGradientStopSchema = z
  .strictObject({
    position: z.number().finite().min(0).max(1),
    color: formatColorTransformSchema,
  })
  .readonly();
export type FormatGradientStop = z.infer<typeof formatGradientStopSchema>;

export const formatFillStyleSchema = z.union([
  z
    .strictObject({
      kind: z.literal("none"),
    })
    .readonly(),
  z
    .strictObject({
      kind: z.literal("solid"),
      color: formatColorTransformSchema,
    })
    .readonly(),
  z
    .strictObject({
      kind: z.literal("gradient"),
      rotWithShape: z.boolean().optional(),
      stops: z.array(formatGradientStopSchema).readonly(),
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
export type FormatFillStyle = z.infer<typeof formatFillStyleSchema>;

export const formatLineStyleSchema = z
  .strictObject({
    widthEmu: z.number().int().nonnegative().optional(),
    cap: z.enum(["rnd", "sq", "flat"]).optional(),
    compound: z
      .enum(["sng", "dbl", "thickThin", "thinThick", "tri"])
      .optional(),
    align: z.enum(["ctr", "in"]).optional(),
    dash: z.string().optional(),
    fill: formatFillStyleSchema.optional(),
  })
  .readonly();
export type FormatLineStyle = z.infer<typeof formatLineStyleSchema>;

export const formatEffectStyleSchema = z
  .strictObject({
    shadow: shadowSchema.optional(),
  })
  .readonly();
export type FormatEffectStyle = z.infer<typeof formatEffectStyleSchema>;

/** DrawingML `<a:fmtScheme>` style matrix stored on a native theme or theme override. */
export const formatSchemeSchema = z
  .strictObject({
    name: z.string().optional(),
    fillStyles: z.array(formatFillStyleSchema).readonly().optional(),
    lineStyles: z.array(formatLineStyleSchema).readonly().optional(),
    effectStyles: z.array(formatEffectStyleSchema).readonly().optional(),
    bgFillStyles: z.array(formatFillStyleSchema).readonly().optional(),
  })
  .readonly();
export type FormatScheme = z.infer<typeof formatSchemeSchema>;

/** Deck-level or master-level theme input; major/minor fonts are written to native PPTX theme. */
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
    formatScheme: formatSchemeSchema.optional(),
  })
  .readonly();
export type ThemeInput = z.infer<typeof themeInputSchema>;

/** Native `<a:themeOverride>` partial override for a SlideLayout or Slide. */
export const themeOverrideSchema = z
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
        major: fontTokenSchema.optional(),
        minor: fontTokenSchema.optional(),
      })
      .readonly()
      .optional(),
    formatScheme: formatSchemeSchema.optional(),
  })
  .readonly();
export type ThemeOverride = z.infer<typeof themeOverrideSchema>;

export const themeDataSchema = z
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
    formatScheme: formatSchemeSchema.optional(),
  })
  .readonly();
export type ThemeData = z.infer<typeof themeDataSchema>;

export type Theme = ThemeData & {
  /** Updates this theme in place while preserving instance identity across all referencing SlideMasters. */
  set(patch: ThemeInput): Theme;
};

export const themeSchema: z.ZodType<Theme> = themeDataSchema.transform((data) =>
  defineTheme(data),
);

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
