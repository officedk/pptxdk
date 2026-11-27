import { z } from "zod";

import { authorIdSchema, readAuthorId, type AuthorId } from "#/author-ids.ts";
import { BuildError, CapabilityError } from "#/diagnostics.ts";
import {
  layoutStyleShape,
  slideLayoutStyleShape,
  type LayoutStyle,
} from "#/layout-schemas.ts";
import { placeholderSizeSchema, type PlaceholderSize } from "#/placeholders.ts";
import {
  backgroundValueSchema,
  colorMapOverrideSchema,
  colorValueSchema,
  textStyleSchema,
  textStyleShape,
  themeInputSchema,
  themeOverrideSchema,
  themeScopeSchema,
  type BackgroundValue,
  type ColorMapOverride,
  type TextStyle,
  type ThemeInput,
  type ThemeOverride,
  type ThemeScope,
} from "#/theme.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

const minimumSlideSizeEmu = 914_400;
const maximumSlideSizeEmu = 51_206_400;
const emuPerPx = 9525;

export const xml10StringSchema = z.string().refine((value) => {
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      codePoint !== 0x9 &&
      codePoint !== 0xa &&
      codePoint !== 0xd &&
      (codePoint < 0x20 ||
        (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
        codePoint === 0xfffe ||
        codePoint === 0xffff)
    ) {
      return false;
    }
  }
  return true;
}, "String contains characters not allowed in XML 1.0");

export const slideDimensionSchema = z
  .number()
  .finite()
  .positive()
  .refine(
    (value) => {
      const sizeEmu = Math.round(value * emuPerPx);
      return sizeEmu >= minimumSlideSizeEmu && sizeEmu <= maximumSlideSizeEmu;
    },
    {
      message: `Converted slide size must be within the closed interval of ${minimumSlideSizeEmu} to ${maximumSlideSizeEmu} EMU`,
    },
  );

/** DrawingML `<p:bgRef idx="1001..1003">` reference resolved against `Theme.formatScheme.bgFillStyles`. */
export type BackgroundRef = {
  readonly ref: 1001 | 1002 | 1003;
  readonly color?: import("#/theme.ts").ColorValue | undefined;
};

const rawBackgroundRefSchema = z
  .strictObject({
    ref: z.union([z.literal(1001), z.literal(1002), z.literal(1003)]),
    color: colorValueSchema.optional(),
  })
  .readonly();
/** DrawingML `<p:bgRef idx="1001..1003">` reference resolved against `Theme.formatScheme.bgFillStyles`. */
export const backgroundRefSchema: z.ZodType<BackgroundRef> =
  rawBackgroundRefSchema satisfies z.ZodType<
    AssertExact<
      BackgroundRef,
      z.infer<typeof rawBackgroundRefSchema>
    > extends true
      ? BackgroundRef
      : never
  >;

/** Slide background can use theme colors, gradients supported by the shared visual writer, formal image resource inputs, `"none"`, or theme background fill references (`p:bgRef`). */
export type SlideBackground = BackgroundValue | BackgroundRef;

const rawSlideBackgroundSchema = z.union([
  backgroundValueSchema,
  rawBackgroundRefSchema,
]);
/** Slide background can use theme colors, gradients supported by the shared visual writer, formal image resource inputs, `"none"`, or theme background fill references (`p:bgRef`). */
export const slideBackgroundSchema: z.ZodType<SlideBackground> =
  rawSlideBackgroundSchema satisfies z.ZodType<
    AssertExact<
      SlideBackground,
      z.infer<typeof rawSlideBackgroundSchema>
    > extends true
      ? SlideBackground
      : never
  >;

export type DeckProps = {
  readonly width?: number | undefined;
  readonly height?: number | undefined;
  readonly theme?: ThemeInput | undefined;
  readonly style?: TextStyle | undefined;
  readonly title?: string | undefined;
};

const deckPropsObjectSchema = z.strictObject({
  width: slideDimensionSchema.optional(),
  height: slideDimensionSchema.optional(),
  theme: themeInputSchema.optional(),
  style: textStyleSchema.optional(),
  title: xml10StringSchema.optional(),
});
const rawDeckPropsSchema = deckPropsObjectSchema.readonly();
export const deckPropsSchema: z.ZodType<DeckProps> =
  rawDeckPropsSchema satisfies z.ZodType<
    AssertExact<DeckProps, z.infer<typeof rawDeckPropsSchema>> extends true
      ? DeckProps
      : never
  >;

export type SlideStyle = TextStyle &
  Omit<LayoutStyle, "zIndex"> & {
    readonly background?: SlideBackground | undefined;
  };

const rawSlideStyleSchema = z
  .strictObject({
    ...textStyleShape,
    ...slideLayoutStyleShape,
    background: rawSlideBackgroundSchema.optional(),
  })
  .readonly();
export const slideStyleSchema: z.ZodType<SlideStyle> =
  rawSlideStyleSchema satisfies z.ZodType<
    AssertExact<SlideStyle, z.infer<typeof rawSlideStyleSchema>> extends true
      ? SlideStyle
      : never
  >;

const slideCorePropsObjectSchema = z.strictObject({
  /** Page author id; unique within the deck and written as the stable slide identifier. */
  id: authorIdSchema,
  /** Optional PowerPoint display name written to `<p:cSld name="...">` without altering the stable `id`. */
  name: xml10StringSchema.optional(),
  layout: z.unknown().optional(),
  theme: themeScopeSchema.optional(),
  colorMapOverride: colorMapOverrideSchema.optional(),
  themeOverride: themeOverrideSchema.optional(),
  style: rawSlideStyleSchema.optional(),
  hidden: z.boolean().optional(),
  showMasterPlaceholderAnimations: z.boolean().optional(),
  showMasterShapes: z.boolean().optional(),
});
export const slideCorePropsSchema: z.ZodType<
  Omit<
    SlideProps,
    | "transition"
    | "advanceAfter"
    | "background"
    | "dateTime"
    | "footer"
    | "slideNumber"
  >
> = slideCorePropsObjectSchema.readonly();

export { placeholderSizeSchema, type PlaceholderSize };

/** Footer placeholder slot style combining inherited text styles and LayoutStyle placement. */
export type FooterPlaceholderStyle = TextStyle & LayoutStyle;

const rawFooterPlaceholderStyleSchema = z
  .strictObject({
    ...textStyleShape,
    ...layoutStyleShape,
  })
  .readonly();
/** Footer placeholder slot style combining inherited text styles and LayoutStyle placement. */
export const footerPlaceholderStyleSchema: z.ZodType<FooterPlaceholderStyle> =
  rawFooterPlaceholderStyleSchema satisfies z.ZodType<
    AssertExact<
      FooterPlaceholderStyle,
      z.infer<typeof rawFooterPlaceholderStyleSchema>
    > extends true
      ? FooterPlaceholderStyle
      : never
  >;

/** Per-slide dateTime/footer override declaration (`false` hides; string sets text; object overrides text/style/slot). */
export type SlideFooterValue =
  | boolean
  | string
  | {
      readonly text?: string | undefined;
      readonly style?: FooterPlaceholderStyle | undefined;
      readonly idx?: number | undefined;
      readonly size?: PlaceholderSize | undefined;
      readonly visible?: boolean | undefined;
    };

const rawSlideFooterValueSchema = z.union([
  z.boolean(),
  xml10StringSchema,
  z
    .strictObject({
      text: xml10StringSchema.optional(),
      style: rawFooterPlaceholderStyleSchema.optional(),
      idx: z.number().int().nonnegative().optional(),
      size: placeholderSizeSchema.optional(),
      visible: z.boolean().optional(),
    })
    .readonly(),
]);
/** Per-slide dateTime/footer override declaration (`false` hides; string sets text; object overrides text/style/slot). */
export const slideFooterValueSchema: z.ZodType<SlideFooterValue> =
  rawSlideFooterValueSchema satisfies z.ZodType<
    AssertExact<
      SlideFooterValue,
      z.infer<typeof rawSlideFooterValueSchema>
    > extends true
      ? SlideFooterValue
      : never
  >;

/** Per-slide slide-number override declaration (`false` hides; `true` enables; object overrides style/slot). */
export type SlideNumberValue =
  | boolean
  | {
      readonly style?: FooterPlaceholderStyle | undefined;
      readonly idx?: number | undefined;
      readonly size?: PlaceholderSize | undefined;
      readonly visible?: boolean | undefined;
    };

const rawSlideNumberValueSchema = z.union([
  z.boolean(),
  z
    .strictObject({
      style: rawFooterPlaceholderStyleSchema.optional(),
      idx: z.number().int().nonnegative().optional(),
      size: placeholderSizeSchema.optional(),
      visible: z.boolean().optional(),
    })
    .readonly(),
]);
/** Per-slide slide-number override declaration (`false` hides; `true` enables; object overrides style/slot). */
export const slideNumberValueSchema: z.ZodType<SlideNumberValue> =
  rawSlideNumberValueSchema satisfies z.ZodType<
    AssertExact<
      SlideNumberValue,
      z.infer<typeof rawSlideNumberValueSchema>
    > extends true
      ? SlideNumberValue
      : never
  >;

export type SlideProps = {
  /** Page author id; unique within the deck and written as the stable slide identifier. */
  readonly id: AuthorId;
  /** Optional PowerPoint display name written to `<p:cSld name="...">` without altering the stable `id`. */
  readonly name?: string | undefined;
  readonly layout?: unknown;
  readonly theme?: ThemeScope | undefined;
  readonly colorMapOverride?: ColorMapOverride | undefined;
  readonly themeOverride?: ThemeOverride | undefined;
  readonly style?: SlideStyle | undefined;
  readonly hidden?: boolean | undefined;
  readonly showMasterPlaceholderAnimations?: boolean | undefined;
  readonly showMasterShapes?: boolean | undefined;
  readonly transition?: unknown;
  readonly advanceAfter?: number | undefined;
  readonly background?: SlideBackground | undefined;
  readonly dateTime?: SlideFooterValue | undefined;
  readonly footer?: SlideFooterValue | undefined;
  readonly slideNumber?: SlideNumberValue | undefined;
};

const rawSlidePropsSchema = slideCorePropsObjectSchema
  .extend({
    transition: z.unknown().optional(),
    advanceAfter: z.number().optional(),
    background: rawSlideBackgroundSchema.optional(),
    dateTime: rawSlideFooterValueSchema.optional(),
    footer: rawSlideFooterValueSchema.optional(),
    slideNumber: rawSlideNumberValueSchema.optional(),
  })
  .readonly();
export const slidePropsSchema: z.ZodType<SlideProps> =
  rawSlidePropsSchema satisfies z.ZodType<
    AssertExact<SlideProps, z.infer<typeof rawSlidePropsSchema>> extends true
      ? SlideProps
      : never
  >;

/** Organizes slides into collapsible flat sections in PowerPoint. */
export type SectionProps = {
  readonly name: string;
};

const rawSectionPropsSchema = z
  .strictObject({
    name: z
      .string()
      .refine(
        (value) => value.trim() !== "",
        "Section.name must be a non-empty string",
      )
      .pipe(xml10StringSchema),
  })
  .readonly();
/** Organizes slides into collapsible flat sections in PowerPoint. */
export const sectionPropsSchema: z.ZodType<SectionProps> =
  rawSectionPropsSchema satisfies z.ZodType<
    AssertExact<
      SectionProps,
      z.infer<typeof rawSectionPropsSchema>
    > extends true
      ? SectionProps
      : never
  >;

/** Target declaration interpreted by PowerPoint upon clicking an object; internal targets are resolved by page author id. */
export type LinkTarget =
  | { readonly id: string }
  | { readonly url: string }
  | { readonly action: "home" };

const rawLinkTargetSchema = z.union([
  z.strictObject({ id: z.string() }).readonly(),
  z.strictObject({ url: z.string() }).readonly(),
  z.strictObject({ action: z.literal("home") }).readonly(),
]);
/** Target declaration interpreted by PowerPoint upon clicking an object; internal targets are resolved by page author id. */
export const linkTargetSchema: z.ZodType<LinkTarget> =
  rawLinkTargetSchema satisfies z.ZodType<
    AssertExact<LinkTarget, z.infer<typeof rawLinkTargetSchema>> extends true
      ? LinkTarget
      : never
  >;

/** Validates link target declarations once during capability bind or resolve. */
export function parseLinkTarget(
  value: unknown,
  path?: string,
): LinkTarget | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new CapabilityError(
      "invalid-link",
      "link must specify id, url, or home action",
      path,
    );
  }
  const definedEntries = Object.entries(value).filter(
    ([, target]) => target !== undefined,
  );
  if (definedEntries.length !== 1) {
    throw new CapabilityError(
      "invalid-link",
      "link must specify exactly one target",
      path,
    );
  }
  const parsed = linkTargetSchema.safeParse(Object.fromEntries(definedEntries));
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-link",
      "link supports { id }, { url }, or { action: 'home' }",
      path,
    );
  }
  return parsed.data;
}

const deckAllowedKeys = new Set(Object.keys(deckPropsObjectSchema.shape));

/**
 * Validates Deck declaration props against `deckPropsSchema` and normalizes default dimensions
 * while preserving original omitted vs explicit declarations on the editable node.
 */
export function parseDeckProps(props: Readonly<Record<string, unknown>>): {
  readonly width: number;
  readonly height: number;
  readonly title?: string;
} {
  const unknown = Object.keys(props).find((key) => !deckAllowedKeys.has(key));
  if (unknown !== undefined) {
    throw new BuildError([
      {
        code: "unknown-prop",
        path: "deck",
        prop: unknown,
        field: unknown,
        scope: "props",
        message: `Deck contains unknown parameter ${unknown}`,
      },
    ]);
  }
  const widthResult =
    props["width"] === undefined
      ? { success: true as const, data: 1280 }
      : slideDimensionSchema.safeParse(props["width"]);
  if (!widthResult.success) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: "deck",
        field: "width",
        message: `Deck.width ${widthResult.error.issues[0]?.message ?? "must be a finite positive pixel value"}`,
      },
    ]);
  }
  const heightResult =
    props["height"] === undefined
      ? { success: true as const, data: 720 }
      : slideDimensionSchema.safeParse(props["height"]);
  if (!heightResult.success) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: "deck",
        field: "height",
        message: `Deck.height ${heightResult.error.issues[0]?.message ?? "must be a finite positive pixel value"}`,
      },
    ]);
  }
  const rawTitle = props["title"];
  if (rawTitle !== undefined) {
    if (typeof rawTitle !== "string") {
      throw new BuildError([
        {
          code: "invalid-prop",
          path: "deck",
          field: "title",
          message: "Deck.title must be a string",
        },
      ]);
    }
    const titleResult = xml10StringSchema.safeParse(rawTitle);
    if (!titleResult.success) {
      throw new BuildError([
        {
          code: "invalid-prop",
          path: "deck",
          field: "title",
          message: "Deck.title contains characters not allowed in XML 1.0",
        },
      ]);
    }
  }
  return {
    width: widthResult.data,
    height: heightResult.data,
    ...(typeof rawTitle === "string" ? { title: rawTitle } : {}),
  };
}

/** Validates Section declaration props against `sectionPropsSchema` with localized BuildError diagnostics. */
export function parseSectionProps(
  props: Readonly<Record<string, unknown>>,
  index: number,
): SectionProps {
  const unknown = Object.keys(props).find((key) => key !== "name");
  if (unknown !== undefined) {
    throw new BuildError([
      {
        code: "unknown-prop",
        path: `deck/section[${index + 1}]`,
        prop: unknown,
        field: unknown,
        scope: "props",
        message: `Section does not support property ${unknown}`,
      },
    ]);
  }
  const rawName = props["name"];
  if (typeof rawName !== "string" || rawName.trim() === "") {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/section[${index + 1}]`,
        field: "name",
        message: "Section.name must be a non-empty string",
      },
    ]);
  }
  const parsed = sectionPropsSchema.safeParse(props);
  if (!parsed.success) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/section[${rawName}]`,
        field: "name",
        message: "Section.name contains characters not allowed in XML 1.0",
      },
    ]);
  }
  return parsed.data;
}

/** Validates core Slide props against `slideCorePropsSchema` while respecting registered slide capability extensions. */
export function parseSlideCoreProps(
  props: Readonly<Record<string, unknown>>,
  index: number,
  registeredSlideProps: readonly string[],
): {
  readonly id: AuthorId;
  readonly name?: string;
  readonly hidden?: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly showMasterShapes?: boolean;
  readonly colorMapOverride?: ColorMapOverride;
  readonly themeOverride?: ThemeOverride;
} {
  const allowed = new Set([
    "id",
    "name",
    "layout",
    "placeholders",
    "theme",
    "colorMapOverride",
    "themeOverride",
    "style",
    "hidden",
    "showMasterPlaceholderAnimations",
    "showMasterShapes",
    ...registeredSlideProps,
  ]);
  const unknown = Object.keys(props).find((key) => !allowed.has(key));
  const authorId = readAuthorId(props["id"], `deck/slide[${index + 1}]`);
  if (unknown !== undefined) {
    const label = authorId ?? index + 1;
    throw new BuildError([
      {
        code: "missing-capability",
        path: `deck/slide[${label}]`,
        field: unknown,
        capability: unknown,
        message: `No capability module registered for Slide.${unknown}`,
      },
    ]);
  }
  if (authorId === undefined) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/slide[${index + 1}]`,
        field: "id",
        message: "Slide must declare id as a non-empty string",
      },
    ]);
  }
  if (!xml10StringSchema.safeParse(authorId).success) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/slide[${index + 1}]`,
        field: "id",
        message: "Slide.id contains characters not allowed in XML 1.0",
      },
    ]);
  }
  const rawName = props["name"];
  if (rawName !== undefined && typeof rawName !== "string") {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/slide[${authorId}]`,
        field: "name",
        message: "Slide.name must be a string",
      },
    ]);
  }
  if (rawName !== undefined && !xml10StringSchema.safeParse(rawName).success) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/slide[${authorId}]`,
        field: "name",
        message: "Slide.name contains characters not allowed in XML 1.0",
      },
    ]);
  }
  const hidden = props["hidden"];
  if (
    hidden !== undefined &&
    !slideCorePropsObjectSchema.shape.hidden.safeParse(hidden).success
  ) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/slide[${authorId}]`,
        field: "hidden",
        message: "Slide.hidden must be a boolean",
      },
    ]);
  }
  const showMasterShapes = props["showMasterShapes"];
  if (
    showMasterShapes !== undefined &&
    !slideCorePropsObjectSchema.shape.showMasterShapes.safeParse(
      showMasterShapes,
    ).success
  ) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/slide[${authorId}]`,
        field: "showMasterShapes",
        message: "Slide.showMasterShapes must be a boolean",
      },
    ]);
  }
  const rawStyle = props["style"];
  if (
    typeof rawStyle === "object" &&
    rawStyle !== null &&
    !Array.isArray(rawStyle) &&
    Object.hasOwn(rawStyle, "zIndex") &&
    (rawStyle as Record<string, unknown>)["zIndex"] !== undefined
  ) {
    throw new BuildError([
      {
        code: "invalid-style",
        path: `deck/slide[${authorId}]`,
        field: "style.zIndex",
        message: "Slide.style does not support zIndex",
      },
    ]);
  }
  const rawColorMapOverride = props["colorMapOverride"];
  let colorMapOverride: ColorMapOverride | undefined;
  if (rawColorMapOverride !== undefined) {
    const parsedMap = colorMapOverrideSchema.safeParse(rawColorMapOverride);
    if (!parsedMap.success) {
      throw new BuildError([
        {
          code: "invalid-prop",
          path: `deck/slide[${authorId}]`,
          field: "colorMapOverride",
          message:
            parsedMap.error.issues[0]?.message ??
            "Slide.colorMapOverride must be 'master' or a 12-slot color map",
        },
      ]);
    }
    colorMapOverride = parsedMap.data;
  }
  const rawThemeOverride = props["themeOverride"];
  let themeOverride: ThemeOverride | undefined;
  if (rawThemeOverride !== undefined) {
    const parsedOverride = themeOverrideSchema.safeParse(rawThemeOverride);
    if (!parsedOverride.success) {
      throw new BuildError([
        {
          code: "invalid-prop",
          path: `deck/slide[${authorId}]`,
          field: "themeOverride",
          message:
            parsedOverride.error.issues[0]?.message ??
            "Slide.themeOverride is invalid",
        },
      ]);
    }
    themeOverride = parsedOverride.data;
  }
  const showMasterPlaceholderAnimations =
    props["showMasterPlaceholderAnimations"];
  if (
    showMasterPlaceholderAnimations !== undefined &&
    !slideCorePropsObjectSchema.shape.showMasterPlaceholderAnimations.safeParse(
      showMasterPlaceholderAnimations,
    ).success
  ) {
    throw new BuildError([
      {
        code: "invalid-prop",
        path: `deck/slide[${authorId}]`,
        field: "showMasterPlaceholderAnimations",
        message: "Slide.showMasterPlaceholderAnimations must be a boolean",
      },
    ]);
  }
  return {
    id: authorId,
    ...(rawName === undefined ? {} : { name: rawName }),
    ...(hidden === undefined ? {} : { hidden: hidden as boolean }),
    ...(showMasterPlaceholderAnimations === undefined
      ? {}
      : {
          showMasterPlaceholderAnimations:
            showMasterPlaceholderAnimations as boolean,
        }),
    ...(showMasterShapes === undefined
      ? {}
      : { showMasterShapes: showMasterShapes as boolean }),
    ...(colorMapOverride === undefined ? {} : { colorMapOverride }),
    ...(themeOverride === undefined ? {} : { themeOverride }),
  };
}
