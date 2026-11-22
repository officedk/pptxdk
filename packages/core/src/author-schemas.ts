import { z } from "zod";

import { authorIdSchema, readAuthorId, type AuthorId } from "#/author-ids.ts";
import { BuildError, CapabilityError } from "#/diagnostics.ts";
import { slideLayoutStyleShape } from "#/layout-schemas.ts";
import {
  colorMapOverrideSchema,
  colorValueSchema,
  linearGradientSchema,
  radialGradientSchema,
  textStyleSchema,
  textStyleShape,
  themeInputSchema,
  themeOverrideSchema,
  themeScopeSchema,
  uint8ArraySchema,
  type ColorMapOverride,
  type ThemeOverride,
} from "#/theme.ts";

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

/** Slide background can use theme colors, gradients supported by the shared visual writer, or formal image resource inputs. */
export const slideBackgroundSchema = z.union([
  colorValueSchema,
  linearGradientSchema,
  radialGradientSchema,
  z
    .strictObject({
      image: z.union([z.string(), z.instanceof(URL), uint8ArraySchema]),
    })
    .readonly(),
]);
export type SlideBackground = z.infer<typeof slideBackgroundSchema>;

const deckPropsObjectSchema = z.strictObject({
  width: slideDimensionSchema.optional(),
  height: slideDimensionSchema.optional(),
  theme: themeInputSchema.optional(),
  style: textStyleSchema.optional(),
  title: xml10StringSchema.optional(),
});
export const deckPropsSchema = deckPropsObjectSchema.readonly();
export type DeckProps = z.infer<typeof deckPropsSchema>;

export const slideStyleSchema = z
  .strictObject({
    ...textStyleShape,
    ...slideLayoutStyleShape,
  })
  .readonly();
export type SlideStyle = z.infer<typeof slideStyleSchema>;

const slideCorePropsObjectSchema = z.strictObject({
  /** Page author id; unique within the deck and written as the stable slide identifier. */
  id: authorIdSchema,
  layout: z.unknown().optional(),
  theme: themeScopeSchema.optional(),
  colorMapOverride: colorMapOverrideSchema.optional(),
  themeOverride: themeOverrideSchema.optional(),
  style: slideStyleSchema.optional(),
  hidden: z.boolean().optional(),
  showMasterPlaceholderAnimations: z.boolean().optional(),
});
export const slideCorePropsSchema = slideCorePropsObjectSchema.readonly();

export const slidePropsSchema = slideCorePropsObjectSchema
  .extend({
    transition: z.unknown().optional(),
    advanceAfter: z.number().optional(),
    background: slideBackgroundSchema.optional(),
  })
  .readonly();
export type SlideProps = z.infer<typeof slidePropsSchema>;

/** Organizes slides into collapsible flat sections in PowerPoint. */
export const sectionPropsSchema = z
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
export type SectionProps = z.infer<typeof sectionPropsSchema>;

/** Target declaration interpreted by PowerPoint upon clicking an object; internal targets are resolved by page author id. */
export const linkTargetSchema = z.union([
  z.strictObject({ id: z.string() }).readonly(),
  z.strictObject({ url: z.string() }).readonly(),
  z.strictObject({ action: z.literal("home") }).readonly(),
]);
export type LinkTarget = z.infer<typeof linkTargetSchema>;

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
        severity: "error",
        path: "deck",
        code: "unknown-parameter",
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
        severity: "error",
        path: "deck",
        code: "invalid-dimension",
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
        severity: "error",
        path: "deck",
        code: "invalid-dimension",
        message: `Deck.height ${heightResult.error.issues[0]?.message ?? "must be a finite positive pixel value"}`,
      },
    ]);
  }
  const rawTitle = props["title"];
  if (rawTitle !== undefined) {
    if (typeof rawTitle !== "string") {
      throw new BuildError([
        {
          severity: "error",
          path: "deck",
          code: "invalid-title",
          message: "Deck.title must be a string",
        },
      ]);
    }
    const titleResult = xml10StringSchema.safeParse(rawTitle);
    if (!titleResult.success) {
      throw new BuildError([
        {
          severity: "error",
          path: "deck",
          code: "invalid-title",
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
        severity: "error",
        path: `deck/section[${index + 1}]`,
        code: "unknown-section-property",
        message: `Section does not support property ${unknown}`,
      },
    ]);
  }
  const rawName = props["name"];
  if (typeof rawName !== "string" || rawName.trim() === "") {
    throw new BuildError([
      {
        severity: "error",
        path: `deck/section[${index + 1}]`,
        code: "invalid-section-name",
        message: "Section.name must be a non-empty string",
      },
    ]);
  }
  const parsed = sectionPropsSchema.safeParse(props);
  if (!parsed.success) {
    throw new BuildError([
      {
        severity: "error",
        path: `deck/section[${rawName}]`,
        code: "invalid-xml-character",
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
  readonly hidden?: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly colorMapOverride?: ColorMapOverride;
  readonly themeOverride?: ThemeOverride;
} {
  const allowed = new Set([
    "id",
    "layout",
    "placeholders",
    "theme",
    "colorMapOverride",
    "themeOverride",
    "style",
    "hidden",
    "showMasterPlaceholderAnimations",
    ...registeredSlideProps,
  ]);
  const unknown = Object.keys(props).find((key) => !allowed.has(key));
  const authorId = readAuthorId(props["id"], `deck/slide[${index + 1}]`);
  if (unknown !== undefined) {
    const label = authorId ?? index + 1;
    throw new BuildError([
      {
        severity: "error",
        path: `deck/slide[${label}]`,
        code: "missing-capability",
        message: `No capability module registered for Slide.${unknown}`,
      },
    ]);
  }
  if (authorId === undefined) {
    throw new BuildError([
      {
        severity: "error",
        path: `deck/slide[${index + 1}]`,
        code: "invalid-author-id",
        message: "Slide must declare id as a non-empty string",
      },
    ]);
  }
  if (!xml10StringSchema.safeParse(authorId).success) {
    throw new BuildError([
      {
        severity: "error",
        path: `deck/slide[${index + 1}]`,
        code: "invalid-xml-character",
        message: "Slide.id contains characters not allowed in XML 1.0",
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
        severity: "error",
        path: `deck/slide[${authorId}]`,
        code: "invalid-hidden-slide",
        message: "Slide.hidden must be a boolean",
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
        severity: "error",
        path: `deck/slide[${authorId}]`,
        code: "unsupported-layout-style",
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
          severity: "error",
          path: `deck/slide[${authorId}]`,
          code: "invalid-color-map-override",
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
          severity: "error",
          path: `deck/slide[${authorId}]`,
          code: "invalid-theme-override",
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
        severity: "error",
        path: `deck/slide[${authorId}]`,
        code: "invalid-slide-prop",
        message: "Slide.showMasterPlaceholderAnimations must be a boolean",
      },
    ]);
  }
  return {
    id: authorId,
    ...(hidden === undefined ? {} : { hidden: hidden as boolean }),
    ...(showMasterPlaceholderAnimations === undefined
      ? {}
      : {
          showMasterPlaceholderAnimations:
            showMasterPlaceholderAnimations as boolean,
        }),
    ...(colorMapOverride === undefined ? {} : { colorMapOverride }),
    ...(themeOverride === undefined ? {} : { themeOverride }),
  };
}
