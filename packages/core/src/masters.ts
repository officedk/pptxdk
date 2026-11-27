import { z } from "zod";

import {
  footerPlaceholderStyleSchema,
  placeholderSizeSchema,
  slideBackgroundSchema,
  slideStyleSchema,
  xml10StringSchema,
  type FooterPlaceholderStyle,
  type PlaceholderSize,
  type SlideBackground,
  type SlideStyle,
} from "#/author-schemas.ts";
import { BuildError, zodIssueToBuildDiagnostic } from "#/diagnostics.ts";
import {
  masterPlaceholderMapSchema,
  masterTextStylesSchema,
  placeholderMapSchema,
  placeholderRoleSchema,
  type ContentPlaceholderRole,
  type EmptyPlaceholderMap,
  type MasterPlaceholderMap,
  type MasterTextStyles,
  type PlaceholderMap,
} from "#/placeholders.ts";
import {
  colorMapOverrideSchema,
  colorMapSchema,
  themeOverrideSchema,
  type ColorMap,
  type ColorMapOverride,
  type Theme,
  type ThemeInput,
  type ThemeOverride,
} from "#/theme.ts";
import type { Rect, ResolvedObject } from "#/types.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

/** OOXML `<p:hf>` header/footer visibility flags on SlideMaster and SlideLayout (ECMA-376 §19.3.1.24). */
export type HeaderFooter = {
  readonly dateTime?: boolean | undefined;
  readonly footer?: boolean | undefined;
  readonly header?: boolean | undefined;
  readonly slideNumber?: boolean | undefined;
};

const rawHeaderFooterSchema = z
  .strictObject({
    dateTime: z.boolean().optional(),
    footer: z.boolean().optional(),
    header: z.boolean().optional(),
    slideNumber: z.boolean().optional(),
  })
  .readonly();
/** OOXML `<p:hf>` header/footer visibility flags on SlideMaster and SlideLayout (ECMA-376 §19.3.1.24). */
export const headerFooterSchema: z.ZodType<HeaderFooter> =
  rawHeaderFooterSchema satisfies z.ZodType<
    AssertExact<
      HeaderFooter,
      z.infer<typeof rawHeaderFooterSchema>
    > extends true
      ? HeaderFooter
      : never
  >;

/** Geometry, placeholder metadata, and default text/style for a `dt`, `ftr`, or `sldNum` slot. */
export type FooterPlaceholderSpec = {
  readonly idx?: number | undefined;
  readonly size?: PlaceholderSize | undefined;
  readonly text?: string | undefined;
  readonly style?: FooterPlaceholderStyle | undefined;
  readonly visible?: boolean | undefined;
};

const rawFooterPlaceholderSpecSchema = z
  .strictObject({
    idx: z.number().int().nonnegative().optional(),
    size: placeholderSizeSchema.optional(),
    text: xml10StringSchema.optional(),
    style: footerPlaceholderStyleSchema.optional(),
    visible: z.boolean().optional(),
  })
  .readonly();
/** Geometry, placeholder metadata, and default text/style for a `dt`, `ftr`, or `sldNum` slot. */
export const footerPlaceholderSpecSchema: z.ZodType<FooterPlaceholderSpec> =
  rawFooterPlaceholderSpecSchema satisfies z.ZodType<
    AssertExact<
      FooterPlaceholderSpec,
      z.infer<typeof rawFooterPlaceholderSpecSchema>
    > extends true
      ? FooterPlaceholderSpec
      : never
  >;

/** Collection of date, footer, and slide-number metadata placeholder definitions on a master or layout. */
export type FooterPlaceholders = {
  readonly dateTime?: FooterPlaceholderSpec | undefined;
  readonly footer?: FooterPlaceholderSpec | undefined;
  readonly slideNumber?: FooterPlaceholderSpec | undefined;
};

const rawFooterPlaceholdersSchema = z
  .strictObject({
    dateTime: rawFooterPlaceholderSpecSchema.optional(),
    footer: rawFooterPlaceholderSpecSchema.optional(),
    slideNumber: rawFooterPlaceholderSpecSchema.optional(),
  })
  .readonly();
/** Collection of date, footer, and slide-number metadata placeholder definitions on a master or layout. */
export const footerPlaceholdersSchema: z.ZodType<FooterPlaceholders> =
  rawFooterPlaceholdersSchema satisfies z.ZodType<
    AssertExact<
      FooterPlaceholders,
      z.infer<typeof rawFooterPlaceholdersSchema>
    > extends true
      ? FooterPlaceholders
      : never
  >;

/** Official OOXML ST_SlideLayoutType enumeration values (ECMA-376 §19.7.15). */
export const slideLayoutTypeValues = [
  "title",
  "tx",
  "twoColTx",
  "tbl",
  "txAndChart",
  "chartAndTx",
  "dgm",
  "chart",
  "txAndClipArt",
  "clipArtAndTx",
  "titleOnly",
  "blank",
  "txAndObj",
  "objAndTx",
  "objOnly",
  "obj",
  "txAndMedia",
  "mediaAndTx",
  "objOverTx",
  "txOverObj",
  "txAndTwoObj",
  "twoObjAndTx",
  "twoObjOverTx",
  "fourObj",
  "vertTx",
  "clipArtAndVertTx",
  "vertTitleAndTx",
  "vertTitleAndTxOverChart",
  "twoObj",
  "objAndTwoObj",
  "twoObjAndObj",
  "cust",
  "secHead",
  "twoTxTwoObj",
  "objTx",
  "picTx",
] as const;

export type SlideLayoutType = (typeof slideLayoutTypeValues)[number];
const rawSlideLayoutTypeSchema = z.enum(slideLayoutTypeValues);
export const slideLayoutTypeSchema: z.ZodType<SlideLayoutType> =
  rawSlideLayoutTypeSchema satisfies z.ZodType<
    AssertExact<
      SlideLayoutType,
      z.infer<typeof rawSlideLayoutTypeSchema>
    > extends true
      ? SlideLayoutType
      : never
  >;

export type SlideMasterDataProps = {
  readonly id?: string | undefined;
  readonly name?: string | undefined;
  readonly preserve?: boolean | undefined;
  readonly placeholders?: MasterPlaceholderMap | undefined;
  readonly textStyles?: MasterTextStyles | undefined;
  readonly colorMap?: ColorMap | undefined;
  readonly style?: SlideStyle | undefined;
  readonly background?: SlideBackground | undefined;
  readonly headerFooter?: HeaderFooter | undefined;
  readonly footerPlaceholders?: FooterPlaceholders | undefined;
  readonly transition?: unknown;
  readonly advanceAfter?: number | undefined;
  readonly motion?: unknown;
};

const rawSlideMasterDataPropsSchema = z
  .strictObject({
    id: z.string().trim().min(1).optional(),
    name: z.string().optional(),
    preserve: z.boolean().optional(),
    placeholders: masterPlaceholderMapSchema.optional(),
    textStyles: masterTextStylesSchema.optional(),
    colorMap: colorMapSchema.optional(),
    style: slideStyleSchema.optional(),
    background: slideBackgroundSchema.optional(),
    headerFooter: rawHeaderFooterSchema.optional(),
    footerPlaceholders: rawFooterPlaceholdersSchema.optional(),
    transition: z.unknown().optional(),
    advanceAfter: z.number().optional(),
    motion: z.unknown().optional(),
  })
  .readonly();
export const slideMasterDataPropsSchema: z.ZodType<SlideMasterDataProps> =
  rawSlideMasterDataPropsSchema satisfies z.ZodType<
    AssertExact<
      SlideMasterDataProps,
      z.infer<typeof rawSlideMasterDataPropsSchema>
    > extends true
      ? SlideMasterDataProps
      : never
  >;

export type SlideMasterProps = SlideMasterDataProps & {
  readonly theme?: ThemeInput | Theme | undefined;
  readonly style?: SlideStyle | undefined;
  readonly placeholders?: MasterPlaceholderMap | undefined;
  readonly textStyles?: MasterTextStyles | undefined;
};

export type SlideLayoutDataProps<
  P extends PlaceholderMap = EmptyPlaceholderMap,
> = {
  readonly id?: string | undefined;
  readonly name?: string | undefined;
  readonly type?: SlideLayoutType | undefined;
  readonly preserve?: boolean | undefined;
  readonly matchingName?: string | undefined;
  readonly userDrawn?: boolean | undefined;
  readonly showMasterShapes?: boolean | undefined;
  readonly showMasterPlaceholderAnimations?: boolean | undefined;
  readonly placeholders?: P | undefined;
  readonly colorMapOverride?: ColorMapOverride | undefined;
  readonly themeOverride?: ThemeOverride | undefined;
  readonly style?: SlideStyle | undefined;
  readonly background?: SlideBackground | undefined;
  readonly headerFooter?: HeaderFooter | undefined;
  readonly footerPlaceholders?: FooterPlaceholders | undefined;
  readonly transition?: unknown;
  readonly advanceAfter?: number | undefined;
  readonly motion?: unknown;
};

const rawSlideLayoutDataPropsSchema = z
  .strictObject({
    id: z.string().trim().min(1).optional(),
    name: z.string().optional(),
    type: rawSlideLayoutTypeSchema.optional(),
    preserve: z.boolean().optional(),
    matchingName: z.string().optional(),
    userDrawn: z.boolean().optional(),
    showMasterShapes: z.boolean().optional(),
    showMasterPlaceholderAnimations: z.boolean().optional(),
    placeholders: placeholderMapSchema.optional(),
    colorMapOverride: colorMapOverrideSchema.optional(),
    themeOverride: themeOverrideSchema.optional(),
    style: slideStyleSchema.optional(),
    background: slideBackgroundSchema.optional(),
    headerFooter: rawHeaderFooterSchema.optional(),
    footerPlaceholders: rawFooterPlaceholdersSchema.optional(),
    transition: z.unknown().optional(),
    advanceAfter: z.number().optional(),
    motion: z.unknown().optional(),
  })
  .readonly();
export const slideLayoutDataPropsSchema: z.ZodType<
  SlideLayoutDataProps<PlaceholderMap>
> = rawSlideLayoutDataPropsSchema satisfies z.ZodType<
  AssertExact<
    SlideLayoutDataProps<PlaceholderMap>,
    z.infer<typeof rawSlideLayoutDataPropsSchema>
  > extends true
    ? SlideLayoutDataProps<PlaceholderMap>
    : never
>;

const rectObservationSchema = z
  .object({
    x: z.number(),
    y: z.number(),
    width: z.number(),
    height: z.number(),
  })
  .readonly();

export type ResolvedPlaceholder = {
  readonly key: string;
  readonly type: ContentPlaceholderRole;
  readonly idx?: number | undefined;
  readonly size?: PlaceholderSize | undefined;
  readonly box?: Rect | undefined;
  readonly style?: Readonly<Record<string, unknown>> | undefined;
};

const rawResolvedPlaceholderSchema = z
  .object({
    key: z.string().min(1),
    type: placeholderRoleSchema,
    idx: z.number().int().nonnegative().optional(),
    size: placeholderSizeSchema.optional(),
    box: rectObservationSchema.optional(),
    style: z.record(z.string(), z.unknown()).readonly().optional(),
  })
  .readonly();
export const resolvedPlaceholderSchema: z.ZodType<ResolvedPlaceholder> =
  rawResolvedPlaceholderSchema satisfies z.ZodType<
    AssertExact<
      ResolvedPlaceholder,
      z.infer<typeof rawResolvedPlaceholderSchema>
    > extends true
      ? ResolvedPlaceholder
      : never
  >;

export type ResolvedSlideLayout = {
  readonly id: string;
  readonly name?: string | undefined;
  readonly type?: SlideLayoutType | undefined;
  readonly preserve?: boolean | undefined;
  readonly matchingName?: string | undefined;
  readonly userDrawn?: boolean | undefined;
  readonly showMasterShapes?: boolean | undefined;
  readonly showMasterPlaceholderAnimations?: boolean | undefined;
  readonly colorMapOverride?: ColorMapOverride | undefined;
  readonly themeOverride?: ThemeOverride | undefined;
  readonly style?: SlideStyle | undefined;
  readonly background?: SlideBackground | undefined;
  readonly headerFooter?: HeaderFooter | undefined;
  readonly objects: readonly ResolvedObject[];
  readonly placeholders: readonly ResolvedPlaceholder[];
  readonly capabilities?: Readonly<Record<string, unknown>> | undefined;
};

const rawResolvedSlideLayoutSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().optional(),
    type: rawSlideLayoutTypeSchema.optional(),
    preserve: z.boolean().optional(),
    matchingName: z.string().optional(),
    userDrawn: z.boolean().optional(),
    showMasterShapes: z.boolean().optional(),
    showMasterPlaceholderAnimations: z.boolean().optional(),
    colorMapOverride: colorMapOverrideSchema.optional(),
    themeOverride: themeOverrideSchema.optional(),
    style: slideStyleSchema.optional(),
    background: slideBackgroundSchema.optional(),
    headerFooter: rawHeaderFooterSchema.optional(),
    objects: z.array(z.custom<ResolvedObject>()).readonly(),
    placeholders: z.array(rawResolvedPlaceholderSchema).readonly(),
    capabilities: z.record(z.string(), z.unknown()).readonly().optional(),
  })
  .readonly();
export const resolvedSlideLayoutSchema: z.ZodType<ResolvedSlideLayout> =
  rawResolvedSlideLayoutSchema satisfies z.ZodType<
    AssertExact<
      ResolvedSlideLayout,
      z.infer<typeof rawResolvedSlideLayoutSchema>
    > extends true
      ? ResolvedSlideLayout
      : never
  >;

export type ResolvedMaster = {
  readonly id: string;
  readonly name?: string | undefined;
  readonly preserve?: boolean | undefined;
  readonly colorMap?: ColorMap | undefined;
  readonly headerFooter?: HeaderFooter | undefined;
  readonly objects: readonly ResolvedObject[];
  readonly placeholders: readonly ResolvedPlaceholder[];
  readonly capabilities?: Readonly<Record<string, unknown>> | undefined;
  readonly layouts: readonly ResolvedSlideLayout[];
};

const rawResolvedMasterSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().optional(),
    preserve: z.boolean().optional(),
    colorMap: colorMapSchema.optional(),
    headerFooter: rawHeaderFooterSchema.optional(),
    objects: z.array(z.custom<ResolvedObject>()).readonly(),
    placeholders: z.array(rawResolvedPlaceholderSchema).readonly(),
    capabilities: z.record(z.string(), z.unknown()).readonly().optional(),
    layouts: z.array(rawResolvedSlideLayoutSchema).readonly(),
  })
  .readonly();
export const resolvedMasterSchema: z.ZodType<ResolvedMaster> =
  rawResolvedMasterSchema satisfies z.ZodType<
    AssertExact<
      ResolvedMaster,
      z.infer<typeof rawResolvedMasterSchema>
    > extends true
      ? ResolvedMaster
      : never
  >;

const slideLayoutDefinitionSymbol = Symbol.for(
  "@pptxdk/core/slide-layout-definition",
);

/** Stateless, schema-validated slide layout specification shareable across JSX and imperative decks. */
export type SlideLayoutDefinition<P extends PlaceholderMap = PlaceholderMap> = {
  readonly [slideLayoutDefinitionSymbol]: true;
  readonly props: SlideLayoutDataProps<P>;
  readonly id?: string | undefined;
  readonly name?: string | undefined;
  readonly type?: SlideLayoutType | undefined;
  readonly preserve?: boolean | undefined;
  readonly matchingName?: string | undefined;
  readonly userDrawn?: boolean | undefined;
  readonly showMasterShapes?: boolean | undefined;
  readonly showMasterPlaceholderAnimations?: boolean | undefined;
  readonly placeholders?: P | undefined;
  readonly colorMapOverride?: ColorMapOverride | undefined;
  readonly themeOverride?: ThemeOverride | undefined;
  readonly style?: SlideStyle | undefined;
  readonly background?: SlideBackground | undefined;
  readonly headerFooter?: HeaderFooter | undefined;
  readonly footerPlaceholders?: FooterPlaceholders | undefined;
  readonly transition?: unknown;
  readonly advanceAfter?: number | undefined;
  readonly motion?: unknown;
};

export type SlideLayoutProps<P extends PlaceholderMap = PlaceholderMap> =
  SlideLayoutDataProps<P> & {
    readonly layout?: SlideLayoutDefinition<P> | undefined;
  };

/** Checks whether a runtime value was produced by `defineSlideLayout`. */
export function isSlideLayoutDefinition(
  value: unknown,
): value is SlideLayoutDefinition {
  return (
    typeof value === "object" &&
    value !== null &&
    slideLayoutDefinitionSymbol in value
  );
}

/**
 * Defines a reusable native slide layout specification that can be mounted into a `SlideMaster`
 * across both JSX (`<SlideLayout layout={def} />`) and imperative (`slideLayout(def)`) entries.
 */
export function defineSlideLayout<
  const P extends PlaceholderMap = EmptyPlaceholderMap,
>(spec: SlideLayoutDataProps<P>): SlideLayoutDefinition<P> {
  const parsed = slideLayoutDataPropsSchema.safeParse(spec);
  if (!parsed.success) {
    const first = parsed.error.issues[0];
    throw new BuildError([
      zodIssueToBuildDiagnostic(
        "slide-layout",
        "invalid-slide-layout",
        first?.message ?? "Invalid slide layout specification",
        first,
      ),
    ]);
  }
  const normalized: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(parsed.data)) {
    if (value !== undefined) {
      normalized[key] = value;
    }
  }
  const frozenProps = Object.freeze(normalized) as SlideLayoutDataProps<P>;
  return Object.freeze({
    [slideLayoutDefinitionSymbol]: true as const,
    props: frozenProps,
    ...(frozenProps.id === undefined ? {} : { id: frozenProps.id }),
    ...(frozenProps.name === undefined ? {} : { name: frozenProps.name }),
    ...(frozenProps.type === undefined ? {} : { type: frozenProps.type }),
    ...(frozenProps.preserve === undefined
      ? {}
      : { preserve: frozenProps.preserve }),
    ...(frozenProps.matchingName === undefined
      ? {}
      : { matchingName: frozenProps.matchingName }),
    ...(frozenProps.userDrawn === undefined
      ? {}
      : { userDrawn: frozenProps.userDrawn }),
    ...(frozenProps.showMasterShapes === undefined
      ? {}
      : { showMasterShapes: frozenProps.showMasterShapes }),
    ...(frozenProps.showMasterPlaceholderAnimations === undefined
      ? {}
      : {
          showMasterPlaceholderAnimations:
            frozenProps.showMasterPlaceholderAnimations,
        }),
    ...(frozenProps.placeholders === undefined
      ? {}
      : { placeholders: frozenProps.placeholders }),
    ...(frozenProps.colorMapOverride === undefined
      ? {}
      : { colorMapOverride: frozenProps.colorMapOverride }),
    ...(frozenProps.themeOverride === undefined
      ? {}
      : { themeOverride: frozenProps.themeOverride }),
    ...(frozenProps.style === undefined ? {} : { style: frozenProps.style }),
    ...(frozenProps.background === undefined
      ? {}
      : { background: frozenProps.background }),
    ...(frozenProps.headerFooter === undefined
      ? {}
      : { headerFooter: frozenProps.headerFooter }),
    ...(frozenProps.footerPlaceholders === undefined
      ? {}
      : { footerPlaceholders: frozenProps.footerPlaceholders }),
    ...(frozenProps.transition === undefined
      ? {}
      : { transition: frozenProps.transition }),
    ...(frozenProps.advanceAfter === undefined
      ? {}
      : { advanceAfter: frozenProps.advanceAfter }),
    ...(frozenProps.motion === undefined ? {} : { motion: frozenProps.motion }),
  });
}
