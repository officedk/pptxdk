import { z } from "zod";

import {
  placeholderSizeSchema,
  slideBackgroundSchema,
  xml10StringSchema,
  type SlideBackground,
} from "#/author-schemas.ts";
import { CapabilityError } from "#/diagnostics.ts";
import {
  masterPlaceholderMapSchema,
  masterTextStylesSchema,
  placeholderMapSchema,
  placeholderRoleSchema,
  placeholderSzSchema,
  type EmptyPlaceholderMap,
  type MasterPlaceholderMap,
  type MasterTextStyles,
  type PlaceholderMap,
} from "#/placeholders.ts";
import {
  colorMapOverrideSchema,
  colorMapSchema,
  textStyleSchema,
  themeOverrideSchema,
  type ColorMapOverride,
  type TextStyle,
  type Theme,
  type ThemeInput,
  type ThemeOverride,
} from "#/theme.ts";
import type { ResolvedObject } from "#/types.ts";

/** OOXML `<p:hf>` header/footer visibility flags on SlideMaster and SlideLayout (ECMA-376 §19.3.1.24). */
export const headerFooterSchema = z
  .strictObject({
    dateTime: z.boolean().optional(),
    footer: z.boolean().optional(),
    header: z.boolean().optional(),
    slideNumber: z.boolean().optional(),
  })
  .readonly();
export type HeaderFooter = z.infer<typeof headerFooterSchema>;

/** Geometry, placeholder metadata, and default text/style for a `dt`, `ftr`, or `sldNum` slot. */
export const footerPlaceholderSpecSchema = z
  .strictObject({
    x: z.number().finite().optional(),
    y: z.number().finite().optional(),
    width: z.number().finite().positive().optional(),
    height: z.number().finite().positive().optional(),
    idx: z.number().int().nonnegative().optional(),
    size: placeholderSizeSchema.optional(),
    text: xml10StringSchema.optional(),
    style: textStyleSchema.optional(),
    visible: z.boolean().optional(),
  })
  .readonly();
export type FooterPlaceholderSpec = z.infer<typeof footerPlaceholderSpecSchema>;

/** Collection of date, footer, and slide-number metadata placeholder definitions on a master or layout. */
export const footerPlaceholdersSchema = z
  .strictObject({
    dateTime: footerPlaceholderSpecSchema.optional(),
    footer: footerPlaceholderSpecSchema.optional(),
    slideNumber: footerPlaceholderSpecSchema.optional(),
  })
  .readonly();
export type FooterPlaceholders = z.infer<typeof footerPlaceholdersSchema>;

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

export const slideLayoutTypeSchema = z.enum(slideLayoutTypeValues);
export type SlideLayoutType = z.infer<typeof slideLayoutTypeSchema>;

export const slideMasterDataPropsSchema = z
  .strictObject({
    id: z.string().trim().min(1).optional(),
    name: z.string().optional(),
    preserve: z.boolean().optional(),
    placeholders: masterPlaceholderMapSchema.optional(),
    textStyles: masterTextStylesSchema.optional(),
    colorMap: colorMapSchema.optional(),
    background: slideBackgroundSchema.optional(),
    headerFooter: headerFooterSchema.optional(),
    footerPlaceholders: footerPlaceholdersSchema.optional(),
    transition: z.unknown().optional(),
    advanceAfter: z.number().optional(),
    motion: z.unknown().optional(),
  })
  .readonly();
export type SlideMasterDataProps = z.infer<typeof slideMasterDataPropsSchema>;

export type SlideMasterProps = SlideMasterDataProps & {
  readonly theme?: ThemeInput | Theme | undefined;
  readonly style?: TextStyle | undefined;
  readonly placeholders?: MasterPlaceholderMap | undefined;
  readonly textStyles?: MasterTextStyles | undefined;
};

export const slideLayoutDataPropsSchema = z
  .strictObject({
    id: z.string().trim().min(1).optional(),
    name: z.string().optional(),
    type: slideLayoutTypeSchema.optional(),
    preserve: z.boolean().optional(),
    matchingName: z.string().optional(),
    userDrawn: z.boolean().optional(),
    showMasterShapes: z.boolean().optional(),
    showMasterPlaceholderAnimations: z.boolean().optional(),
    placeholders: placeholderMapSchema.optional(),
    colorMapOverride: colorMapOverrideSchema.optional(),
    themeOverride: themeOverrideSchema.optional(),
    background: slideBackgroundSchema.optional(),
    headerFooter: headerFooterSchema.optional(),
    footerPlaceholders: footerPlaceholdersSchema.optional(),
    transition: z.unknown().optional(),
    advanceAfter: z.number().optional(),
    motion: z.unknown().optional(),
  })
  .readonly();

type BaseSlideLayoutDataProps = Omit<
  z.infer<typeof slideLayoutDataPropsSchema>,
  "placeholders"
>;

export type SlideLayoutDataProps<
  P extends PlaceholderMap = EmptyPlaceholderMap,
> = BaseSlideLayoutDataProps & {
  readonly placeholders?: P | undefined;
};

const rectObservationSchema = z
  .object({
    x: z.number(),
    y: z.number(),
    width: z.number(),
    height: z.number(),
  })
  .readonly();

export const resolvedPlaceholderSchema = z
  .object({
    key: z.string().min(1),
    type: placeholderRoleSchema,
    idx: z.number().int().nonnegative().optional(),
    size: placeholderSzSchema.optional(),
    box: rectObservationSchema.optional(),
    style: z.record(z.string(), z.unknown()).readonly().optional(),
  })
  .readonly();
export type ResolvedPlaceholder = z.infer<typeof resolvedPlaceholderSchema>;

export const resolvedSlideLayoutSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().optional(),
    type: slideLayoutTypeSchema.optional(),
    preserve: z.boolean().optional(),
    matchingName: z.string().optional(),
    userDrawn: z.boolean().optional(),
    showMasterShapes: z.boolean().optional(),
    showMasterPlaceholderAnimations: z.boolean().optional(),
    colorMapOverride: colorMapOverrideSchema.optional(),
    themeOverride: themeOverrideSchema.optional(),
    background: slideBackgroundSchema.optional(),
    headerFooter: headerFooterSchema.optional(),
    objects: z.array(z.custom<ResolvedObject>()).readonly(),
    placeholders: z.array(resolvedPlaceholderSchema).readonly(),
    capabilities: z.record(z.string(), z.unknown()).readonly().optional(),
  })
  .readonly();
export type ResolvedSlideLayout = z.infer<typeof resolvedSlideLayoutSchema>;

/** Transitional aliases retained until Ticket 11. */
export const slideLayoutObservationSchema = resolvedSlideLayoutSchema;
export type SlideLayoutObservation = ResolvedSlideLayout;

export const resolvedMasterSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().optional(),
    preserve: z.boolean().optional(),
    colorMap: colorMapSchema.optional(),
    headerFooter: headerFooterSchema.optional(),
    objects: z.array(z.custom<ResolvedObject>()).readonly(),
    placeholders: z.array(resolvedPlaceholderSchema).readonly(),
    capabilities: z.record(z.string(), z.unknown()).readonly().optional(),
    layouts: z.array(resolvedSlideLayoutSchema).readonly(),
  })
  .readonly();
export type ResolvedMaster = z.infer<typeof resolvedMasterSchema>;

/** Transitional aliases retained until Ticket 11. */
export const slideMasterObservationSchema = resolvedMasterSchema;
export type SlideMasterObservation = ResolvedMaster;

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
    throw new CapabilityError(
      "invalid-slide-layout",
      first?.message ?? "Invalid slide layout specification",
    );
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
