import { z } from "zod";

import { CapabilityError } from "#/diagnostics.ts";
import {
  colorMapOverrideSchema,
  colorMapSchema,
  themeOverrideSchema,
  type ColorMapOverride,
  type TextStyle,
  type Theme,
  type ThemeInput,
  type ThemeOverride,
} from "#/theme.ts";

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
    colorMap: colorMapSchema.optional(),
  })
  .readonly();
export type SlideMasterDataProps = z.infer<typeof slideMasterDataPropsSchema>;

export type SlideMasterProps = SlideMasterDataProps & {
  readonly theme?: ThemeInput | Theme | undefined;
  readonly style?: TextStyle | undefined;
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
    colorMapOverride: colorMapOverrideSchema.optional(),
    themeOverride: themeOverrideSchema.optional(),
  })
  .readonly();
export type SlideLayoutDataProps = z.infer<typeof slideLayoutDataPropsSchema>;

export const slideLayoutObservationSchema = z
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
  })
  .readonly();
export type SlideLayoutObservation = z.infer<
  typeof slideLayoutObservationSchema
>;

export const slideMasterObservationSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().optional(),
    preserve: z.boolean().optional(),
    colorMap: colorMapSchema.optional(),
    layouts: z.array(slideLayoutObservationSchema).readonly(),
  })
  .readonly();
export type SlideMasterObservation = z.infer<
  typeof slideMasterObservationSchema
>;

const slideLayoutDefinitionSymbol = Symbol.for(
  "@pptxdk/core/slide-layout-definition",
);

/** Stateless, schema-validated slide layout specification shareable across JSX and imperative decks. */
export type SlideLayoutDefinition<
  Spec extends SlideLayoutDataProps = SlideLayoutDataProps,
> = {
  readonly [slideLayoutDefinitionSymbol]: true;
  readonly props: Spec;
  readonly id?: string | undefined;
  readonly name?: string | undefined;
  readonly type?: SlideLayoutType | undefined;
  readonly preserve?: boolean | undefined;
  readonly matchingName?: string | undefined;
  readonly userDrawn?: boolean | undefined;
  readonly showMasterShapes?: boolean | undefined;
  readonly showMasterPlaceholderAnimations?: boolean | undefined;
  readonly colorMapOverride?: ColorMapOverride | undefined;
  readonly themeOverride?: ThemeOverride | undefined;
};

export type SlideLayoutProps = SlideLayoutDataProps & {
  readonly layout?: SlideLayoutDefinition | undefined;
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
export function defineSlideLayout<const Spec extends SlideLayoutDataProps>(
  spec: Spec,
): SlideLayoutDefinition<Spec> {
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
  const frozenProps = Object.freeze(normalized) as Spec;
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
    ...(frozenProps.colorMapOverride === undefined
      ? {}
      : { colorMapOverride: frozenProps.colorMapOverride }),
    ...(frozenProps.themeOverride === undefined
      ? {}
      : { themeOverride: frozenProps.themeOverride }),
  });
}
