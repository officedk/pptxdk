import { z } from "zod";

import { geometrySchema } from "#/geometry-schemas.ts";
import { resolvePercent } from "#/layout-protocol.ts";
import {
  layoutStyleSchema,
  paddingStyleSchema,
  type LayoutLength,
} from "#/layout-schemas.ts";
import {
  textStyleShape,
  visualStyleSchema,
  type ThemeContext,
} from "#/theme.ts";

const singleUnicodeCharacterSchema = z.string().refine((character) => {
  const codePoint = character.codePointAt(0);
  return (
    codePoint !== undefined && String.fromCodePoint(codePoint) === character
  );
}, "Bullet must be a single Unicode character");

export const textListBulletSchema = z
  .strictObject({
    type: z.literal("bullet"),
    character: singleUnicodeCharacterSchema.optional(),
    level: z.number().int().min(0).max(8).optional(),
  })
  .readonly();

export const textListNumberSchema = z
  .strictObject({
    type: z.literal("number"),
    style: z.literal("arabicPeriod").optional(),
    startAt: z.number().int().min(1).max(32767).optional(),
    level: z.number().int().min(0).max(8).optional(),
  })
  .readonly();

/** Native paragraph list schema (bullets and Arabic period numbering). */
export const textListSchema = z.discriminatedUnion("type", [
  textListBulletSchema,
  textListNumberSchema,
]);
export type TextList = z.infer<typeof textListSchema>;

export const textPlaceholderRoles = [
  "title",
  "ctrTitle",
  "subTitle",
  "body",
  "dt",
  "ftr",
  "hdr",
] as const;

export const textPlaceholderRoleSchema = z.enum(textPlaceholderRoles);
export type TextPlaceholderRole = z.infer<typeof textPlaceholderRoleSchema>;

export const contentPlaceholderRoles = [
  ...textPlaceholderRoles,
  "pic",
  "clipArt",
  "tbl",
  "chart",
  "obj",
  "sldNum",
  "dgm",
  "media",
] as const;

export const placeholderRoleSchema = z.enum(contentPlaceholderRoles);
export type ContentPlaceholderRole = z.infer<typeof placeholderRoleSchema>;

/** Official OOXML ST_PlaceholderSize values for placeholder slots. */
export const placeholderSizeSchema = z.enum(["full", "half", "quarter"]);
export type PlaceholderSize = z.infer<typeof placeholderSizeSchema>;

export const placeholderOrientSchema = z.enum(["horz", "vert"]);
export type PlaceholderOrient = z.infer<typeof placeholderOrientSchema>;

export const placeholderLevelStyleSchema = z
  .strictObject({
    ...textStyleShape,
    list: textListSchema.optional(),
    marginLeft: z.number().finite().nonnegative().optional(),
    indent: z.number().finite().optional(),
    spacingBefore: z.number().finite().nonnegative().optional(),
  })
  .readonly();
export type PlaceholderLevelStyle = z.infer<typeof placeholderLevelStyleSchema>;

export const placeholderListStylesSchema = z
  .strictObject({
    1: placeholderLevelStyleSchema.optional(),
    2: placeholderLevelStyleSchema.optional(),
    3: placeholderLevelStyleSchema.optional(),
    4: placeholderLevelStyleSchema.optional(),
    5: placeholderLevelStyleSchema.optional(),
    6: placeholderLevelStyleSchema.optional(),
    7: placeholderLevelStyleSchema.optional(),
    8: placeholderLevelStyleSchema.optional(),
    9: placeholderLevelStyleSchema.optional(),
  })
  .readonly();
export type PlaceholderListStyles = z.infer<typeof placeholderListStylesSchema>;

export const placeholderStyleSchema = z
  .strictObject({
    ...layoutStyleSchema.unwrap().shape,
    ...paddingStyleSchema.unwrap().shape,
    ...textStyleShape,
    ...visualStyleSchema.unwrap().shape,
  })
  .readonly();
export type PlaceholderStyle = z.infer<typeof placeholderStyleSchema>;

const basePlaceholderSpecSchema = z.strictObject({
  type: placeholderRoleSchema,
  explicitType: z.boolean().optional(),
  idx: z.number().int().nonnegative().optional(),
  size: placeholderSizeSchema.optional(),
  orient: placeholderOrientSchema.optional(),
  hasCustomPrompt: z.boolean().optional(),
  name: z.string().optional(),
  geometry: geometrySchema.optional(),
  style: placeholderStyleSchema.optional(),
  listStyles: placeholderListStylesSchema.optional(),
  promptText: z.string().optional(),
});

export const layoutPlaceholderSpecSchema = basePlaceholderSpecSchema.readonly();

export type LayoutPlaceholderSpec<
  Role extends ContentPlaceholderRole = ContentPlaceholderRole,
> = Omit<z.infer<typeof layoutPlaceholderSpecSchema>, "type"> & {
  readonly type: Role;
};

export const masterPlaceholderSpecSchema = layoutPlaceholderSpecSchema;
export type MasterPlaceholderSpec = LayoutPlaceholderSpec;

export const placeholderMapSchema = z
  .record(z.string().min(1), layoutPlaceholderSpecSchema)
  .readonly();

export type PlaceholderMap = Readonly<Record<string, LayoutPlaceholderSpec>>;
export type EmptyPlaceholderMap = Readonly<
  Record<never, LayoutPlaceholderSpec>
>;
export type MasterPlaceholderMap = Readonly<
  Record<string, MasterPlaceholderSpec>
>;

export const masterPlaceholderMapSchema = placeholderMapSchema;

export const masterTextStylesSchema = z
  .strictObject({
    titleStyle: placeholderListStylesSchema.optional(),
    bodyStyle: placeholderListStylesSchema.optional(),
    otherStyle: placeholderListStylesSchema.optional(),
  })
  .readonly();
export type MasterTextStyles = z.infer<typeof masterTextStylesSchema>;

export const nativePlaceholderRefSchema = z
  .strictObject({
    type: placeholderRoleSchema,
    explicitType: z.boolean().optional(),
    idx: z.number().int().nonnegative().optional(),
    size: placeholderSizeSchema.optional(),
    orient: placeholderOrientSchema.optional(),
    hasCustomPrompt: z.boolean().optional(),
  })
  .readonly();
export type NativePlaceholderRef = z.infer<typeof nativePlaceholderRefSchema>;

export type PlaceholderResolutionContext = {
  readonly key?: string;
  readonly nativeRef: NativePlaceholderRef;
  readonly masterTheme: ThemeContext;
  readonly masterTextStyles?: MasterTextStyles;
  readonly masterPlaceholder?: MasterPlaceholderSpec;
  readonly layoutPlaceholder?: LayoutPlaceholderSpec;
  readonly slideDimensions?: {
    readonly width: number;
    readonly height: number;
  };
};

export type PlaceholderBoxLayer = {
  readonly style?: Readonly<Record<string, unknown>> | undefined;
};

function resolvePlaceholderScalar(
  value: unknown,
  basis: number | undefined,
): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (basis !== undefined && typeof value === "string") {
    return resolvePercent(value as LayoutLength, basis);
  }
  return undefined;
}

/**
 * Cascades box coordinates (`left`/`right`/`inset`/`width` and `top`/`bottom`/`inset`/`height`)
 * across Master -> Layout -> Slide placeholder style layers, resolving percentages against `slideDimensions`.
 */
export function resolvePlaceholderBoxFromLayers(
  layers: readonly (PlaceholderBoxLayer | undefined)[],
  slideDimensions?: { readonly width: number; readonly height: number },
): {
  readonly x: unknown;
  readonly y: unknown;
  readonly width: unknown;
  readonly height: unknown;
} {
  let accLeft: unknown;
  let accRight: unknown;
  let accWidth: unknown;
  let accTop: unknown;
  let accBottom: unknown;
  let accHeight: unknown;

  for (const layer of layers) {
    if (layer === undefined) continue;
    const style =
      typeof layer.style === "object" &&
      layer.style !== null &&
      !Array.isArray(layer.style)
        ? layer.style
        : undefined;
    if (style === undefined) continue;
    const hasOwn = (obj: Readonly<Record<string, unknown>>, key: string) =>
      Object.hasOwn(obj, key) && obj[key] !== undefined;
    const hasLeft = hasOwn(style, "left") || hasOwn(style, "inset");
    const hasRight = hasOwn(style, "right") || hasOwn(style, "inset");
    const hasWidth = hasOwn(style, "width");
    const hasTop = hasOwn(style, "top") || hasOwn(style, "inset");
    const hasBottom = hasOwn(style, "bottom") || hasOwn(style, "inset");
    const hasHeight = hasOwn(style, "height");

    if (hasRight && !hasLeft) accLeft = undefined;
    if (hasLeft && !hasRight && hasWidth) accRight = undefined;
    if (hasLeft) {
      accLeft = hasOwn(style, "left") ? style["left"] : style["inset"];
    }
    if (hasRight) {
      accRight = hasOwn(style, "right") ? style["right"] : style["inset"];
    }
    if (hasWidth) {
      accWidth = style["width"];
    }

    if (hasBottom && !hasTop) accTop = undefined;
    if (hasTop && !hasBottom && hasHeight) accBottom = undefined;
    if (hasTop) {
      accTop = hasOwn(style, "top") ? style["top"] : style["inset"];
    }
    if (hasBottom) {
      accBottom = hasOwn(style, "bottom") ? style["bottom"] : style["inset"];
    }
    if (hasHeight) {
      accHeight = style["height"];
    }
  }

  const leftNum = resolvePlaceholderScalar(accLeft, slideDimensions?.width);
  const rightNum = resolvePlaceholderScalar(accRight, slideDimensions?.width);
  const widthNum = resolvePlaceholderScalar(accWidth, slideDimensions?.width);
  let resolvedX: unknown = leftNum ?? accLeft;
  let resolvedWidth: unknown = widthNum ?? accWidth;
  if (
    slideDimensions !== undefined &&
    leftNum !== undefined &&
    rightNum !== undefined &&
    widthNum === undefined
  ) {
    resolvedWidth = slideDimensions.width - leftNum - rightNum;
  } else if (
    slideDimensions !== undefined &&
    leftNum === undefined &&
    rightNum !== undefined &&
    widthNum !== undefined
  ) {
    resolvedX = slideDimensions.width - rightNum - widthNum;
  }

  const topNum = resolvePlaceholderScalar(accTop, slideDimensions?.height);
  const bottomNum = resolvePlaceholderScalar(
    accBottom,
    slideDimensions?.height,
  );
  const heightNum = resolvePlaceholderScalar(
    accHeight,
    slideDimensions?.height,
  );
  let resolvedY: unknown = topNum ?? accTop;
  let resolvedHeight: unknown = heightNum ?? accHeight;
  if (
    slideDimensions !== undefined &&
    topNum !== undefined &&
    bottomNum !== undefined &&
    heightNum === undefined
  ) {
    resolvedHeight = slideDimensions.height - topNum - bottomNum;
  } else if (
    slideDimensions !== undefined &&
    topNum === undefined &&
    bottomNum !== undefined &&
    heightNum !== undefined
  ) {
    resolvedY = slideDimensions.height - bottomNum - heightNum;
  }

  return {
    x: resolvedX,
    y: resolvedY,
    width: resolvedWidth,
    height: resolvedHeight,
  };
}

const baseAuthorKeyByRole: Readonly<Record<ContentPlaceholderRole, string>> = {
  title: "title",
  ctrTitle: "ctrTitle",
  subTitle: "subTitle",
  body: "body",
  obj: "content",
  pic: "pic",
  clipArt: "clipArt",
  tbl: "tbl",
  chart: "chart",
  dt: "dt",
  ftr: "ftr",
  sldNum: "sldNum",
  hdr: "hdr",
  dgm: "dgm",
  media: "media",
};

/**
 * Allocates deterministic unique author keys for placeholders on a SlideLayout or SlideMaster.
 *
 * Never infers semantic column/side names from coordinates. When multiple placeholders share the
 * same base role (e.g. two `body` slots with `idx=1` and `idx=3`), disambiguates deterministically
 * via `${base}-${idx}` (falling back to `-2`, `-3` in XML order if needed).
 */
export function allocatePlaceholderAuthorKeys(
  specs: readonly {
    readonly type: ContentPlaceholderRole;
    readonly explicitType?: boolean | undefined;
    readonly idx?: number | undefined;
  }[],
): readonly string[] {
  const counts = new Map<string, number>();
  const bases = specs.map((spec) =>
    spec.explicitType === false ? "content" : baseAuthorKeyByRole[spec.type],
  );
  for (const base of bases) {
    counts.set(base, (counts.get(base) ?? 0) + 1);
  }
  const used = new Set<string>();
  return specs.map((spec, index) => {
    const base = bases[index]!;
    const count = counts.get(base) ?? 1;
    const preferred =
      count === 1
        ? base
        : spec.idx !== undefined
          ? `${base}-${spec.idx}`
          : base;
    if (!used.has(preferred)) {
      used.add(preferred);
      return preferred;
    }
    let suffix = 2;
    while (used.has(`${preferred}-${suffix}`)) {
      suffix += 1;
    }
    const unique = `${preferred}-${suffix}`;
    used.add(unique);
    return unique;
  });
}

/**
 * Matches a slide `<p:ph>` reference onto a layout placeholder slot per ECMA-376 §19.3.1.36:
 * `@idx` matches first; when `@idx` is absent, matches `@type` (preferring an unindexed slot).
 */
export function matchLayoutPlaceholderKey(
  placeholders: PlaceholderMap | undefined,
  ref: {
    readonly type?: ContentPlaceholderRole | undefined;
    readonly idx?: number | undefined;
  },
): string | undefined {
  if (placeholders === undefined) return undefined;
  const entries = Object.entries(placeholders);
  if (ref.idx !== undefined) {
    const byIdx = entries.find(([, spec]) => spec.idx === ref.idx);
    if (byIdx !== undefined) return byIdx[0];
    return undefined;
  }
  if (ref.type !== undefined) {
    const unindexedByType = entries.find(
      ([, spec]) => spec.idx === undefined && spec.type === ref.type,
    );
    if (unindexedByType !== undefined) return unindexedByType[0];
    const anyByType = entries.find(([, spec]) => spec.type === ref.type);
    if (anyByType !== undefined) return anyByType[0];
  }
  return undefined;
}

/**
 * Matches a layout or slide placeholder onto its parent master placeholder by role (not `@idx`).
 */
export function matchMasterPlaceholderSpec(
  masterPlaceholders:
    | Readonly<Record<string, MasterPlaceholderSpec>>
    | undefined,
  role: ContentPlaceholderRole,
): MasterPlaceholderSpec | undefined {
  if (masterPlaceholders === undefined) return undefined;
  const specs = Object.values(masterPlaceholders);
  const exact = specs.find((spec) => spec.type === role);
  if (exact !== undefined) return exact;
  switch (role) {
    case "ctrTitle":
      return specs.find((spec) => spec.type === "title");
    case "subTitle":
    case "obj":
    case "pic":
    case "clipArt":
    case "tbl":
    case "chart":
    case "dgm":
    case "media":
      return specs.find((spec) => spec.type === "body");
    default:
      return undefined;
  }
}

/** Selects the master `<p:txStyles>` bucket (`titleStyle`, `bodyStyle`, `otherStyle`) for a role. */
export function masterTextStyleCategoryForRole(
  role: ContentPlaceholderRole,
): "titleStyle" | "bodyStyle" | "otherStyle" {
  switch (role) {
    case "title":
    case "ctrTitle":
      return "titleStyle";
    case "body":
    case "subTitle":
    case "obj":
    case "pic":
    case "clipArt":
    case "tbl":
    case "chart":
    case "dgm":
    case "media":
      return "bodyStyle";
    case "dt":
    case "ftr":
    case "sldNum":
    case "hdr":
      return "otherStyle";
    default: {
      const exhaustive: never = role;
      return exhaustive;
    }
  }
}
