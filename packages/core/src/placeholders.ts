import { z } from "zod";

import { paddingStyleSchema } from "#/layout-schemas.ts";
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

export const placeholderSzSchema = z.enum(["full", "half", "quarter"]);
export type PlaceholderSz = z.infer<typeof placeholderSzSchema>;

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
  sz: placeholderSzSchema.optional(),
  orient: placeholderOrientSchema.optional(),
  hasCustomPrompt: z.boolean().optional(),
  name: z.string().optional(),
  x: z.number().finite().optional(),
  y: z.number().finite().optional(),
  width: z.number().finite().positive().optional(),
  height: z.number().finite().positive().optional(),
  geometry: z.unknown().optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
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
    sz: placeholderSzSchema.optional(),
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
};

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
