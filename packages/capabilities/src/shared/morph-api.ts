import {
  isMorphKey,
  linkTargetSchema,
  type LinkTarget,
  type MorphKey,
} from "@pptxdk/core/capability";
import { z } from "zod";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

/** Native OOXML object name prefix denoting explicit cross-slide Morph pairing (`ECMA-376` / PowerPoint convention). */
export const nativeMorphPrefix = "!!";

/** Runtime-identity-preserving Zod schema for handles created by `createMorphScope().handle(name)`. */
export const morphKeySchema: z.ZodType<MorphKey> = z.custom<MorphKey>(
  isMorphKey,
  "morphKey must be created by createMorphScope().handle(name)",
);

export type MorphProps = {
  readonly morphKey?: MorphKey | undefined;
};

const rawMorphPropsSchema = z
  .strictObject({
    morphKey: morphKeySchema.optional(),
  })
  .readonly();
/**
 * Explicit cross-slide continuity declaration available to native objects.
 *
 * Morph writes handles into the OOXML native object name, so the same object cannot declare a separate name.
 */
export const morphPropsSchema: z.ZodType<MorphProps> =
  rawMorphPropsSchema satisfies z.ZodType<
    AssertExact<MorphProps, z.infer<typeof rawMorphPropsSchema>> extends true
      ? MorphProps
      : never
  >;

export type LinkProps = {
  readonly link?: LinkTarget | undefined;
};

const rawLinkPropsSchema = z
  .strictObject({
    link: linkTargetSchema.optional(),
  })
  .readonly();
/** Click target schema for public native objects; core resolves identifiers into package relationships. */
export const linkPropsSchema: z.ZodType<LinkProps> =
  rawLinkPropsSchema satisfies z.ZodType<
    AssertExact<LinkProps, z.infer<typeof rawLinkPropsSchema>> extends true
      ? LinkProps
      : never
  >;

/**
 * Extracts a cross-slide Morph token exclusively from native `!!...` `<p:cNvPr name="!!...">` markers.
 * Plain display names never synthesize Morph continuity.
 */
export function parseNativeMorphToken(
  objectName: string | undefined,
): string | undefined {
  if (objectName === undefined || !objectName.startsWith(nativeMorphPrefix)) {
    return undefined;
  }
  const token = objectName.slice(nativeMorphPrefix.length);
  return token.trim() === "" ? undefined : token;
}

export { createMorphScope } from "@pptxdk/core";
export type { LinkTarget, MorphKey, MorphScope } from "@pptxdk/core";
