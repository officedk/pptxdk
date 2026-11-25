import {
  isMorphKey,
  linkTargetSchema,
  type MorphKey,
} from "@pptxdk/core/capability";
import { z } from "zod";

/** Native OOXML object name prefix denoting explicit cross-slide Morph pairing (`ECMA-376` / PowerPoint convention). */
export const nativeMorphPrefix = "!!";

/** Runtime-identity-preserving Zod schema for handles created by `createMorphScope().handle(name)`. */
export const morphKeySchema = z.custom<MorphKey>(
  isMorphKey,
  "morphKey must be created by createMorphScope().handle(name)",
);

/**
 * Explicit cross-slide continuity declaration available to native objects.
 *
 * Morph writes handles into the OOXML native object name, so the same object cannot declare a separate name.
 */
export const morphPropsSchema = z
  .strictObject({
    morphKey: morphKeySchema.optional(),
  })
  .readonly();
export type MorphProps = z.infer<typeof morphPropsSchema>;

/** Click target schema for public native objects; core resolves identifiers into package relationships. */
export const linkPropsSchema = z
  .strictObject({
    link: linkTargetSchema.optional(),
  })
  .readonly();
export type LinkProps = z.infer<typeof linkPropsSchema>;

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
