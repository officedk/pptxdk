import type { LinkTarget, MorphKey } from "@react-pptx/core";

/**
 * Explicit cross-slide continuity declaration available to native objects.
 *
 * Morph writes handles into the OOXML native object name, so the same object cannot declare a separate name.
 */
export type MorphProps = {
  readonly morphKey?: MorphKey;
};

/** Click target for public native objects; core is responsible for resolving identifiers into package relationships. */
export type LinkProps = {
  readonly link?: LinkTarget;
};

export { createMorphScope } from "@react-pptx/core";
export type { LinkTarget, MorphKey, MorphScope } from "@react-pptx/core";
