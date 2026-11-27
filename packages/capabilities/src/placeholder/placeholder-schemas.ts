import {
  authorIdSchema,
  layoutStyleSchema,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  visualStyleSchema,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import { z } from "zod";

import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema } from "../shared/morph-api.ts";

export const placeholderObjectStyleSchema = z
  .strictObject({
    ...layoutStyleSchema.unwrap().shape,
    ...visualStyleSchema.unwrap().shape,
  })
  .readonly();

const placeholderDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: morphKeySchema.optional(),
  placeholder: z.string().min(1).optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  link: linkTargetSchema.optional(),
  style: placeholderObjectStyleSchema.optional(),
});

export const placeholderDeclarationPropsSchema =
  placeholderDeclarationPropsObjectSchema.readonly();

export const placeholderAllowedPropKeys = new Set(
  Object.keys(placeholderDeclarationPropsObjectSchema.shape),
);

export type PlaceholderDeclarationProps = z.infer<
  typeof placeholderDeclarationPropsSchema
>;

export type PlaceholderNode = CommonNode<
  "placeholder",
  PlaceholderDeclarationProps
>;

/** Creates an editable AST node representing an unpopulated native content placeholder (`<p:sp><p:ph .../></p:sp>` without `<p:txBody>`). */
export function createPlaceholderNode(
  props: PlaceholderDeclarationProps,
): PlaceholderNode {
  return new CommonNode("placeholder", "placeholder", props);
}
