import {
  authorIdSchema,
  layoutStyleShape,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  visualStyleShape,
  type AuthorId,
  type LayoutStyle,
  type LinkTarget,
  type NativePlaceholderRef,
  type VisualStyle,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import { z } from "zod";

import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema, type MorphKey } from "../shared/morph-api.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type PlaceholderObjectStyle = LayoutStyle & VisualStyle;

const placeholderObjectStyleObjectSchema = z.strictObject({
  ...layoutStyleShape,
  ...visualStyleShape,
});
export const placeholderStyleKeys: readonly string[] = Object.keys(
  placeholderObjectStyleObjectSchema.shape,
);
export const placeholderObjectStyleSchema: z.ZodType<PlaceholderObjectStyle> =
  placeholderObjectStyleObjectSchema.readonly();

export type PlaceholderDeclarationProps = {
  readonly id?: AuthorId | undefined;
  readonly name?: string | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly placeholder?: string | undefined;
  readonly nativePlaceholder?: NativePlaceholderRef | undefined;
  readonly link?: LinkTarget | undefined;
  readonly style?: PlaceholderObjectStyle | undefined;
};

const placeholderDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: morphKeySchema.optional(),
  placeholder: z.string().min(1).optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  link: linkTargetSchema.optional(),
  style: placeholderObjectStyleSchema.optional(),
});

const rawPlaceholderDeclarationPropsSchema =
  placeholderDeclarationPropsObjectSchema.readonly();
export const placeholderDeclarationPropsSchema: z.ZodType<PlaceholderDeclarationProps> =
  rawPlaceholderDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      PlaceholderDeclarationProps,
      z.infer<typeof rawPlaceholderDeclarationPropsSchema>
    > extends true
      ? PlaceholderDeclarationProps
      : never
  >;

export const placeholderAllowedPropKeys = new Set(
  Object.keys(placeholderDeclarationPropsObjectSchema.shape),
);

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
