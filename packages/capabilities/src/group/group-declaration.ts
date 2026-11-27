import {
  authorIdSchema,
  layoutStyleShape,
  linkTargetSchema,
  textStyleShape,
  visualStyleShape,
  type AuthorId,
  type LayoutStyle,
  type LinkTarget,
  type TextStyle,
  type VisualStyle,
} from "@pptxdk/core/capability";
import { CommonContainer } from "@pptxdk/core/document";
import { z } from "zod";

import {
  motionPropsShape,
  type MotionProps,
} from "../motion/motion-schemas.ts";
import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema, type MorphKey } from "../shared/morph-api.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type GroupStyle = LayoutStyle &
  TextStyle &
  Pick<VisualStyle, "rotate" | "flipH" | "flipV">;

const groupStyleObjectSchema = z.strictObject({
  ...layoutStyleShape,
  ...textStyleShape,
  rotate: visualStyleShape.rotate,
  flipH: visualStyleShape.flipH,
  flipV: visualStyleShape.flipV,
});

export const groupStyleKeys: readonly string[] = Object.keys(
  groupStyleObjectSchema.shape,
);
export const groupTransformShape = {
  rotate: visualStyleShape.rotate,
  flipH: visualStyleShape.flipH,
  flipV: visualStyleShape.flipV,
};
const rawGroupStyleSchema = groupStyleObjectSchema.readonly();
export const groupStyleSchema: z.ZodType<GroupStyle> =
  rawGroupStyleSchema satisfies z.ZodType<
    AssertExact<GroupStyle, z.infer<typeof rawGroupStyleSchema>> extends true
      ? GroupStyle
      : never
  >;

/** Native Group preserves member relationships; style allows it to arrange as a layout child, with output bounds computed from the union of member geometries. */
export type GroupDeclarationProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly name?: string | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  readonly style?: GroupStyle | undefined;
};

const groupDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  style: rawGroupStyleSchema.optional(),
  enter: motionPropsShape.enter,
  emphasis: motionPropsShape.emphasis,
  motion: motionPropsShape.motion,
});

export const groupDeclarationPropsKeys = Object.keys(
  groupDeclarationPropsObjectSchema.shape,
);

const rawGroupDeclarationPropsSchema =
  groupDeclarationPropsObjectSchema.readonly();
export const groupDeclarationPropsSchema: z.ZodType<GroupDeclarationProps> =
  rawGroupDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      GroupDeclarationProps,
      z.infer<typeof rawGroupDeclarationPropsSchema>
    > extends true
      ? GroupDeclarationProps
      : never
  >;

/** Creates the editable native group shared by the command factory and the Group JSX element. */
export function createGroupNode(
  props: GroupDeclarationProps,
): CommonContainer<
  "group",
  GroupDeclarationProps,
  NonNullable<GroupDeclarationProps["style"]>
> {
  return new CommonContainer<
    "group",
    GroupDeclarationProps,
    NonNullable<GroupDeclarationProps["style"]>
  >("group", "group", props);
}
