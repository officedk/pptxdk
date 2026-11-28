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
import { morphKeySchema, type MorphKey } from "../shared/morph-api.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type BoxStyle = LayoutStyle & TextStyle & VisualStyle;

const boxStyleObjectSchema = z.strictObject({
  ...layoutStyleShape,
  ...textStyleShape,
  ...visualStyleShape,
});
export const boxStyleKeys: readonly string[] = Object.keys(
  boxStyleObjectSchema.shape,
);
const rawBoxStyleSchema = boxStyleObjectSchema.readonly();
export const boxStyleSchema: z.ZodType<BoxStyle> =
  rawBoxStyleSchema satisfies z.ZodType<
    AssertExact<BoxStyle, z.infer<typeof rawBoxStyleSchema>> extends true
      ? BoxStyle
      : never
  >;

/** Layout container appearance with visual styling that optionally generates a native background shape. */
export type BoxDeclarationProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  readonly style?: BoxStyle | undefined;
};

const boxDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  style: rawBoxStyleSchema.optional(),
  ...motionPropsShape,
});

export const boxDeclarationPropsKeys = Object.keys(
  boxDeclarationPropsObjectSchema.shape,
);

const rawBoxDeclarationPropsSchema = boxDeclarationPropsObjectSchema.readonly();
export const boxDeclarationPropsSchema: z.ZodType<BoxDeclarationProps> =
  rawBoxDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      BoxDeclarationProps,
      z.infer<typeof rawBoxDeclarationPropsSchema>
    > extends true
      ? BoxDeclarationProps
      : never
  >;

/** Creates the editable layout container shared by the command factory and the Box JSX element. */
export function createBoxNode(
  props: BoxDeclarationProps,
): CommonContainer<
  "box",
  BoxDeclarationProps,
  NonNullable<BoxDeclarationProps["style"]>
> {
  return new CommonContainer<
    "box",
    BoxDeclarationProps,
    NonNullable<BoxDeclarationProps["style"]>
  >("box", "box", props);
}
