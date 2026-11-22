import {
  authorIdSchema,
  layoutStyleSchema,
  linkTargetSchema,
  textStyleSchema,
  themeScopeSchema,
  visualStyleSchema,
} from "@pptxdk/core/capability";
import { CommonContainer } from "@pptxdk/core/document";
import { z } from "zod";

import { MotionPropsSchema } from "../motion/motion-schemas.ts";
import { MorphKeySchema } from "../shared/morph-api.ts";

const boxStyleObjectSchema = z.strictObject({
  ...layoutStyleSchema.unwrap().shape,
  ...textStyleSchema.unwrap().shape,
  ...visualStyleSchema.unwrap().shape,
});

export const boxStyleSchema = boxStyleObjectSchema.readonly();
export type BoxStyle = z.infer<typeof boxStyleSchema>;

const motionPropsShape = MotionPropsSchema.unwrap().shape;

const boxDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  morphKey: MorphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  style: boxStyleSchema.optional(),
  theme: themeScopeSchema.optional(),
  enter: motionPropsShape.enter,
  motion: motionPropsShape.motion,
});

export const boxDeclarationPropsKeys = Object.keys(
  boxDeclarationPropsObjectSchema.shape,
);

export const boxDeclarationPropsSchema =
  boxDeclarationPropsObjectSchema.readonly();

/** Layout container appearance with visual styling that optionally generates a native background shape. */
export type BoxDeclarationProps = z.infer<typeof boxDeclarationPropsSchema>;

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
