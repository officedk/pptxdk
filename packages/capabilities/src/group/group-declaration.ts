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

import { motionPropsSchema } from "../motion/motion-schemas.ts";
import {
  shapeCoordinateSchema,
  shapeSizeSchema,
} from "../shape/shape-declaration.ts";
import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema } from "../shared/morph-api.ts";

const groupStyleObjectSchema = z.strictObject({
  ...layoutStyleSchema.unwrap().shape,
  ...textStyleSchema.unwrap().shape,
  rotate: visualStyleSchema.unwrap().shape.rotate,
});

export const groupStyleSchema = groupStyleObjectSchema.readonly();
export type GroupStyle = z.infer<typeof groupStyleSchema>;

const motionPropsShape = motionPropsSchema.unwrap().shape;

const groupBoundsSchema = z
  .strictObject({
    x: shapeCoordinateSchema,
    y: shapeCoordinateSchema,
    width: shapeSizeSchema,
    height: shapeSizeSchema,
  })
  .readonly();

const groupDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  bounds: groupBoundsSchema.optional(),
  style: groupStyleSchema.optional(),
  theme: themeScopeSchema.optional(),
  enter: motionPropsShape.enter,
  emphasis: motionPropsShape.emphasis,
  motion: motionPropsShape.motion,
});

export const groupDeclarationPropsKeys = Object.keys(
  groupDeclarationPropsObjectSchema.shape,
);

export const groupDeclarationPropsSchema =
  groupDeclarationPropsObjectSchema.readonly();

/** Native Group preserves member relationships; style allows it to arrange as a layout child, with output bounds computed from the union of member geometries. */
export type GroupDeclarationProps = z.infer<typeof groupDeclarationPropsSchema>;

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
