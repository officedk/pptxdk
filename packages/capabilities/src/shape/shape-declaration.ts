import {
  authorIdSchema,
  layoutStyleSchema,
  linkTargetSchema,
  textStyleSchema,
  visualStyleSchema,
} from "@pptxdk/core/capability";
import { RichTextContainer } from "@pptxdk/core/document";
import { pxToEmu } from "@pptxdk/emit";
import { z } from "zod";

import { motionPropsSchema } from "../motion/motion-schemas.ts";
import { geometrySchema } from "../shared/geometry.ts";
import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema } from "../shared/morph-api.ts";
import { shapeBorderValueSchema } from "../shared/visual.ts";

export const shapeCoordinateSchema = z
  .number()
  .finite()
  .min(-1_000_000)
  .max(1_000_000);

export const shapeSizeSchema = z
  .number()
  .finite()
  .positive()
  .refine((value) => {
    const extentEmu = pxToEmu(value);
    return extentEmu >= 1 && extentEmu <= 2_147_483_647;
  });

const shapeStyleObjectSchema = z.strictObject({
  ...layoutStyleSchema.unwrap().shape,
  ...visualStyleSchema.unwrap().shape,
  ...textStyleSchema.unwrap().shape,
  border: shapeBorderValueSchema.optional(),
});

export const shapeStyleSchema = shapeStyleObjectSchema.readonly();
export type ShapeStyle = z.infer<typeof shapeStyleSchema>;

const motionPropsShape = motionPropsSchema.unwrap().shape;

const shapeDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  geometry: geometrySchema.optional(),
  style: shapeStyleSchema.optional(),
  enter: motionPropsShape.enter,
  emphasis: motionPropsShape.emphasis,
  motion: motionPropsShape.motion,
});

export const shapeDeclarationPropsKeys = [
  ...Object.keys(shapeDeclarationPropsObjectSchema.shape),
  "children",
] as const;

export const shapeDeclarationPropsSchema =
  shapeDeclarationPropsObjectSchema.readonly();

export type ShapeDeclarationProps = z.infer<typeof shapeDeclarationPropsSchema>;

/** Creates the editable native shape shared by the command factory and the Shape JSX element. */
export function createShapeNode(
  props: ShapeDeclarationProps,
): RichTextContainer<
  "shape",
  ShapeDeclarationProps,
  NonNullable<ShapeDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "shape",
    ShapeDeclarationProps,
    NonNullable<ShapeDeclarationProps["style"]>
  >("shape", "shape", props);
}
