import {
  authorIdSchema,
  colorValueSchema,
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
import { lineEndStyleSchema } from "../shared/visual.ts";

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
  ...lineEndStyleSchema.unwrap().shape,
  ...textStyleSchema.unwrap().shape,
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
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
  x: shapeCoordinateSchema.optional(),
  y: shapeCoordinateSchema.optional(),
  width: shapeSizeSchema.optional(),
  height: shapeSizeSchema.optional(),
  style: shapeStyleSchema.optional(),
  fill: colorValueSchema.optional(),
  /** Directly writes native text into this shape without creating an additional text box; mutually exclusive with children. */
  text: z.string().optional(),
  /** Shape text uses shared text services and can still be emitted when text box capability is omitted. */
  textStyle: textStyleSchema.optional(),
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
