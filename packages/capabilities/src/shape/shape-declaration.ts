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
import { RichTextContainer } from "@pptxdk/core/document";
import { pxToEmu } from "@pptxdk/emit";
import { z } from "zod";

import {
  motionPropsShape,
  type MotionProps,
} from "../motion/motion-schemas.ts";
import { geometrySchema, type Geometry } from "../shared/geometry.ts";
import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema, type MorphKey } from "../shared/morph-api.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

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

export type ShapeStyle = LayoutStyle & VisualStyle & TextStyle;

const shapeStyleObjectSchema = z.strictObject({
  ...layoutStyleShape,
  ...visualStyleShape,
  ...textStyleShape,
});
export const shapeStyleKeyList: readonly string[] = Object.keys(
  shapeStyleObjectSchema.shape,
);
const rawShapeStyleSchema = shapeStyleObjectSchema.readonly();
export const shapeStyleSchema: z.ZodType<ShapeStyle> =
  rawShapeStyleSchema satisfies z.ZodType<
    AssertExact<ShapeStyle, z.infer<typeof rawShapeStyleSchema>> extends true
      ? ShapeStyle
      : never
  >;

export type ShapeDeclarationProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly name?: string | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  readonly geometry?: Geometry | undefined;
  readonly style?: ShapeStyle | undefined;
};

const shapeDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  geometry: geometrySchema.optional(),
  style: rawShapeStyleSchema.optional(),
  enter: motionPropsShape.enter,
  emphasis: motionPropsShape.emphasis,
  motion: motionPropsShape.motion,
});

export const shapeAllowedPropKeyList: readonly string[] = Object.keys(
  shapeDeclarationPropsObjectSchema.shape,
);

export const shapeDeclarationPropsKeys = [
  ...shapeAllowedPropKeyList,
  "children",
] as const;

const rawShapeDeclarationPropsSchema =
  shapeDeclarationPropsObjectSchema.readonly();
export const shapeDeclarationPropsSchema: z.ZodType<ShapeDeclarationProps> =
  rawShapeDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      ShapeDeclarationProps,
      z.infer<typeof rawShapeDeclarationPropsSchema>
    > extends true
      ? ShapeDeclarationProps
      : never
  >;

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
