import {
  colorValueSchema,
  isMorphKey,
  linkTargetSchema,
  textStyleSchema,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type MorphKey,
  type PaddingStyle,
} from "@pptxdk/core/capability";
import { RichTextContainer } from "@pptxdk/core/document";
import { pxToEmu } from "@pptxdk/emit";
import { z } from "zod";

import type { MotionProps } from "../motion/motion-api.ts";
import { geometrySchema, type Geometry } from "../shared/geometry.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import type { TextStyle } from "../shared/text-service.ts";
import type { ShapeVisualStyle } from "../shared/visual.ts";

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

const shapeDeclarationPropsObjectSchema = z.strictObject({
  id: z.string().optional(),
  name: z.string().optional(),
  morphKey: z.custom<MorphKey>(isMorphKey).optional(),
  link: linkTargetSchema.optional(),
  geometry: geometrySchema.optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
  x: shapeCoordinateSchema.optional(),
  y: shapeCoordinateSchema.optional(),
  width: shapeSizeSchema.optional(),
  height: shapeSizeSchema.optional(),
  style: z.record(z.string(), z.unknown()).readonly().optional(),
  fill: colorValueSchema.optional(),
  text: z.string().optional(),
  textStyle: textStyleSchema.optional(),
  enter: z.unknown().optional(),
  emphasis: z.unknown().optional(),
  motion: z.unknown().optional(),
});

export const shapeDeclarationPropsKeys = [
  ...Object.keys(shapeDeclarationPropsObjectSchema.shape),
  "children",
] as const;

export const shapeDeclarationPropsSchema =
  shapeDeclarationPropsObjectSchema.readonly();

export type ShapeDeclarationProps = MotionProps &
  MorphProps & {
    readonly link?: LinkTarget;
    readonly id?: string;
    readonly name?: string;
    readonly geometry?: Geometry;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly style?: LayoutStyle & ShapeVisualStyle & TextStyle & PaddingStyle;
    readonly fill?: ColorValue;
    /** Directly writes native text into this shape without creating an additional text box; mutually exclusive with children. */
    readonly text?: string;
    /** Shape text uses shared text services and can still be emitted when text box capability is omitted. */
    readonly textStyle?: TextStyle;
  };

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
