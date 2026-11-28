import {
  authorIdSchema,
  layoutStyleShape,
  linkTargetSchema,
  visualStyleShape,
  type AuthorId,
  type LayoutStyle,
  type LinkTarget,
  type PaddingStyle,
  type VisualStyle,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import { z } from "zod";

import {
  motionPropsShape,
  type MotionProps,
} from "../motion/motion-schemas.ts";
import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema, type MorphKey } from "../shared/morph-api.ts";
import {
  lineStrokeSchema,
  lineStrokeValueSchema,
  type ArrowType,
  type LineStroke,
  type LineStrokeValue,
} from "../shared/visual.ts";

export {
  lineStrokeSchema,
  lineStrokeValueSchema,
  type ArrowType,
  type LineStroke,
  type LineStrokeValue,
};

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type LineRouting = "straight" | "elbow" | "curved";
const rawLineRoutingSchema = z.enum(["straight", "elbow", "curved"]);
export const lineRoutingSchema: z.ZodType<LineRouting> =
  rawLineRoutingSchema satisfies z.ZodType<
    AssertExact<LineRouting, z.infer<typeof rawLineRoutingSchema>> extends true
      ? LineRouting
      : never
  >;

export type LinePort = "top" | "right" | "bottom" | "left" | number;
const rawLinePortSchema = z.union([
  z.enum(["top", "right", "bottom", "left"]),
  z.number().int().min(0),
]);
export const linePortSchema: z.ZodType<LinePort> =
  rawLinePortSchema satisfies z.ZodType<
    AssertExact<LinePort, z.infer<typeof rawLinePortSchema>> extends true
      ? LinePort
      : never
  >;

const lineCoordinateSchema = z.number().finite().min(-1_000_000).max(1_000_000);

export type LineEndpoint =
  | { readonly id: AuthorId; readonly port: LinePort }
  | { readonly x: number; readonly y: number };

const rawLineEndpointSchema = z.union([
  z
    .strictObject({
      id: authorIdSchema,
      port: rawLinePortSchema,
    })
    .readonly(),
  z
    .strictObject({
      x: lineCoordinateSchema,
      y: lineCoordinateSchema,
    })
    .readonly(),
]);
export const lineEndpointSchema: z.ZodType<LineEndpoint> =
  rawLineEndpointSchema satisfies z.ZodType<
    AssertExact<
      LineEndpoint,
      z.infer<typeof rawLineEndpointSchema>
    > extends true
      ? LineEndpoint
      : never
  >;

export type LineAdjust = {
  readonly adj1?: number | undefined;
  readonly adj2?: number | undefined;
  readonly adj3?: number | undefined;
};

const lineAdjustValueSchema = z
  .number()
  .refine(
    Number.isSafeInteger,
    "Geometry adjust must be an integer DrawingML parameter",
  );
const rawLineAdjustSchema = z
  .strictObject({
    adj1: lineAdjustValueSchema.optional(),
    adj2: lineAdjustValueSchema.optional(),
    adj3: lineAdjustValueSchema.optional(),
  })
  .readonly();
export const lineAdjustSchema: z.ZodType<LineAdjust> =
  rawLineAdjustSchema satisfies z.ZodType<
    AssertExact<LineAdjust, z.infer<typeof rawLineAdjustSchema>> extends true
      ? LineAdjust
      : never
  >;

export type LineVisualStyle = Pick<
  VisualStyle,
  "rotate" | "flipH" | "flipV" | "opacity" | "boxShadow"
>;
export type LineLayoutStyle = Omit<LayoutStyle, keyof PaddingStyle>;
export type LineStyle = LineLayoutStyle & LineVisualStyle;

const {
  padding: _padding,
  paddingTop: _paddingTop,
  paddingRight: _paddingRight,
  paddingBottom: _paddingBottom,
  paddingLeft: _paddingLeft,
  ...lineLayoutStyleShape
} = layoutStyleShape;

const lineStyleObjectSchema = z.strictObject({
  ...lineLayoutStyleShape,
  rotate: visualStyleShape.rotate,
  flipH: visualStyleShape.flipH,
  flipV: visualStyleShape.flipV,
  opacity: visualStyleShape.opacity,
  boxShadow: visualStyleShape.boxShadow,
});
export const lineStyleKeyList: readonly string[] = Object.keys(
  lineStyleObjectSchema.shape,
);
const rawLineStyleSchema = lineStyleObjectSchema.readonly();
export const lineStyleSchema: z.ZodType<LineStyle> =
  rawLineStyleSchema satisfies z.ZodType<
    AssertExact<LineStyle, z.infer<typeof rawLineStyleSchema>> extends true
      ? LineStyle
      : never
  >;

export type LineDeclarationProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly name?: string | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  readonly routing?: LineRouting | undefined;
  readonly adjust?: LineAdjust | undefined;
  readonly stroke?: LineStrokeValue | undefined;
  readonly from?: LineEndpoint | undefined;
  readonly to?: LineEndpoint | undefined;
  readonly style?: LineStyle | undefined;
};

const lineDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  routing: rawLineRoutingSchema.optional(),
  adjust: rawLineAdjustSchema.optional(),
  stroke: lineStrokeValueSchema.optional(),
  from: rawLineEndpointSchema.optional(),
  to: rawLineEndpointSchema.optional(),
  style: rawLineStyleSchema.optional(),
  ...motionPropsShape,
});

export const lineAllowedPropKeyList: readonly string[] = Object.keys(
  lineDeclarationPropsObjectSchema.shape,
);

const rawLineDeclarationPropsSchema =
  lineDeclarationPropsObjectSchema.readonly();
export const lineDeclarationPropsSchema: z.ZodType<LineDeclarationProps> =
  rawLineDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      LineDeclarationProps,
      z.infer<typeof rawLineDeclarationPropsSchema>
    > extends true
      ? LineDeclarationProps
      : never
  >;

/** Creates the editable native line/connector shared by the command factory and the Line JSX element. */
export function createLineNode(
  props: LineDeclarationProps,
): CommonNode<
  "line",
  LineDeclarationProps,
  NonNullable<LineDeclarationProps["style"]>
> {
  return new CommonNode<
    "line",
    LineDeclarationProps,
    NonNullable<LineDeclarationProps["style"]>
  >("line", "line", props);
}
