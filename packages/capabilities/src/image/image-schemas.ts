import {
  authorIdSchema,
  layoutStyleSchema,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  visualStyleSchema,
} from "@pptxdk/core/capability";
import {
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  OoxmlPositiveIntSchema,
} from "@pptxdk/import";
import { z } from "zod";

import { MotionPropsSchema } from "../motion/motion-schemas.ts";
import { geometrySchema } from "../shared/geometry.ts";
import { OptionalNameSchema } from "../shared/identity.ts";
import { MorphKeySchema } from "../shared/morph-api.ts";

export const ImageFitSchema = z.enum(["fill", "cover", "contain"]);
export type ImageFit = z.infer<typeof ImageFitSchema>;

export const CropFractionSchema = z.number().finite().min(0).max(1);

export const ImageCropInputSchema = z
  .strictObject({
    left: CropFractionSchema.optional(),
    top: CropFractionSchema.optional(),
    right: CropFractionSchema.optional(),
    bottom: CropFractionSchema.optional(),
  })
  .superRefine((crop, ctx) => {
    if ((crop.left ?? 0) + (crop.right ?? 0) > 1) {
      ctx.addIssue({
        code: "custom",
        message: "Sum of crop left and right cannot exceed 1",
        path: ["left"],
      });
    }
    if ((crop.top ?? 0) + (crop.bottom ?? 0) > 1) {
      ctx.addIssue({
        code: "custom",
        message: "Sum of crop top and bottom cannot exceed 1",
        path: ["top"],
      });
    }
  });
export type ImageCropInput = z.infer<typeof ImageCropInputSchema>;

export const ExternalImageSourceSchema = z.strictObject({
  externalUrl: z.string().min(1),
});
export type ExternalImageSource = z.infer<typeof ExternalImageSourceSchema>;

export const ImageSourceSchema = z.union([
  z.string(),
  z.instanceof(URL),
  z.custom<Uint8Array>((value) => value instanceof Uint8Array),
  ExternalImageSourceSchema,
]);
export type ImageSource = z.infer<typeof ImageSourceSchema>;

export const ImageCoordinateValueSchema = z
  .number()
  .finite()
  .min(-1_000_000)
  .max(1_000_000);

export const ImageExtentValueSchema = z.number().finite().positive();

const coreVisualShape = visualStyleSchema.unwrap().shape;

const imageVisualStyleObjectSchema = z.strictObject({
  borderColor: coreVisualShape.borderColor,
  borderWidth: coreVisualShape.borderWidth,
  borderStyle: coreVisualShape.borderStyle,
  opacity: coreVisualShape.opacity,
  boxShadow: coreVisualShape.boxShadow,
  softEdge: coreVisualShape.softEdge,
  rotate: coreVisualShape.rotate,
});

/** Visual styling supported on Image objects. */
export const imageVisualStyleSchema = imageVisualStyleObjectSchema.readonly();
export type ImageVisualStyle = z.infer<typeof imageVisualStyleSchema>;

export const imageVisualStyleKeys = new Set(
  Object.keys(imageVisualStyleObjectSchema.shape),
);

export const imageStyleSchema = z
  .strictObject({
    ...layoutStyleSchema.unwrap().shape,
    ...imageVisualStyleObjectSchema.shape,
  })
  .readonly();

const motionPropsShape = MotionPropsSchema.unwrap().shape;

const imageDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: OptionalNameSchema.optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: MorphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  src: ImageSourceSchema,
  /** Alternative text for the image. */
  alt: z.string().optional(),
  /** Position and bounds in px; position defaults to 0, width and height default to original image dimensions. */
  x: ImageCoordinateValueSchema.optional(),
  y: ImageCoordinateValueSchema.optional(),
  width: ImageExtentValueSchema.optional(),
  height: ImageExtentValueSchema.optional(),
  geometry: geometrySchema.optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
  style: imageStyleSchema.optional(),
  /** `fill` stretches, `cover` crops centered, `contain` preserves aspect ratio; defaults to `fill`. */
  fit: ImageFitSchema.optional(),
  /** Explicit crop fraction for each side of the original image; takes precedence over `fit` when provided. */
  crop: ImageCropInputSchema.optional(),
  enter: motionPropsShape.enter,
  emphasis: motionPropsShape.emphasis,
  motion: motionPropsShape.motion,
});

export const imageDeclarationPropsSchema =
  imageDeclarationPropsObjectSchema.readonly();

export type ImageDeclarationProps = z.infer<typeof imageDeclarationPropsSchema>;

export const imageAllowedParameters: readonly string[] = Object.keys(
  imageDeclarationPropsObjectSchema.shape,
);

export const NonVisualPicIdSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
  name: z.string().optional(),
  descr: z.string().optional(),
});

const OoxmlCropFractionSchema = OoxmlNonNegativeIntSchema.refine(
  (value) => value <= 100_000,
  { message: "Crop value cannot exceed 100000" },
);

export const SrcRectFragmentSchema = z
  .object({
    l: OoxmlCropFractionSchema.optional(),
    t: OoxmlCropFractionSchema.optional(),
    r: OoxmlCropFractionSchema.optional(),
    b: OoxmlCropFractionSchema.optional(),
  })
  .refine(
    (value) =>
      (value.l ?? 0) + (value.r ?? 0) <= 100_000 &&
      (value.t ?? 0) + (value.b ?? 0) <= 100_000,
    { message: "Crop sum cannot exceed 100000" },
  );

export const AlphaModFixFragmentSchema = z.object({
  amt: OoxmlCropFractionSchema,
});

export const PresetGeomFragmentSchema = z.object({
  prst: z.string().min(1),
});

export const GeomGuideFragmentSchema = z.object({
  name: z.string().min(1),
  fmla: z.string().regex(/^val (-?\d+)$/u),
});

export const CustomPathAttributesSchema = z.object({
  w: OoxmlPositiveIntSchema.optional(),
  h: OoxmlPositiveIntSchema.optional(),
  fill: z
    .enum(["none", "norm", "lighten", "lightenLess", "darken", "darkenLess"])
    .optional(),
  stroke: z.enum(["0", "1", "true", "false"]).optional(),
});

export const CustomPathPointSchema = z.object({
  x: OoxmlIntSchema,
  y: OoxmlIntSchema,
});

export const LineAttributesFragmentSchema = z.object({
  w: OoxmlPositiveIntSchema.optional(),
});

export const PresetDashFragmentSchema = z.object({
  val: z.enum(["solid", "dash", "sysDot", "dot"]),
});

export const OuterShadowFragmentSchema = z.object({
  blurRad: OoxmlNonNegativeIntSchema.optional(),
  dist: OoxmlNonNegativeIntSchema.optional(),
  dir: OoxmlNonNegativeIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: z.string().optional(),
});

export const SoftEdgeFragmentSchema = z.object({
  rad: OoxmlNonNegativeIntSchema,
});
