import {
  authorIdSchema,
  layoutStyleSchema,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  visualStyleSchema,
} from "@pptxdk/core/capability";
import {
  ooxmlIntSchema,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
} from "@pptxdk/import";
import { z } from "zod";

import { motionPropsSchema } from "../motion/motion-schemas.ts";
import { geometrySchema } from "../shared/geometry.ts";
import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema } from "../shared/morph-api.ts";

export const imageFitSchema = z.enum(["fill", "cover", "contain"]);
export type ImageFit = z.infer<typeof imageFitSchema>;

export const cropFractionSchema = z.number().finite().min(0).max(1);

export const imageCropInputSchema = z
  .strictObject({
    left: cropFractionSchema.optional(),
    top: cropFractionSchema.optional(),
    right: cropFractionSchema.optional(),
    bottom: cropFractionSchema.optional(),
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
export type ImageCropInput = z.infer<typeof imageCropInputSchema>;

export const externalImageSourceSchema = z.strictObject({
  externalUrl: z.string().min(1),
});
export type ExternalImageSource = z.infer<typeof externalImageSourceSchema>;

export const imageSourceSchema = z.union([
  z.string(),
  z.instanceof(URL),
  z.custom<Uint8Array>((value) => value instanceof Uint8Array),
  externalImageSourceSchema,
]);
export type ImageSource = z.infer<typeof imageSourceSchema>;

export const imageCoordinateValueSchema = z
  .number()
  .finite()
  .min(-1_000_000)
  .max(1_000_000);

export const imageExtentValueSchema = z.number().finite().positive();

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

const motionPropsShape = motionPropsSchema.unwrap().shape;

const imageDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  src: imageSourceSchema,
  /** Alternative text for the image. */
  alt: z.string().optional(),
  /** Position and bounds in px; position defaults to 0, width and height default to original image dimensions. */
  x: imageCoordinateValueSchema.optional(),
  y: imageCoordinateValueSchema.optional(),
  width: imageExtentValueSchema.optional(),
  height: imageExtentValueSchema.optional(),
  geometry: geometrySchema.optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
  style: imageStyleSchema.optional(),
  /** `fill` stretches, `cover` crops centered, `contain` preserves aspect ratio; defaults to `fill`. */
  fit: imageFitSchema.optional(),
  /** Explicit crop fraction for each side of the original image; takes precedence over `fit` when provided. */
  crop: imageCropInputSchema.optional(),
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

export const nonVisualPicIdSchema = z.object({
  id: ooxmlNonNegativeIntSchema,
  name: z.string().optional(),
  descr: z.string().optional(),
});

const OoxmlCropFractionSchema = ooxmlNonNegativeIntSchema.refine(
  (value) => value <= 100_000,
  { message: "Crop value cannot exceed 100000" },
);

export const srcRectFragmentSchema = z
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

export const alphaModFixFragmentSchema = z.object({
  amt: OoxmlCropFractionSchema,
});

export const presetGeomFragmentSchema = z.object({
  prst: z.string().min(1),
});

export const geomGuideFragmentSchema = z.object({
  name: z.string().min(1),
  fmla: z.string().regex(/^val (-?\d+)$/u),
});

export const customPathAttributesSchema = z.object({
  w: ooxmlPositiveIntSchema.optional(),
  h: ooxmlPositiveIntSchema.optional(),
  fill: z
    .enum(["none", "norm", "lighten", "lightenLess", "darken", "darkenLess"])
    .optional(),
  stroke: z.enum(["0", "1", "true", "false"]).optional(),
});

export const customPathPointSchema = z.object({
  x: ooxmlIntSchema,
  y: ooxmlIntSchema,
});

export const lineAttributesFragmentSchema = z.object({
  w: ooxmlPositiveIntSchema.optional(),
});

export const presetDashFragmentSchema = z.object({
  val: z.enum(["solid", "dash", "sysDot", "dot"]),
});

export const outerShadowFragmentSchema = z.object({
  blurRad: ooxmlNonNegativeIntSchema.optional(),
  dist: ooxmlNonNegativeIntSchema.optional(),
  dir: ooxmlNonNegativeIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: z.string().optional(),
});

export const softEdgeFragmentSchema = z.object({
  rad: ooxmlNonNegativeIntSchema,
});
