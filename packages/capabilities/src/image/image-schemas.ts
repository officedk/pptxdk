import {
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  OoxmlPositiveIntSchema,
} from "@pptxdk/import";
import { z } from "zod";

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
