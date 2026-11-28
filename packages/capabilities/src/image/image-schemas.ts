import {
  authorIdSchema,
  layoutStyleShape,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  visualStyleShape,
  type AuthorId,
  type BackgroundValue,
  type BorderValue,
  type BoxShadowValue,
  type LayoutStyle,
  type LinkTarget,
  type NativePlaceholderRef,
} from "@pptxdk/core/capability";
import {
  ooxmlIntSchema,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
  ooxmlPositiveLineWidthEmuSchema,
  ooxmlWritableExtentEmuSchema,
} from "@pptxdk/import";
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

export type ImageFit = "fill" | "cover" | "contain";
const rawImageFitSchema = z.enum(["fill", "cover", "contain"]);
export const imageFitSchema: z.ZodType<ImageFit> =
  rawImageFitSchema satisfies z.ZodType<
    AssertExact<ImageFit, z.infer<typeof rawImageFitSchema>> extends true
      ? ImageFit
      : never
  >;

export const cropFractionSchema = z.number().finite().min(0).max(1);

export type ImageCropInput = {
  readonly left?: number | undefined;
  readonly top?: number | undefined;
  readonly right?: number | undefined;
  readonly bottom?: number | undefined;
};

const rawImageCropInputSchema = z
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
export const imageCropInputSchema: z.ZodType<ImageCropInput> =
  rawImageCropInputSchema satisfies z.ZodType<
    AssertExact<
      ImageCropInput,
      z.infer<typeof rawImageCropInputSchema>
    > extends true
      ? ImageCropInput
      : never
  >;

export type ExternalImageSource = {
  readonly externalUrl: string;
};

const rawExternalImageSourceSchema = z.strictObject({
  externalUrl: z.string().min(1),
});
export const externalImageSourceSchema: z.ZodType<ExternalImageSource> =
  rawExternalImageSourceSchema satisfies z.ZodType<
    AssertExact<
      ExternalImageSource,
      z.infer<typeof rawExternalImageSourceSchema>
    > extends true
      ? ExternalImageSource
      : never
  >;

/**
 * Image source accepted by `Image` / `image()`.
 *
 * Supports raster bitmap formats (**PNG, JPEG, GIF**) passed as a local file path, `file:` `URL`,
 * or `Uint8Array` bytes, or `{ externalUrl }` for linked external images. `.svg` files are not
 * supported; use `Shape` with `path()` custom geometry (`geometry={path(...)}`) for native
 * editable vector graphics.
 */
export type ImageSource = string | URL | Uint8Array | ExternalImageSource;

const rawImageSourceSchema = z.union([
  z.string(),
  z.instanceof(URL),
  z.custom<Uint8Array>((value) => value instanceof Uint8Array),
  rawExternalImageSourceSchema,
]);
export const imageSourceSchema: z.ZodType<ImageSource> =
  rawImageSourceSchema satisfies z.ZodType<
    AssertExact<ImageSource, z.infer<typeof rawImageSourceSchema>> extends true
      ? ImageSource
      : never
  >;

export const imageCoordinateValueSchema = z
  .number()
  .finite()
  .min(-1_000_000)
  .max(1_000_000);

export const imageExtentValueSchema = z.number().finite().positive();

/** Visual styling supported on Image objects. */
export type ImageVisualStyle = {
  readonly background?: BackgroundValue | undefined;
  readonly border?: BorderValue | undefined;
  readonly opacity?: number | undefined;
  readonly boxShadow?: BoxShadowValue | undefined;
  readonly softEdge?: number | undefined;
  readonly rotate?: number | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
};

const imageVisualStyleObjectSchema = z.strictObject({
  background: visualStyleShape.background,
  border: visualStyleShape.border,
  opacity: visualStyleShape.opacity,
  boxShadow: visualStyleShape.boxShadow,
  softEdge: visualStyleShape.softEdge,
  rotate: visualStyleShape.rotate,
  flipH: visualStyleShape.flipH,
  flipV: visualStyleShape.flipV,
});

const rawImageVisualStyleSchema = imageVisualStyleObjectSchema.readonly();
/** Visual styling supported on Image objects. */
export const imageVisualStyleSchema: z.ZodType<ImageVisualStyle> =
  rawImageVisualStyleSchema satisfies z.ZodType<
    AssertExact<
      ImageVisualStyle,
      z.infer<typeof rawImageVisualStyleSchema>
    > extends true
      ? ImageVisualStyle
      : never
  >;

export const imageVisualStyleKeys = new Set(
  Object.keys(imageVisualStyleObjectSchema.shape),
);

export type ImageStyle = LayoutStyle & ImageVisualStyle;

const imageStyleObjectSchema = z.strictObject({
  ...layoutStyleShape,
  ...imageVisualStyleObjectSchema.shape,
});
export const imageStyleKeys: readonly string[] = Object.keys(
  imageStyleObjectSchema.shape,
);
const rawImageStyleSchema = imageStyleObjectSchema.readonly();
export const imageStyleSchema: z.ZodType<ImageStyle> =
  rawImageStyleSchema satisfies z.ZodType<
    AssertExact<ImageStyle, z.infer<typeof rawImageStyleSchema>> extends true
      ? ImageStyle
      : never
  >;

export type ImageDeclarationProps = MotionProps & {
  readonly id?: AuthorId | undefined;
  readonly name?: string | undefined;
  readonly placeholder?: string | undefined;
  readonly nativePlaceholder?: NativePlaceholderRef | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  /**
   * Image source (PNG, JPEG, or GIF local path, `file:` `URL`, or `Uint8Array` bytes; or `{ externalUrl }`).
   *
   * `.svg` files are not supported; for native editable vector graphics, use `Shape` with `path()`
   * custom geometry (`geometry={path(...)}`).
   */
  readonly src: ImageSource;
  /** Alternative text for the image. */
  readonly alt?: string | undefined;
  readonly geometry?: Geometry | undefined;
  readonly style?: ImageStyle | undefined;
  /** `fill` stretches, `cover` crops centered, `contain` preserves aspect ratio; defaults to `fill`. */
  readonly fit?: ImageFit | undefined;
  /** Explicit crop fraction for each side of the original image; takes precedence over `fit` when provided. */
  readonly crop?: ImageCropInput | undefined;
};

const imageDeclarationPropsObjectSchema = z.strictObject({
  id: authorIdSchema.optional(),
  name: optionalNameSchema.optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  src: rawImageSourceSchema,
  /** Alternative text for the image. */
  alt: z.string().optional(),
  geometry: geometrySchema.optional(),
  style: rawImageStyleSchema.optional(),
  /** `fill` stretches, `cover` crops centered, `contain` preserves aspect ratio; defaults to `fill`. */
  fit: rawImageFitSchema.optional(),
  /** Explicit crop fraction for each side of the original image; takes precedence over `fit` when provided. */
  crop: rawImageCropInputSchema.optional(),
  ...motionPropsShape,
});

export const imageAltSchema = imageDeclarationPropsObjectSchema.shape.alt;
const rawImageDeclarationPropsSchema =
  imageDeclarationPropsObjectSchema.readonly();
export const imageDeclarationPropsSchema: z.ZodType<ImageDeclarationProps> =
  rawImageDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      ImageDeclarationProps,
      z.infer<typeof rawImageDeclarationPropsSchema>
    > extends true
      ? ImageDeclarationProps
      : never
  >;

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
  w: ooxmlPositiveIntSchema.refine((v) => v <= 2_147_483_647).optional(),
  h: ooxmlPositiveIntSchema.refine((v) => v <= 2_147_483_647).optional(),
  fill: z
    .enum(["none", "norm", "lighten", "lightenLess", "darken", "darkenLess"])
    .optional(),
  stroke: z.enum(["0", "1", "true", "false"]).optional(),
});

export const customPathPointSchema = z.object({
  x: ooxmlIntSchema.refine((v) => v >= -2_147_483_647 && v <= 2_147_483_647),
  y: ooxmlIntSchema.refine((v) => v >= -2_147_483_647 && v <= 2_147_483_647),
});

export const lineAttributesFragmentSchema = z.object({
  w: ooxmlPositiveLineWidthEmuSchema.optional(),
});

export const presetDashFragmentSchema = z.object({
  val: z.enum(["solid", "dash", "sysDot", "dot"]),
});

export const outerShadowFragmentSchema = z.object({
  blurRad: ooxmlWritableExtentEmuSchema.optional(),
  dist: ooxmlWritableExtentEmuSchema.optional(),
  dir: ooxmlIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: z.string().optional(),
});

export const softEdgeFragmentSchema = z.object({
  rad: ooxmlWritableExtentEmuSchema,
});
