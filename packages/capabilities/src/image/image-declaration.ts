import type {
  LayoutStyle,
  LinkTarget,
  VisualStyle,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";

import type { MotionProps } from "../motion/motion-api.ts";
import type { Geometry } from "../shared/geometry.ts";
import type { MorphProps } from "../shared/morph-api.ts";

/** Visual styling supported on Image objects. */
export type ImageVisualStyle = Pick<
  VisualStyle,
  | "borderColor"
  | "borderWidth"
  | "borderStyle"
  | "opacity"
  | "boxShadow"
  | "softEdge"
  | "rotate"
>;

export const imageVisualStyleKeys = new Set([
  "borderColor",
  "borderWidth",
  "borderStyle",
  "opacity",
  "boxShadow",
  "softEdge",
  "rotate",
]);

/** Local file path, file URL, or raw image bytes. */
export type ImageSource = string | URL | Uint8Array;

export type ImageDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly src: ImageSource;
    /** Alternative text for the image. */
    readonly alt?: string;
    /** Position and bounds in px; position defaults to 0, width and height default to original image dimensions. */
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly geometry?: Geometry;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
    readonly style?: LayoutStyle & ImageVisualStyle;
    /** `fill` stretches, `cover` crops centered, `contain` preserves aspect ratio; defaults to `fill`. */
    readonly fit?: "fill" | "cover" | "contain";
    /** Explicit crop fraction for each side of the original image; takes precedence over `fit` when provided. */
    readonly crop?: {
      readonly left?: number;
      readonly top?: number;
      readonly right?: number;
      readonly bottom?: number;
    };
  };

/** Creates the editable native image shared by the command factory and the Image JSX element. */
export function createImageNode(
  props: ImageDeclarationProps,
): CommonNode<
  "image",
  ImageDeclarationProps,
  NonNullable<ImageDeclarationProps["style"]>
> {
  return new CommonNode<
    "image",
    ImageDeclarationProps,
    NonNullable<ImageDeclarationProps["style"]>
  >("image", "rp:image", props);
}
