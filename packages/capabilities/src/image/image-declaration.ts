import { CommonNode } from "@pptxdk/core/document";
import type { z } from "zod";

import type { MotionProps } from "../motion/motion-api.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import type {
  ExternalImageSource,
  ImageCropInput,
  ImageFit,
  ImageSource,
  ImageVisualStyle,
  imageDeclarationPropsSchema,
} from "./image-schemas.ts";
import { imageVisualStyleKeys } from "./image-schemas.ts";

export {
  imageVisualStyleKeys,
  type ExternalImageSource,
  type ImageCropInput,
  type ImageFit,
  type ImageSource,
  type ImageVisualStyle,
};

type ImageDeclarationDataProps = Omit<
  z.infer<typeof imageDeclarationPropsSchema>,
  "enter" | "emphasis" | "motion" | "morphKey"
>;

export type ImageDeclarationProps = MotionProps &
  MorphProps &
  ImageDeclarationDataProps;

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
  >("image", "image", props);
}
