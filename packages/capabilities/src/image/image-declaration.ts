import { CommonNode } from "@pptxdk/core/document";

import type {
  ExternalImageSource,
  ImageCropInput,
  ImageDeclarationProps,
  ImageFit,
  ImageSource,
  ImageVisualStyle,
} from "./image-schemas.ts";
import { imageVisualStyleKeys } from "./image-schemas.ts";

export {
  imageVisualStyleKeys,
  type ExternalImageSource,
  type ImageCropInput,
  type ImageDeclarationProps,
  type ImageFit,
  type ImageSource,
  type ImageVisualStyle,
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
  >("image", "image", props);
}
