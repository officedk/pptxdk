import { hostNodeCreator } from "@pptxdk/core/host";
import type { ReactNode } from "react";

import { motionElement } from "../motion/motion-elements.ts";
import type { TypedElement } from "../shared/typed-element.ts";
import {
  createImageNode,
  type ImageDeclarationProps,
} from "./image-declaration.ts";

export type { TypedElement } from "../shared/typed-element.ts";

export type ImageProps = ImageDeclarationProps & {
  readonly children?: ReactNode;
};

/** Declares a native image handled by the image module within the React tree. */
export function Image(props: ImageProps): TypedElement<"image"> {
  return motionElement("image", {
    ...props,
    [hostNodeCreator]: createImageNode,
  }) as TypedElement<"image">;
}
