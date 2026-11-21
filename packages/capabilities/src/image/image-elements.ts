import { hostNodeCreator } from "@pptxdk/core/host";
import type { ReactNode } from "react";

import { motionElement } from "../motion/motion-elements.ts";
import { createImageNode, type ImageDeclarationProps } from "./image.ts";

export type ImageProps = ImageDeclarationProps & {
  readonly children?: ReactNode;
};

/** Declares a native image handled by the image module within the React tree. */
export function Image(props: ImageProps): ReactNode {
  return motionElement("image", {
    ...props,
    [hostNodeCreator]: createImageNode,
  });
}
