import type { ReactNode } from "react";

import type { ImageDeclarationProps } from "./image.ts";
import { motionElement } from "./motion-elements.ts";

export type ImageProps = ImageDeclarationProps & {
  readonly children?: ReactNode;
};

/** Declares a native image handled by the image module within the React tree. */
export function Image(props: ImageProps): ReactNode {
  return motionElement("rp:image", props);
}
