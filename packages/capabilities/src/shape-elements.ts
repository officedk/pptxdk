import { hostNodeCreator } from "@react-pptx/core/host";
import type { ReactNode } from "react";

import { motionElement } from "./motion-elements.ts";
import { createShapeNode, type ShapeDeclarationProps } from "./shape.ts";

export type ShapeProps = ShapeDeclarationProps & {
  /** Native text nodes shared with Text; may contain P, Span, and SlideNumber. */
  readonly children?: ReactNode;
};

/** Declares a native object handled by the shape module within the React tree. */
export function Shape(props: ShapeProps): ReactNode {
  return motionElement("rp:shape", {
    ...props,
    [hostNodeCreator]: createShapeNode,
  });
}
