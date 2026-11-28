import { hostNodeCreator } from "@pptxdk/core/host";

import { motionElement } from "../motion/motion-elements.ts";
import type { TypedElement } from "../shared/typed-element.ts";
import {
  createLineNode,
  type LineDeclarationProps,
} from "./line-declaration.ts";

export type { TypedElement } from "../shared/typed-element.ts";

export type LineProps = LineDeclarationProps & {
  /** Lines are 1D primitives and do not support nested children or text. */
  readonly children?: never;
};

/** Declares a native `<p:cxnSp>` line or connector within the React tree. */
export function Line(props: LineProps): TypedElement<"line"> {
  return motionElement("line", {
    ...props,
    [hostNodeCreator]: createLineNode,
  }) as TypedElement<"line">;
}
