import { hostNodeCreator } from "@pptxdk/core/host";
import { createElement, type Key, type ReactNode } from "react";

import { createBoxNode, type BoxDeclarationProps } from "./box-declaration.ts";

export type BoxProps = BoxDeclarationProps & {
  readonly key?: Key | null;
  readonly children?: ReactNode;
};

/**
 * Visible Box backgrounds and descendants inherit its entrance only when they do not declare one.
 *
 * The local `enter` and `motion` values stay on the editable node. Inheritance is derived by the next
 * formal build from the current document, so a moved Box does not leave a stale provider context.
 */
export function Box({ children, ...props }: BoxProps): ReactNode {
  return createElement(
    "rp:box",
    { ...props, [hostNodeCreator]: createBoxNode },
    children,
  );
}
