import { hostNodeCreator } from "@react-pptx/core/document";
import { createElement, useContext, type Key, type ReactNode } from "react";

import { createBoxNode, type BoxDeclarationProps } from "./box.ts";
import { motionContext } from "./motion-elements.ts";

export type BoxProps = BoxDeclarationProps & {
  readonly key?: Key | null;
  readonly children?: ReactNode;
};

/** Visible Box's own background and descendants without entrance effects can inherit default entrance. */
export function Box({
  enter,
  motion,
  children,
  ...props
}: BoxProps): ReactNode {
  const parent = useContext(motionContext);
  const defaultEnter = enter === undefined ? parent.defaultEnter : enter;
  const context = Object.freeze({
    ...parent,
    enabled: parent.enabled && motion !== false,
    ...(defaultEnter === undefined ? {} : { defaultEnter }),
  });
  const boxContext = Object.freeze({
    ...context,
    ...(motion === undefined ? {} : { motionValue: motion }),
  });
  return createElement(
    motionContext.Provider,
    { value: context },
    createElement(
      "rp:box",
      {
        ...props,
        __motionContext: boxContext,
        [hostNodeCreator]: createBoxNode,
      },
      children,
    ),
  );
}
