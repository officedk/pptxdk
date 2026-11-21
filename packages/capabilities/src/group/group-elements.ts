import { hostNodeCreator } from "@pptxdk/core/host";
import type { ReactNode } from "react";

import { motionElement } from "../motion/motion-elements.ts";
import { createGroupNode, type GroupDeclarationProps } from "./group.ts";

export type GroupProps = GroupDeclarationProps & {
  readonly children?: ReactNode;
};

/** Declares an editable group preserving native member structure within the React tree. */
export function Group(props: GroupProps): ReactNode {
  return motionElement("group", {
    ...props,
    [hostNodeCreator]: createGroupNode,
  });
}
