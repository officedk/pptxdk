import { hostNodeCreator } from "@react-pptx/core/host";
import type { ReactNode } from "react";

import { createGroupNode, type GroupDeclarationProps } from "./group.ts";
import { motionElement } from "./motion-elements.ts";

export type GroupProps = GroupDeclarationProps & {
  readonly children?: ReactNode;
};

/** Declares an editable group preserving native member structure within the React tree. */
export function Group(props: GroupProps): ReactNode {
  return motionElement("rp:group", {
    ...props,
    [hostNodeCreator]: createGroupNode,
  });
}
