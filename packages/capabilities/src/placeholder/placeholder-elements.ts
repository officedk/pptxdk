import { hostNodeCreator } from "@pptxdk/core/host";
import { createElement, type ReactNode } from "react";

import {
  createPlaceholderNode,
  type PlaceholderDeclarationProps,
} from "./placeholder.ts";

export type PlaceholderProps = PlaceholderDeclarationProps;

/** Declares an unpopulated native content placeholder bound to a layout placeholder slot. */
export function Placeholder(props: PlaceholderProps): ReactNode {
  return createElement("placeholder", {
    ...props,
    [hostNodeCreator]: createPlaceholderNode,
  });
}
