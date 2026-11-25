import { hostNodeCreator } from "@pptxdk/core/host";
import { createElement, type ReactNode } from "react";

import {
  createNotesNode,
  type NotesDeclarationProps,
} from "./notes-declaration.ts";

export type NotesProps = NotesDeclarationProps & {
  readonly children?: ReactNode;
};

/** Declares one page's rich-text speaker notes; repeated declarations on a page are a build error. */
export function Notes(props: NotesProps): ReactNode {
  return createElement(
    "notes",
    { ...props, [hostNodeCreator]: createNotesNode },
    props.children,
  );
}
