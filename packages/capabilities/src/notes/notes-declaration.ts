import { RichTextContainer } from "@pptxdk/core/document";

import type { TextStyle } from "../shared/text-service.ts";
import {
  notesDeclarationPropsSchema,
  type NotesDeclarationProps,
} from "./notes-schemas.ts";

export { notesDeclarationPropsSchema, type NotesDeclarationProps };

export type NotesNode = RichTextContainer<
  "notes",
  NotesDeclarationProps,
  TextStyle
>;

/** Creates the editable notes container shared by the command factory and Notes JSX. */
export function createNotesNode(props: NotesDeclarationProps): NotesNode {
  return new RichTextContainer<"notes", NotesDeclarationProps, TextStyle>(
    "notes",
    "notes",
    props,
    { pageAffiliated: true },
  );
}
