import { RichTextContainer } from "@react-pptx/core/document";

import type { TextStyle } from "./text-service.ts";

/**
 * Page-affiliated speaker notes declaration shared by JSX and command authoring.
 *
 * Notes carry only an author id and shared rich-text style/content. Layout, native object,
 * Morph, link, and motion properties belong to slide canvas objects, not to the notes body.
 */
export type NotesDeclarationProps = {
  readonly id?: string;
  readonly style?: TextStyle;
};

export type NotesNode = RichTextContainer<
  "notes",
  NotesDeclarationProps,
  TextStyle
>;

/** Creates the editable notes container shared by the command factory and Notes JSX. */
export function createNotesNode(props: NotesDeclarationProps): NotesNode {
  return new RichTextContainer<"notes", NotesDeclarationProps, TextStyle>(
    "notes",
    "rp:notes",
    props,
  );
}
