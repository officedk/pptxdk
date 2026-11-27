import { defineObjectCodegen } from "@pptxdk/codegen";

import {
  blockRichTextChildKinds,
  createRichTextCodegenHandler,
} from "../text/text-codegen.ts";
const notesAllowedPropKeys = new Set(["id", "style"]);

/**
 * Source codegen contribution for page-affiliated `notes` (`notes()` / `<Notes>`) declarations.
 *
 * Registered via `declarations.notes` so `@pptxdk/codegen` resolves `Notes` child nodes inside
 * `Slide` while preserving all rich-text children and speaker-note styles.
 */
export const notesCodegen = defineObjectCodegen({
  declarations: {
    notes: createRichTextCodegenHandler({
      vanillaFactory: "notes",
      reactComponent: "Notes",
      allowedPropKeys: notesAllowedPropKeys,
      allowedChildKinds: blockRichTextChildKinds,
      supportObjectExtras: false,
    }),
  },
});
