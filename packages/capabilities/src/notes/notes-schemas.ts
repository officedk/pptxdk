import { authorIdSchema, textStyleSchema } from "@pptxdk/core/capability";
import { ooxmlBoolSchema, ooxmlNonNegativeIntSchema } from "@pptxdk/import";
import { z } from "zod";

/** Pure-data Zod v4 contract for page-affiliated speaker notes declarations (`Notes` / `notes()`). */
export const notesDeclarationPropsSchema = z
  .strictObject({
    id: authorIdSchema.optional(),
    style: textStyleSchema.optional(),
  })
  .readonly();

export type NotesDeclarationProps = z.infer<typeof notesDeclarationPropsSchema>;

/** Validates `<p:notes>` root attributes in Transitional OOXML. */
export const notesRootAttrsXmlSchema = z.object({
  showMasterSp: ooxmlBoolSchema.optional(),
  showMasterPhAnim: ooxmlBoolSchema.optional(),
});

/** Validates `<p:cNvPr>` attributes inside notes slide shapes. */
export const notesShapeIdentityXmlSchema = z.object({
  id: ooxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

/** Validates `<p:ph>` placeholder attributes inside notes slides. */
export const notesPlaceholderXmlSchema = z.object({
  type: z.string().optional(),
  idx: ooxmlNonNegativeIntSchema.optional(),
  sz: z.enum(["full", "half", "quarter"]).optional(),
  orient: z.enum(["horz", "vert"]).optional(),
  hasCustomPrompt: ooxmlBoolSchema.optional(),
});
