import { textStyleSchema } from "@pptxdk/core/capability";
import { OoxmlBoolSchema, OoxmlNonNegativeIntSchema } from "@pptxdk/import";
import { z } from "zod";

/** Pure-data Zod v4 contract for page-affiliated speaker notes declarations (`Notes` / `notes()`). */
export const notesDeclarationPropsSchema = z
  .strictObject({
    id: z.string().optional(),
    style: textStyleSchema.optional(),
  })
  .readonly();

export type NotesDeclarationProps = z.infer<typeof notesDeclarationPropsSchema>;

/** Validates `<p:notes>` root attributes in Transitional OOXML. */
export const NotesRootAttrsXmlSchema = z.object({
  showMasterSp: OoxmlBoolSchema.optional(),
  showMasterPhAnim: OoxmlBoolSchema.optional(),
});

/** Validates `<p:cNvPr>` attributes inside notes slide shapes. */
export const NotesShapeIdentityXmlSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

/** Validates `<p:ph>` placeholder attributes inside notes slides. */
export const NotesPlaceholderXmlSchema = z.object({
  type: z.string().optional(),
  idx: OoxmlNonNegativeIntSchema.optional(),
  sz: z.enum(["full", "half", "quarter"]).optional(),
  orient: z.enum(["horz", "vert"]).optional(),
  hasCustomPrompt: OoxmlBoolSchema.optional(),
});
