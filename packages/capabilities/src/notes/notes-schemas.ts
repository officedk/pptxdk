import {
  authorIdSchema,
  textStyleSchema,
  type AuthorId,
  type TextStyle,
} from "@pptxdk/core/capability";
import { ooxmlBoolSchema, ooxmlNonNegativeIntSchema } from "@pptxdk/import";
import { z } from "zod";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type NotesDeclarationProps = {
  readonly id?: AuthorId | undefined;
  readonly style?: TextStyle | undefined;
};

const rawNotesDeclarationPropsSchema = z
  .strictObject({
    id: authorIdSchema.optional(),
    style: textStyleSchema.optional(),
  })
  .readonly();

/** Allowed declaration prop keys for `Notes` / `notes()`, derived from `rawNotesDeclarationPropsSchema`. */
export const notesAllowedPropKeys: ReadonlySet<string> = new Set(
  Object.keys(rawNotesDeclarationPropsSchema.unwrap().shape),
);

/** Pure-data Zod v4 contract for page-affiliated speaker notes declarations (`Notes` / `notes()`). */
export const notesDeclarationPropsSchema: z.ZodType<NotesDeclarationProps> =
  rawNotesDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      NotesDeclarationProps,
      z.infer<typeof rawNotesDeclarationPropsSchema>
    > extends true
      ? NotesDeclarationProps
      : never
  >;

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
