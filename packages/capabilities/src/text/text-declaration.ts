import { CommonNode, RichTextContainer } from "@pptxdk/core/document";
import type { z } from "zod";

import type {
  paragraphAuthorPropsSchema,
  slideNumberAuthorPropsSchema,
  spanAuthorPropsSchema,
  textAuthorPropsSchema,
  TextStyle,
} from "../shared/text-schemas.ts";

/** Pure text box declaration shared by JSX and command authoring. */
export type TextDeclarationProps = z.infer<typeof textAuthorPropsSchema>;

export type ParagraphDeclarationProps = z.infer<
  typeof paragraphAuthorPropsSchema
>;

export type SpanDeclarationProps = z.infer<typeof spanAuthorPropsSchema>;

export type SlideNumberDeclarationProps = z.infer<
  typeof slideNumberAuthorPropsSchema
>;

/** Pure text style type for Shape; does not depend on the text box capability module. */
export type ShapeTextStyle = TextStyle;

/** Creates the editable text box shared by the command factory and the Text JSX element. */
export function createTextNode(
  props: TextDeclarationProps,
): RichTextContainer<
  "text",
  TextDeclarationProps,
  NonNullable<TextDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "text",
    TextDeclarationProps,
    NonNullable<TextDeclarationProps["style"]>
  >("text", "text", props);
}

/** Creates the editable paragraph shared by P JSX and command authoring. */
export function createParagraphNode(
  props: ParagraphDeclarationProps,
): RichTextContainer<
  "paragraph",
  ParagraphDeclarationProps,
  NonNullable<ParagraphDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "paragraph",
    ParagraphDeclarationProps,
    NonNullable<ParagraphDeclarationProps["style"]>
  >("paragraph", "paragraph", props);
}

/** Creates the editable span shared by Span JSX and command authoring. */
export function createSpanNode(
  props: SpanDeclarationProps,
): RichTextContainer<
  "span",
  SpanDeclarationProps,
  NonNullable<SpanDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "span",
    SpanDeclarationProps,
    NonNullable<SpanDeclarationProps["style"]>
  >("span", "span", props);
}

/** Creates the editable slide number field shared by the command factory and JSX element. */
export function createSlideNumberNode(
  props: SlideNumberDeclarationProps,
): CommonNode<
  "slide-number",
  SlideNumberDeclarationProps,
  NonNullable<SlideNumberDeclarationProps["style"]>
> {
  return new CommonNode<
    "slide-number",
    SlideNumberDeclarationProps,
    NonNullable<SlideNumberDeclarationProps["style"]>
  >("slide-number", "slide-number", props);
}
