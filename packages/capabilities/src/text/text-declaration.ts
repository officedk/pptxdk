import { CommonNode, RichTextContainer } from "@pptxdk/core/document";

import type {
  ParagraphDeclarationProps,
  SlideNumberDeclarationProps,
  SpanDeclarationProps,
  SubDeclarationProps,
  SupDeclarationProps,
  TextDeclarationProps,
  TextStyle,
} from "../shared/text-schemas.ts";

/** Pure text box declaration shared by JSX and command authoring. */
export type {
  ParagraphDeclarationProps,
  SlideNumberDeclarationProps,
  SpanDeclarationProps,
  SubDeclarationProps,
  SupDeclarationProps,
  TextDeclarationProps,
};

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

/** Creates the editable superscript inline node shared by Sup JSX and command authoring. */
export function createSupNode(
  props: SupDeclarationProps,
): RichTextContainer<
  "sup",
  SupDeclarationProps,
  NonNullable<SupDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "sup",
    SupDeclarationProps,
    NonNullable<SupDeclarationProps["style"]>
  >("sup", "sup", props);
}

/** Creates the editable subscript inline node shared by Sub JSX and command authoring. */
export function createSubNode(
  props: SubDeclarationProps,
): RichTextContainer<
  "sub",
  SubDeclarationProps,
  NonNullable<SubDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "sub",
    SubDeclarationProps,
    NonNullable<SubDeclarationProps["style"]>
  >("sub", "sub", props);
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
