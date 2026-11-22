import type {
  LayoutStyle,
  LinkTarget,
  NativePlaceholderRef,
  PaddingStyle,
  VisualStyle,
} from "@pptxdk/core/capability";
import { CommonNode, RichTextContainer } from "@pptxdk/core/document";

import type { MotionProps } from "../motion/motion-api.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import type {
  ParagraphStyle,
  RunTextStyle,
  TextStyle,
} from "../shared/text-schemas.ts";

/** Pure text box declaration shared by JSX and command authoring. */
export type TextDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly placeholder?: string;
    readonly nativePlaceholder?: NativePlaceholderRef;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number | "auto";
    readonly style?: LayoutStyle & PaddingStyle & TextStyle & VisualStyle;
  };

export type ParagraphDeclarationProps = {
  readonly id?: string;
  readonly style?: ParagraphStyle;
};

export type SpanDeclarationProps = {
  readonly id?: string;
  readonly style?: RunTextStyle;
};

export type SlideNumberDeclarationProps = {
  readonly id?: string;
  readonly style?: RunTextStyle;
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
