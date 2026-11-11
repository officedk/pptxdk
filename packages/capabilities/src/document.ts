import type {
  LayoutStyle,
  LinkTarget,
  PaddingStyle,
  TextStyle,
  ThemeInput,
  ThemeScope,
  VisualStyle,
} from "@react-pptx/core/capability";
import {
  CommonContainer,
  CommonNode,
  RichTextContainer,
  type RichTextItem,
} from "@react-pptx/core/document";

import { boxCapability, type BoxDeclarationProps } from "./box.ts";
import { font } from "./font.ts";
import type {
  ParagraphDeclarationProps,
  SlideNumberDeclarationProps,
  SpanDeclarationProps,
  TextDeclarationProps,
} from "./text-declaration.ts";
import { textCapability, type ParsedText } from "./text.ts";
import { transitionCapability, type Transition } from "./transition.ts";

/**
 * External addresses such as link URLs are stored as address values at authoring time. Content
 * referenced by an address (including resource paths accepted by a capability) is read during each
 * formal build; creating an authoring node does not snapshot external file content.
 */

export type DeckAuthorProps = {
  readonly width?: number;
  readonly height?: number;
  readonly theme?: ThemeInput;
  readonly style?: TextStyle;
  readonly title?: string;
};

export type SlideAuthorProps = {
  readonly slideId: string;
  readonly transition?: Transition;
  readonly theme?: ThemeScope;
  readonly style?: TextStyle & Omit<LayoutStyle, "zIndex">;
  readonly hidden?: boolean;
  readonly advanceAfter?: number;
};

export type SectionAuthorProps = {
  readonly name: string;
};

export type BoxAuthorProps = BoxDeclarationProps;
export type TextAuthorProps = TextDeclarationProps;
export type ParagraphAuthorProps = ParagraphDeclarationProps;
export type SpanAuthorProps = SpanDeclarationProps;
export type SlideNumberAuthorProps = SlideNumberDeclarationProps;

export type DeckNode = CommonContainer<"deck", DeckAuthorProps, TextStyle>;
export type SlideNode = CommonContainer<
  "slide",
  SlideAuthorProps,
  TextStyle & Omit<LayoutStyle, "zIndex">
>;
export type SectionNode = CommonContainer<"section", SectionAuthorProps>;
export type BoxNode = CommonContainer<
  "box",
  BoxAuthorProps,
  LayoutStyle & TextStyle & VisualStyle
>;
export type TextNode = RichTextContainer<
  "text",
  TextAuthorProps,
  LayoutStyle & PaddingStyle & TextStyle & VisualStyle
>;
export type ParagraphNode = RichTextContainer<
  "paragraph",
  ParagraphAuthorProps,
  NonNullable<ParagraphDeclarationProps["style"]>
>;
export type SpanNode = RichTextContainer<
  "span",
  SpanAuthorProps,
  NonNullable<SpanDeclarationProps["style"]>
>;
export type SlideNumberNode = CommonNode<
  "slide-number",
  SlideNumberAuthorProps,
  NonNullable<SlideNumberDeclarationProps["style"]>
>;

export type TextContent = RichTextItem | readonly RichTextItem[];
export type ParagraphContent = RichTextItem | readonly RichTextItem[];
export type SpanContent = RichTextItem | readonly RichTextItem[];

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  if (value instanceof CommonNode) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function splitContent<Props extends object>(
  args: readonly (TextContent | Props | undefined)[],
): {
  readonly content: readonly RichTextItem[];
  readonly props: Props | undefined;
} {
  const last = args.at(-1);
  const hasProps = isPlainRecord(last);
  const values = hasProps ? args.slice(0, -1) : args;
  const content = values.flatMap((value) =>
    value === undefined
      ? []
      : Array.isArray(value)
        ? value
        : [value as RichTextItem],
  );
  return {
    content,
    props: hasProps ? (last as Props) : undefined,
  };
}

export function deck(props: DeckAuthorProps = {}): DeckNode {
  return new CommonContainer("deck", "rp:deck", props);
}

export function slide(props: SlideAuthorProps): SlideNode {
  return new CommonContainer("slide", "rp:slide", props);
}

export function section(props: SectionAuthorProps): SectionNode {
  return new CommonContainer("section", "rp:section", props);
}

export function box(props: BoxAuthorProps = {}): BoxNode {
  return new CommonContainer("box", "rp:box", props);
}

/** Creates a text box whose original reference remains editable after attachment. */
export function text(content?: TextContent, props?: TextAuthorProps): TextNode;
export function text(...content: readonly TextContent[]): TextNode;
export function text(
  ...args: readonly (TextContent | TextAuthorProps | undefined)[]
): TextNode {
  const { content, props } = splitContent(args);
  const node = new RichTextContainer<
    "text",
    TextAuthorProps,
    LayoutStyle & PaddingStyle & TextStyle & VisualStyle
  >("text", "rp:text", props ?? {});
  node.content(...content);
  return node;
}

export function paragraph(
  content?: ParagraphContent,
  props?: ParagraphAuthorProps,
): ParagraphNode;
export function paragraph(
  ...content: readonly ParagraphContent[]
): ParagraphNode;
export function paragraph(
  ...args: readonly (ParagraphContent | ParagraphAuthorProps | undefined)[]
): ParagraphNode {
  const { content, props } = splitContent(args);
  const node = new RichTextContainer<
    "paragraph",
    ParagraphAuthorProps,
    NonNullable<ParagraphDeclarationProps["style"]>
  >("paragraph", "rp:paragraph", props ?? {});
  node.content(...content);
  return node;
}

export function span(content?: SpanContent, props?: SpanAuthorProps): SpanNode;
export function span(...content: readonly SpanContent[]): SpanNode;
export function span(
  ...args: readonly (SpanContent | SpanAuthorProps | undefined)[]
): SpanNode {
  const { content, props } = splitContent(args);
  const node = new RichTextContainer<
    "span",
    SpanAuthorProps,
    NonNullable<SpanDeclarationProps["style"]>
  >("span", "rp:span", props ?? {});
  node.content(...content);
  return node;
}

/** Emits the native slide number field, resolved by the presentation application. */
export function slideNumber(
  props: SlideNumberAuthorProps = {},
): SlideNumberNode {
  return new CommonNode<
    "slide-number",
    SlideNumberAuthorProps,
    NonNullable<SlideNumberDeclarationProps["style"]>
  >("slide-number", "rp:slide-number", props);
}

/** Baseline capabilities enabled by the public document build; factories never register capabilities. */
export const documentCapabilities = Object.freeze([
  boxCapability,
  textCapability,
  transitionCapability,
] as const);

export { font };
export type { LinkTarget, ParsedText, Transition };
