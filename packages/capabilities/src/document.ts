import type {
  LayoutStyle,
  LinkTarget,
  PaddingStyle,
  SlideBackground,
  TextStyle,
  ThemeInput,
  ThemeScope,
  VisualStyle,
} from "@react-pptx/core/capability";
import {
  CommonContainer,
  CommonNode,
  createDeckNode,
  createSectionNode,
  createSlideNode,
  RichTextContainer,
  type RichTextItem,
} from "@react-pptx/core/document";

import {
  boxCapability,
  createBoxNode,
  type BoxDeclarationProps,
} from "./box.ts";
import { font } from "./font.ts";
import {
  createGroupNode,
  groupCapability,
  type GroupDeclarationProps,
} from "./group.ts";
import {
  createImageNode,
  imageCapability,
  type ImageDeclarationProps,
} from "./image.ts";
import {
  createShapeNode,
  shapeCapability,
  type ShapeDeclarationProps,
} from "./shape.ts";
import { slideBackgroundCapability } from "./slide-background.ts";
import {
  createParagraphNode,
  createSlideNumberNode,
  createSpanNode,
  createTextNode,
  type ParagraphDeclarationProps,
  type SlideNumberDeclarationProps,
  type SpanDeclarationProps,
  type TextDeclarationProps,
} from "./text-declaration.ts";
import { textCapability, type ParsedText } from "./text.ts";
import { transitionCapability, type Transition } from "./transition.ts";

/**
 * External addresses such as link URLs are stored as address values at authoring time. Content
 * referenced by an address (including resource paths accepted by a capability) is read during each
 * formal build; creating an authoring node does not snapshot external file content.
 *
 * Image byte sources are isolated when accepted by `image()` or `set()`, so later changes to the
 * caller's buffer do not alter the document. Path and URL sources keep their address value instead
 * and are loaded again by every formal build.
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
  readonly background?: SlideBackground;
  readonly theme?: ThemeScope;
  readonly style?: TextStyle & Omit<LayoutStyle, "zIndex">;
  readonly hidden?: boolean;
  readonly advanceAfter?: number;
};

export type SectionAuthorProps = {
  readonly name: string;
};

export type BoxAuthorProps = BoxDeclarationProps;
export type ImageAuthorProps = ImageDeclarationProps;
export type GroupAuthorProps = GroupDeclarationProps;
export type ShapeAuthorProps = ShapeDeclarationProps;
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
export type ShapeNode = RichTextContainer<
  "shape",
  ShapeDeclarationProps,
  NonNullable<ShapeDeclarationProps["style"]>
>;
export type TextNode = RichTextContainer<
  "text",
  TextAuthorProps,
  LayoutStyle & PaddingStyle & TextStyle & VisualStyle
>;
export type ImageNode = CommonNode<
  "image",
  ImageDeclarationProps,
  NonNullable<ImageDeclarationProps["style"]>
>;
export type GroupNode = CommonContainer<
  "group",
  GroupDeclarationProps,
  NonNullable<GroupDeclarationProps["style"]>
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
  return createDeckNode(props);
}

export function slide(props: SlideAuthorProps): SlideNode {
  return createSlideNode(props);
}

export function section(props: SectionAuthorProps): SectionNode {
  return createSectionNode(props);
}

export function box(props: BoxAuthorProps = {}): BoxNode {
  return createBoxNode(props);
}

/** Creates a native shape whose geometry, visuals, and rich text remain editable after attachment. */
export function shape(props: ShapeAuthorProps = {}): ShapeNode {
  return createShapeNode(props);
}

/** Creates a native image whose source bytes, crop, fit, and visuals remain editable after attachment. */
export function image(props: ImageAuthorProps): ImageNode {
  return createImageNode(props);
}

/** Creates a native group preserving member identity, local theme, and layout children. */
export function group(props: GroupAuthorProps = {}): GroupNode {
  return createGroupNode(props);
}

/** Creates a text box whose original reference remains editable after attachment. */
export function text(content?: TextContent, props?: TextAuthorProps): TextNode;
export function text(...content: readonly TextContent[]): TextNode;
export function text(
  ...args: readonly (TextContent | TextAuthorProps | undefined)[]
): TextNode {
  const { content, props } = splitContent(args);
  const node = createTextNode(props ?? {});
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
  const node = createParagraphNode(props ?? {});
  node.content(...content);
  return node;
}

export function span(content?: SpanContent, props?: SpanAuthorProps): SpanNode;
export function span(...content: readonly SpanContent[]): SpanNode;
export function span(
  ...args: readonly (SpanContent | SpanAuthorProps | undefined)[]
): SpanNode {
  const { content, props } = splitContent(args);
  const node = createSpanNode(props ?? {});
  node.content(...content);
  return node;
}

/** Emits the native slide number field, resolved by the presentation application. */
export function slideNumber(
  props: SlideNumberAuthorProps = {},
): SlideNumberNode {
  return createSlideNumberNode(props);
}

/** Baseline capabilities enabled by the public document build; factories never register capabilities. */
export const documentCapabilities = Object.freeze([
  boxCapability,
  groupCapability,
  imageCapability,
  shapeCapability,
  slideBackgroundCapability,
  textCapability,
  transitionCapability,
] as const);

export { font };
export type { LinkTarget, ParsedText, Transition };
