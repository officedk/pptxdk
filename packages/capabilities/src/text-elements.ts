import { createElement, type ReactNode } from "react";

import { motionElement } from "./motion-elements.ts";
import type {
  ParagraphDeclarationProps,
  SlideNumberDeclarationProps,
  SpanDeclarationProps,
  TextDeclarationProps,
} from "./text-declaration.ts";

export type {
  ParagraphDeclarationProps,
  ShapeTextStyle,
  SlideNumberDeclarationProps,
  SpanDeclarationProps,
  TextDeclarationProps,
} from "./text-declaration.ts";

export type TextProps = TextDeclarationProps & {
  readonly children?: ReactNode;
};

export type ParagraphProps = ParagraphDeclarationProps & {
  readonly children?: ReactNode;
};

export type SpanProps = SpanDeclarationProps & {
  readonly children?: ReactNode;
};

export type SlideNumberProps = SlideNumberDeclarationProps;

/** Native text box; numeric height preserves a fixed region, while auto height uses explicit font measurement. */
export function Text(props: TextProps): ReactNode {
  return motionElement("rp:text", props);
}

/** Declares a native paragraph within a text box, with line spacing and paragraph list semantics. */
export function P(props: ParagraphProps): ReactNode {
  return createElement("rp:paragraph", props, props.children);
}

/** Declares an individually styled text run within a paragraph. */
export function Span(props: SpanProps): ReactNode {
  return createElement("rp:span", props, props.children);
}

/** Emits a native slide number field resolved by presentation applications. */
export function SlideNumber(props: SlideNumberProps): ReactNode {
  return createElement("rp:slide-number", props);
}
