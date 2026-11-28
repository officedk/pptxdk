import { hostNodeCreator } from "@pptxdk/core/host";
import { createElement, type ReactNode } from "react";

import { motionElement } from "../motion/motion-elements.ts";
import type { TypedElement } from "../shared/typed-element.ts";
import {
  createParagraphNode,
  createSlideNumberNode,
  createSpanNode,
  createSubNode,
  createSupNode,
  createTextNode,
  type ParagraphDeclarationProps,
  type SlideNumberDeclarationProps,
  type SpanDeclarationProps,
  type SubDeclarationProps,
  type SupDeclarationProps,
  type TextDeclarationProps,
} from "./text-declaration.ts";

export type {
  ParagraphDeclarationProps,
  ShapeTextStyle,
  SlideNumberDeclarationProps,
  SpanDeclarationProps,
  SubDeclarationProps,
  SupDeclarationProps,
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

export type SupProps = SupDeclarationProps & {
  readonly children?: ReactNode;
};

export type SubProps = SubDeclarationProps & {
  readonly children?: ReactNode;
};

export type SlideNumberProps = SlideNumberDeclarationProps;

export type { TypedElement } from "../shared/typed-element.ts";

/** Native text box; numeric height preserves a fixed region, while auto height uses explicit font measurement. */
export function Text(props: TextProps): TypedElement<"text"> {
  return motionElement("text", {
    ...props,
    [hostNodeCreator]: createTextNode,
  }) as TypedElement<"text">;
}

/** Declares a native paragraph within a text box, with line spacing and paragraph list semantics. */
export function P(props: ParagraphProps): TypedElement<"paragraph"> {
  return createElement(
    "paragraph",
    { ...props, [hostNodeCreator]: createParagraphNode },
    props.children,
  ) as unknown as TypedElement<"paragraph">;
}

/** Declares an individually styled text run within a paragraph. */
export function Span(props: SpanProps): TypedElement<"span"> {
  return createElement(
    "span",
    { ...props, [hostNodeCreator]: createSpanNode },
    props.children,
  ) as unknown as TypedElement<"span">;
}

/** Declares an inline superscript run (`baseline="30000"`). */
export function Sup(props: SupProps): TypedElement<"sup"> {
  return createElement(
    "sup",
    { ...props, [hostNodeCreator]: createSupNode },
    props.children,
  ) as unknown as TypedElement<"sup">;
}

/** Declares an inline subscript run (`baseline="-25000"`). */
export function Sub(props: SubProps): TypedElement<"sub"> {
  return createElement(
    "sub",
    { ...props, [hostNodeCreator]: createSubNode },
    props.children,
  ) as unknown as TypedElement<"sub">;
}

/** Emits a native slide number field resolved by presentation applications. */
export function SlideNumber(
  props: SlideNumberProps,
): TypedElement<"slide-number"> {
  return createElement("slide-number", {
    ...props,
    [hostNodeCreator]: createSlideNumberNode,
  }) as unknown as TypedElement<"slide-number">;
}

const inlineTextMarker = Symbol.for("@pptxdk/text-inline-element");
Object.defineProperty(P, inlineTextMarker, { value: true });
Object.defineProperty(Span, inlineTextMarker, { value: true });
Object.defineProperty(Sup, inlineTextMarker, { value: true });
Object.defineProperty(Sub, inlineTextMarker, { value: true });
Object.defineProperty(SlideNumber, inlineTextMarker, { value: true });
