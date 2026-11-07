import type {
  LayoutStyle,
  LinkTarget,
  PaddingStyle,
  VisualStyle as CoreVisualStyle,
} from "@react-pptx/core";
import { createElement, type ReactNode } from "react";

import type { MorphProps } from "./morph-api.ts";
import { motionElement, type MotionProps } from "./motion-api.ts";
import type { ParagraphStyle, TextStyle } from "./text-service.ts";

export type TextProps = MotionProps &
  MorphProps & {
    readonly testId?: string;
    readonly part?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number | "auto";
    readonly style?: LayoutStyle & PaddingStyle & TextStyle & CoreVisualStyle;
    readonly children?: ReactNode;
  };

export type ParagraphProps = {
  readonly style?: ParagraphStyle;
  readonly children?: ReactNode;
};

export type SpanProps = {
  readonly style?: Omit<TextStyle, "verticalAlign" | "lineHeight" | "textWrap">;
  readonly children?: ReactNode;
};

export type SlideNumberProps = {
  readonly style?: Omit<TextStyle, "verticalAlign" | "lineHeight" | "textWrap">;
};

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

/** Pure text style type for Shape; does not depend on the text box capability module. */
export type ShapeTextStyle = TextStyle;
