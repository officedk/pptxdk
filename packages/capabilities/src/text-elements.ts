import type { LayoutStyle, VisualStyle } from "@react-pptx/core";
import { createElement, type ReactNode } from "react";

import type { MorphProps } from "./morph-api.ts";
import { motionElement, type MotionProps } from "./motion-api.ts";
import type { ParagraphStyle, TextStyle } from "./text-service.ts";

export type TextProps = MotionProps &
  MorphProps & {
    readonly testId?: string;
    readonly name?: string;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number | "auto";
    readonly style?: TextStyle & LayoutStyle & VisualStyle;
    readonly children?: ReactNode;
  };

export type ParagraphProps = {
  readonly style?: ParagraphStyle;
  readonly children?: ReactNode;
};

export type SpanProps = {
  readonly style?: Omit<TextStyle, "verticalAlign" | "lineHeight">;
  readonly children?: ReactNode;
};

/** 原生文本框；数值高度保留固定区域，auto 高度使用明确字体度量。 */
export function Text(props: TextProps): ReactNode {
  return motionElement("rp:text", props);
}

/** 在文本框中声明原生段落，可设置行距和段落列表语义。 */
export function P(props: ParagraphProps): ReactNode {
  return createElement("rp:paragraph", props, props.children);
}

/** 在段落中声明单独样式的文字片段。 */
export function Span(props: SpanProps): ReactNode {
  return createElement("rp:span", props, props.children);
}

/** Shape 的纯文字样式类型；它不依赖文本框组件能力。 */
export type ShapeTextStyle = TextStyle;
