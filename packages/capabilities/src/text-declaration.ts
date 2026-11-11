import type {
  LayoutStyle,
  LinkTarget,
  PaddingStyle,
  VisualStyle,
} from "@react-pptx/core/capability";

import type { MorphProps } from "./morph-api.ts";
import type { MotionProps } from "./motion-api.ts";
import type { ParagraphStyle, TextStyle } from "./text-service.ts";

/** Pure text box declaration shared by JSX and command authoring. */
export type TextDeclarationProps = MotionProps &
  MorphProps & {
    readonly testId?: string;
    readonly part?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number | "auto";
    readonly style?: LayoutStyle & PaddingStyle & TextStyle & VisualStyle;
  };

export type ParagraphDeclarationProps = {
  readonly style?: ParagraphStyle;
};

export type SpanDeclarationProps = {
  readonly style?: Omit<TextStyle, "verticalAlign" | "lineHeight" | "textWrap">;
};

export type SlideNumberDeclarationProps = {
  readonly style?: Omit<TextStyle, "verticalAlign" | "lineHeight" | "textWrap">;
};

/** Pure text style type for Shape; does not depend on the text box capability module. */
export type ShapeTextStyle = TextStyle;
