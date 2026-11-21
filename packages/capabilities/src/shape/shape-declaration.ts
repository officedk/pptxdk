import type {
  ColorValue,
  LayoutStyle,
  LinkTarget,
  PaddingStyle,
} from "@pptxdk/core/capability";
import { RichTextContainer } from "@pptxdk/core/document";

import type { MotionProps } from "../motion/motion-api.ts";
import type { Geometry } from "../shared/geometry.ts";
import type { MorphProps } from "../shared/morph-api.ts";
import type { TextStyle } from "../shared/text-service.ts";
import type { ShapeVisualStyle } from "../shared/visual.ts";

export type ShapeDeclarationProps = MotionProps &
  MorphProps & {
    readonly link?: LinkTarget;
    readonly id?: string;
    readonly name?: string;
    readonly geometry?: Geometry;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly style?: LayoutStyle & ShapeVisualStyle & TextStyle & PaddingStyle;
    readonly fill?: ColorValue;
    /** Directly writes native text into this shape without creating an additional text box; mutually exclusive with children. */
    readonly text?: string;
    /** Shape text uses shared text services and can still be emitted when text box capability is omitted. */
    readonly textStyle?: TextStyle;
  };

/** Creates the editable native shape shared by the command factory and the Shape JSX element. */
export function createShapeNode(
  props: ShapeDeclarationProps,
): RichTextContainer<
  "shape",
  ShapeDeclarationProps,
  NonNullable<ShapeDeclarationProps["style"]>
> {
  return new RichTextContainer<
    "shape",
    ShapeDeclarationProps,
    NonNullable<ShapeDeclarationProps["style"]>
  >("shape", "shape", props);
}
