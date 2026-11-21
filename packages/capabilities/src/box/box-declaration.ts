import type {
  LayoutStyle,
  LinkTarget,
  TextStyle,
  ThemeScope,
  VisualStyle,
} from "@pptxdk/core/capability";
import { CommonContainer } from "@pptxdk/core/document";

import type { MotionProps } from "../motion/motion-api.ts";
import type { MorphProps } from "../shared/morph-api.ts";

/** Layout container appearance with visual styling that optionally generates a native background shape. */
export type BoxDeclarationProps = MorphProps &
  Pick<MotionProps, "enter" | "motion"> & {
    readonly id?: string;
    readonly link?: LinkTarget;
    readonly style?: LayoutStyle & TextStyle & VisualStyle;
    readonly theme?: ThemeScope;
  };

/** Creates the editable layout container shared by the command factory and the Box JSX element. */
export function createBoxNode(
  props: BoxDeclarationProps,
): CommonContainer<
  "box",
  BoxDeclarationProps,
  NonNullable<BoxDeclarationProps["style"]>
> {
  return new CommonContainer<
    "box",
    BoxDeclarationProps,
    NonNullable<BoxDeclarationProps["style"]>
  >("box", "rp:box", props);
}
