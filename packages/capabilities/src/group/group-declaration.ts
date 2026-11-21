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

/** Native Group preserves member relationships; style allows it to arrange as a layout child, with output bounds computed from the union of member geometries. */
export type GroupDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly style?: LayoutStyle & TextStyle & Pick<VisualStyle, "rotate">;
    readonly theme?: ThemeScope;
  };

/** Creates the editable native group shared by the command factory and the Group JSX element. */
export function createGroupNode(
  props: GroupDeclarationProps,
): CommonContainer<
  "group",
  GroupDeclarationProps,
  NonNullable<GroupDeclarationProps["style"]>
> {
  return new CommonContainer<
    "group",
    GroupDeclarationProps,
    NonNullable<GroupDeclarationProps["style"]>
  >("group", "rp:group", props);
}
