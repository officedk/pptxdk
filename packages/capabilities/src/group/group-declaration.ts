import {
  isMorphKey,
  linkTargetSchema,
  themeScopeSchema,
  type LayoutStyle,
  type LinkTarget,
  type MorphKey,
  type TextStyle,
  type ThemeScope,
  type VisualStyle,
} from "@pptxdk/core/capability";
import { CommonContainer } from "@pptxdk/core/document";
import { z } from "zod";

import type { MotionProps } from "../motion/motion-api.ts";
import type { MorphProps } from "../shared/morph-api.ts";

const groupDeclarationPropsObjectSchema = z.strictObject({
  id: z.string().optional(),
  name: z.string().optional(),
  morphKey: z.custom<MorphKey>(isMorphKey).optional(),
  link: linkTargetSchema.optional(),
  style: z.record(z.string(), z.unknown()).readonly().optional(),
  theme: themeScopeSchema.optional(),
  enter: z.unknown().optional(),
  emphasis: z.unknown().optional(),
  motion: z.unknown().optional(),
});

export const groupDeclarationPropsKeys = Object.keys(
  groupDeclarationPropsObjectSchema.shape,
);

export const groupDeclarationPropsSchema =
  groupDeclarationPropsObjectSchema.readonly();

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
  >("group", "group", props);
}
