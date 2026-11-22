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

const boxDeclarationPropsObjectSchema = z.strictObject({
  id: z.string().optional(),
  morphKey: z.custom<MorphKey>(isMorphKey).optional(),
  link: linkTargetSchema.optional(),
  style: z.record(z.string(), z.unknown()).readonly().optional(),
  theme: themeScopeSchema.optional(),
  enter: z.unknown().optional(),
  motion: z.unknown().optional(),
});

export const boxDeclarationPropsKeys = Object.keys(
  boxDeclarationPropsObjectSchema.shape,
);

export const boxDeclarationPropsSchema =
  boxDeclarationPropsObjectSchema.readonly();

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
  >("box", "box", props);
}
