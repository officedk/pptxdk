import type { MorphKey } from "@react-pptx/core";

/**
 * 原生对象可用的显式跨页连续性声明。
 *
 * Morph 会把句柄写入 OOXML 原生对象名称，因此同一对象不能再另设 name。
 */
export type MorphProps = {
  readonly morphKey?: MorphKey;
};

export { createMorphScope } from "@react-pptx/core";
export type { MorphKey, MorphScope } from "@react-pptx/core";
