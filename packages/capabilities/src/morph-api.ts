import type { LinkTarget, MorphKey } from "@react-pptx/core";

/**
 * 原生对象可用的显式跨页连续性声明。
 *
 * Morph 会把句柄写入 OOXML 原生对象名称，因此同一对象不能再另设 name。
 */
export type MorphProps = {
  readonly morphKey?: MorphKey;
};

/** 公开原生对象的点击目标；核心负责把标识解析为文件关系。 */
export type LinkProps = {
  readonly link?: LinkTarget;
};

export { createMorphScope } from "@react-pptx/core";
export type { LinkTarget, MorphKey, MorphScope } from "@react-pptx/core";
