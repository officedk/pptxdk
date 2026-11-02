export type FontFaceStyle = "normal" | "italic";

/** 一种字体文件声明；资源只供构建中的字体用途读取，默认不会写入 PPTX。 */
export type FontFace = {
  readonly weight?: number;
  readonly style?: FontFaceStyle;
  readonly data?: Uint8Array;
};

/** 可在主题和组件间引用的字体族及其字重、样式资源。 */
export type Font = {
  readonly family: string;
  readonly faces?: readonly FontFace[];
};

/** 创建一个字体值；调用方显式传入资源，不触发下载或系统字体扫描。 */
export function font(family: string, faces: readonly FontFace[] = []): Font {
  return Object.freeze({
    family,
    faces: Object.freeze(faces.map((face) => Object.freeze({ ...face }))),
  });
}
