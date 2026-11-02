export type FontFaceStyle = "normal" | "italic";

/** 一种字体文件声明；资源只供构建中的字体用途读取，默认不会写入 PPTX。 */
export type FontFace = {
  readonly weight?: number;
  readonly style?: FontFaceStyle;
  /** 单字体 TTF/OTF、WOFF 或 WOFF2 数据；内容度量不读取系统字体。 */
  readonly data?: Uint8Array;
};

/** 可在主题和组件间引用的输出字体族、字重、样式及度量资源。 */
export type Font = {
  readonly family: string;
  readonly faces?: readonly FontFace[];
  /** 文件中的字体只供度量；输出仍声明 family。 */
  readonly substitute?: boolean;
};

export type FontOptions = {
  /** 明确允许用文件字体度量，同时保留 family 作为输出字体名。 */
  readonly substitute?: boolean;
};

/** 创建一个字体值；资源只从调用方提供的字节读取。 */
export function font(
  family: string,
  faces: readonly FontFace[] = [],
  options: FontOptions = {},
): Font {
  if (
    options === null ||
    typeof options !== "object" ||
    (options.substitute !== undefined &&
      typeof options.substitute !== "boolean")
  ) {
    throw new TypeError("Font options.substitute 必须是布尔值");
  }
  return Object.freeze({
    family,
    faces: Object.freeze(faces.map((face) => Object.freeze({ ...face }))),
    ...(options.substitute === true ? { substitute: true } : {}),
  });
}
