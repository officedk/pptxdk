import type { Font, FontFace, FontOptions } from "@react-pptx/core";

export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
} from "@react-pptx/core";

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
      typeof options.substitute !== "boolean") ||
    (options.embeddable !== undefined &&
      typeof options.embeddable !== "boolean")
  ) {
    throw new TypeError("Font options.substitute 必须是布尔值");
  }
  return Object.freeze({
    family,
    faces: Object.freeze(faces.map((face) => Object.freeze({ ...face }))),
    ...(options.substitute === true ? { substitute: true } : {}),
    ...(options.embeddable === false ? { embeddable: false } : {}),
  });
}
