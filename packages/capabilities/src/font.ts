import type { Font, FontFace, FontOptions } from "@react-pptx/core/capability";

export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
} from "@react-pptx/core/capability";

/** Creates a Font value; resources are read only from caller-provided bytes. */
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
    throw new TypeError("Font options.substitute must be a boolean");
  }
  return Object.freeze({
    family,
    faces: Object.freeze(
      faces.map((face) =>
        Object.freeze({
          ...face,
          ...(face.data === undefined
            ? {}
            : { data: new Uint8Array(face.data) }),
        }),
      ),
    ),
    ...(options.substitute === true ? { substitute: true } : {}),
    ...(options.embeddable === false ? { embeddable: false } : {}),
  });
}
