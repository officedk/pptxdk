import {
  fontOptionsSchema,
  type Font,
  type FontFace,
  type FontOptions,
} from "@pptxdk/core/capability";

export type {
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
} from "@pptxdk/core/capability";

/** Creates a Font value; resources are read only from caller-provided bytes. */
export function font(
  family: string,
  faces: readonly FontFace[] = [],
  options: FontOptions = {},
): Font {
  const parsedOptions = fontOptionsSchema.safeParse(options);
  if (!parsedOptions.success) {
    throw new TypeError("Font options.substitute must be a boolean");
  }
  const validOptions = parsedOptions.data;
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
    ...(validOptions.substitute === true ? { substitute: true } : {}),
    ...(validOptions.embeddable === false ? { embeddable: false } : {}),
  });
}
