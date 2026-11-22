import {
  fontFaceSchema,
  fontFaceStyleSchema,
  fontOptionsSchema,
  fontSchema,
  type Font,
  type FontFace,
  type FontOptions,
} from "@pptxdk/core/capability";

export { fontFaceSchema, fontFaceStyleSchema, fontOptionsSchema, fontSchema };
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
  const candidate = {
    family,
    faces,
    ...(validOptions.substitute === true ? { substitute: true } : {}),
    ...(validOptions.embeddable === false ? { embeddable: false } : {}),
  };
  const parsedFont = fontSchema.safeParse(candidate);
  if (!parsedFont.success) {
    throw new TypeError(
      parsedFont.error.issues[0]?.message ?? "Invalid Font declaration",
    );
  }
  return Object.freeze({
    ...parsedFont.data,
    faces: Object.freeze(
      (parsedFont.data.faces ?? []).map((face) =>
        Object.freeze({
          ...face,
          ...(face.data === undefined
            ? {}
            : { data: new Uint8Array(face.data) }),
        }),
      ),
    ),
  });
}
