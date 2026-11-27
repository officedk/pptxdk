import {
  bytesSha256,
  createSealedFontFace,
  fontFaceSchema,
  fontFaceStyleSchema,
  fontOptionsSchema,
  fontSchema,
  markImmutableBytes,
  markSealedFont,
  unsealedFontSchema,
  type Font,
  type FontFace,
  type FontOptions,
} from "@pptxdk/core/capability";

import { tryInferStaticFontFaceMetadata } from "./font-engine.ts";

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
  const parsedFont = unsealedFontSchema.safeParse(candidate);
  if (!parsedFont.success) {
    throw new TypeError(
      parsedFont.error.issues[0]?.message ?? "Invalid Font declaration",
    );
  }
  const sealedFaces = Object.freeze(
    (parsedFont.data.faces ?? []).map((face) => {
      if (face.data === undefined) {
        return createSealedFontFace({
          weight: face.weight,
          style: face.style,
        });
      }
      // Perform the single isolating copy of caller bytes and compute its SHA-256 once.
      const copiedData = markImmutableBytes(new Uint8Array(face.data));
      const sha256 = bytesSha256(copiedData);
      const inferred =
        face.weight === undefined || face.style === undefined
          ? tryInferStaticFontFaceMetadata(copiedData, sha256)
          : undefined;
      return createSealedFontFace({
        weight: face.weight ?? inferred?.weight,
        style: face.style ?? inferred?.style,
        bytes: copiedData,
        sha256,
      });
    }),
  );
  return markSealedFont({
    ...parsedFont.data,
    faces: sealedFaces,
  });
}
