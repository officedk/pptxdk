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

import {
  tryInferStaticFontFaceMetadata,
  unpackFontCollectionFaces,
} from "./font-engine.ts";

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
      const inferred = tryInferStaticFontFaceMetadata(copiedData, sha256);
      return createSealedFontFace({
        weight: face.weight ?? inferred?.weight,
        style: face.style ?? inferred?.style,
        bytes: copiedData,
        sha256,
        panose: inferred?.panose,
        pitchFamily: inferred?.pitchFamily,
      });
    }),
  );
  return markSealedFont({
    ...parsedFont.data,
    faces: sealedFaces,
  });
}

/**
 * Unpacks a TrueType/OpenType Collection (`.ttc` / `.otc`) buffer by extracting all static
 * faces matching `family` (or the sole family when `options.substitute` is `true`) and returns
 * a sealed `Font` value with isolated face byte buffers.
 */
export function fromFontCollection(
  bytes: Uint8Array,
  family: string,
  options: FontOptions = {},
): Font {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError("Font collection bytes must be a Uint8Array");
  }
  if (typeof family !== "string" || family.trim() === "") {
    throw new TypeError("Font family must be a non-empty string");
  }
  const parsedOptions = fontOptionsSchema.safeParse(options);
  if (!parsedOptions.success) {
    throw new TypeError("Font options.substitute must be a boolean");
  }
  const validOptions = parsedOptions.data;
  const unpacked = unpackFontCollectionFaces(bytes, family, {
    substitute: validOptions.substitute,
  });
  const sealedFaces = Object.freeze(
    unpacked.map((face) => {
      const ownedBytes = markImmutableBytes(face.bytes);
      const sha256 = bytesSha256(ownedBytes);
      return createSealedFontFace({
        weight: face.weight,
        style: face.style,
        bytes: ownedBytes,
        sha256,
        panose: face.panose,
        pitchFamily: face.pitchFamily,
      });
    }),
  );
  return markSealedFont({
    family,
    faces: sealedFaces,
    ...(validOptions.substitute === true ? { substitute: true } : {}),
    ...(validOptions.embeddable === false ? { embeddable: false } : {}),
  });
}
