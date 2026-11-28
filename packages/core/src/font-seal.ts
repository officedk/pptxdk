import { createHash } from "node:crypto";

import type { Font, FontFace, FontFaceStyle } from "#/theme.ts";

/** PANOSE and PitchAndFamily classification extracted from an SFNT `OS/2` + `post` table. */
export type FontClassification = {
  readonly panose?: string | undefined;
  readonly pitchFamily: number;
};

/** Internal read-only view of a sealed or unsealed font face's bytes and cached SHA-256 digest. */
export type FontFaceResourceView = {
  readonly bytes: Uint8Array;
  readonly sha256: string;
  readonly panose?: string | undefined;
  readonly pitchFamily?: number | undefined;
};

const immutableBytes = new WeakSet<Uint8Array>();
const sha256ByBytes = new WeakMap<Uint8Array, string>();
const classificationByBytes = new WeakMap<
  Uint8Array,
  FontClassification | null
>();
const sealedFonts = new WeakSet<object>();
const sealedFaceResources = new WeakMap<FontFace, FontFaceResourceView>();

const sansSerifPanoseStyles = new Set([11, 12, 13, 14, 15]);

/** Computes DrawingML `pitchFamily` from 10-byte PANOSE and `post.isFixedPitch`. */
export function computePitchFamilyFromPanose(
  panose: readonly number[] | Uint8Array,
  isFixedPitch: number,
): number {
  const familyType = panose[0] ?? 0;
  const serifStyle = panose[1] ?? 0;
  const proportion = panose[3] ?? 0;
  const fixedPitch = isFixedPitch !== 0 || proportion === 9;
  const family = fixedPitch
    ? 0x30
    : familyType === 2
      ? sansSerifPanoseStyles.has(serifStyle)
        ? 0x20
        : serifStyle >= 2
          ? 0x10
          : 0
      : familyType === 3
        ? 0x40
        : familyType === 4 || familyType === 5
          ? 0x50
          : 0;
  return family | (fixedPitch ? 1 : 2);
}

/** Formats 10-byte `OS/2.panose` as a 20-character uppercase hexadecimal string. */
export function formatPanoseHex(
  panose: readonly number[] | Uint8Array,
): string {
  let out = "";
  for (let index = 0; index < 10; index += 1) {
    out += ((panose[index] ?? 0) & 0xff)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  }
  return out;
}

/**
 * Extracts `OS/2` `panose` and `pitchFamily` from raw SFNT (`.ttf`/`.otf`) bytes using pure
 * `DataView` reads so `@pptxdk/core` and `@pptxdk/emit` stay free of `fontkit`.
 */
function parseSfntClassificationUncached(
  bytes: Uint8Array,
): FontClassification | undefined {
  if (bytes.byteLength < 12) return undefined;
  const signature = String.fromCharCode(
    bytes[0]!,
    bytes[1]!,
    bytes[2]!,
    bytes[3]!,
  );
  if (
    signature !== "\0\x01\0\0" &&
    signature !== "true" &&
    signature !== "OTTO"
  ) {
    return undefined;
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numTables = view.getUint16(4);
  if (12 + numTables * 16 > bytes.byteLength) return undefined;
  let os2Offset: number | undefined;
  let os2Length = 0;
  let postOffset: number | undefined;
  let postLength = 0;
  for (let index = 0; index < numTables; index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(
      bytes[record]!,
      bytes[record + 1]!,
      bytes[record + 2]!,
      bytes[record + 3]!,
    );
    const offset = view.getUint32(record + 8);
    const length = view.getUint32(record + 12);
    if (offset > bytes.byteLength || length > bytes.byteLength - offset) {
      return undefined;
    }
    if (tag === "OS/2") {
      os2Offset = offset;
      os2Length = length;
    } else if (tag === "post") {
      postOffset = offset;
      postLength = length;
    }
  }
  if (os2Offset === undefined || os2Length < 42) {
    return undefined;
  }
  const panoseBytes = bytes.subarray(os2Offset + 32, os2Offset + 42);
  const isFixedPitch =
    postOffset !== undefined && postLength >= 16
      ? view.getUint32(postOffset + 12)
      : 0;
  return {
    panose: formatPanoseHex(panoseBytes),
    pitchFamily: computePitchFamilyFromPanose(panoseBytes, isFixedPitch),
  };
}

function getOrParseSfntClassification(
  bytes: Uint8Array,
): FontClassification | undefined {
  if (immutableBytes.has(bytes)) {
    const cached = classificationByBytes.get(bytes);
    if (cached !== undefined) {
      return cached ?? undefined;
    }
    const parsed = parseSfntClassificationUncached(bytes);
    classificationByBytes.set(bytes, parsed ?? null);
    return parsed;
  }
  return parseSfntClassificationUncached(bytes);
}

/**
 * Marks an internally owned `Uint8Array` as immutable so `bytesSha256` can memoize its digest
 * across pipeline stages and `snapshotValue` / `readImage` can reuse it zero-copy.
 */
export function markImmutableBytes(bytes: Uint8Array): Uint8Array {
  immutableBytes.add(bytes);
  return bytes;
}

/** Returns `true` when `bytes` is an internal immutable buffer registered with `markImmutableBytes`. */
export function isImmutableBytes(bytes: Uint8Array): boolean {
  return immutableBytes.has(bytes);
}

/**
 * Computes the lowercase hex SHA-256 digest for a `Uint8Array`.
 *
 * Caching in `sha256ByBytes` is restricted to buffers registered via `markImmutableBytes`;
 * unmarked caller-owned buffers are hashed fresh every time to prevent cache poisoning.
 */
export function bytesSha256(bytes: Uint8Array): string {
  if (!immutableBytes.has(bytes)) {
    return createHash("sha256").update(bytes).digest("hex");
  }
  const cached = sha256ByBytes.get(bytes);
  if (cached !== undefined) {
    return cached;
  }
  const digest = createHash("sha256").update(bytes).digest("hex");
  sha256ByBytes.set(bytes, digest);
  return digest;
}

/**
 * Creates a frozen `FontFace` whose private bytes live in a module-private `WeakMap` and whose
 * enumerable `.data` getter returns a fresh defensive `Uint8Array` copy on public reads.
 */
export function createSealedFontFace(input: {
  readonly weight?: number | undefined;
  readonly style?: FontFaceStyle | undefined;
  readonly bytes?: Uint8Array | undefined;
  readonly sha256?: string | undefined;
  readonly panose?: string | undefined;
  readonly pitchFamily?: number | undefined;
}): FontFace {
  const faceRecord: Record<string, unknown> = {
    ...(input.weight === undefined ? {} : { weight: input.weight }),
    ...(input.style === undefined ? {} : { style: input.style }),
  };

  if (input.bytes !== undefined) {
    const privateBytes = markImmutableBytes(input.bytes);
    const sha256 = input.sha256 ?? bytesSha256(privateBytes);
    sha256ByBytes.set(privateBytes, sha256);
    Object.defineProperty(faceRecord, "data", {
      enumerable: true,
      configurable: false,
      get(): Uint8Array {
        return new Uint8Array(privateBytes);
      },
    });
    const frozenFace = Object.freeze(faceRecord) as FontFace;
    sealedFaceResources.set(frozenFace, {
      bytes: privateBytes,
      sha256,
      ...(input.panose === undefined ? {} : { panose: input.panose }),
      ...(input.pitchFamily === undefined
        ? {}
        : { pitchFamily: input.pitchFamily }),
    });
    return frozenFace;
  }

  return Object.freeze(faceRecord) as FontFace;
}

/**
 * Marks a frozen `Font` object as sealed so input snapshotting and Zod validation short-circuit
 * in $O(1)$ without adding extra own properties or symbols to the object.
 */
export function markSealedFont(fontValue: Font): Font {
  Object.freeze(fontValue);
  sealedFonts.add(fontValue);
  return fontValue;
}

/** Returns `true` when `value` is a sealed `Font` created by `font()`. */
export function isSealedFont(value: unknown): value is Font {
  return typeof value === "object" && value !== null && sealedFonts.has(value);
}

/**
 * Checks whether a `FontFace` carries binary bytes without triggering the defensive `.data`
 * getter on sealed faces.
 */
export function hasFontFaceBytes(face: FontFace): boolean {
  if (sealedFaceResources.has(face)) {
    return true;
  }
  return face.data !== undefined;
}

/**
 * Returns the internal bytes and cached SHA-256 digest for a sealed or unsealed `FontFace`
 * without invoking the public defensive `.data` copy getter on sealed faces.
 */
export function peekFontFaceResource(
  face: FontFace,
): FontFaceResourceView | undefined {
  const sealed = sealedFaceResources.get(face);
  if (sealed !== undefined) {
    return sealed;
  }
  const rawBytes = face.data;
  if (rawBytes === undefined) {
    return undefined;
  }
  return {
    bytes: rawBytes,
    sha256: bytesSha256(rawBytes),
  };
}

/**
 * Returns cached or parsed `OS/2` `panose` and `pitchFamily` metadata for a sealed or unsealed
 * `FontFace` without triggering defensive `.data` copies on sealed faces.
 */
export function peekFontFaceClassification(
  face: FontFace,
): FontClassification | undefined {
  const sealed = sealedFaceResources.get(face);
  if (sealed !== undefined) {
    if (sealed.pitchFamily !== undefined) {
      return {
        ...(sealed.panose === undefined ? {} : { panose: sealed.panose }),
        pitchFamily: sealed.pitchFamily,
      };
    }
    return getOrParseSfntClassification(sealed.bytes);
  }
  const rawBytes = face.data;
  if (rawBytes === undefined) return undefined;
  return getOrParseSfntClassification(rawBytes);
}

/**
 * Resolves `OS/2` classification metadata (`panose` and `pitchFamily`) for a `Font` value,
 * preferring the face matching `(weight, style)`, then `(400, "normal")`, then the first face
 * with classification metadata.
 */
export function resolveFontClassification(
  fontValue: Font,
  weight?: number,
  faceStyle?: FontFaceStyle,
): FontClassification | undefined {
  const faces = fontValue.faces;
  if (faces === undefined || faces.length === 0) return undefined;
  const targetWeight = weight ?? 400;
  const targetStyle = faceStyle ?? "normal";
  const exact = faces.find(
    (candidate) =>
      (candidate.weight ?? 400) === targetWeight &&
      (candidate.style ?? "normal") === targetStyle &&
      peekFontFaceClassification(candidate) !== undefined,
  );
  if (exact !== undefined) {
    return peekFontFaceClassification(exact);
  }
  const regular =
    targetWeight === 400 && targetStyle === "normal"
      ? undefined
      : faces.find(
          (candidate) =>
            (candidate.weight ?? 400) === 400 &&
            (candidate.style ?? "normal") === "normal" &&
            peekFontFaceClassification(candidate) !== undefined,
        );
  if (regular !== undefined) {
    return peekFontFaceClassification(regular);
  }
  for (const candidate of faces) {
    const cls = peekFontFaceClassification(candidate);
    if (cls !== undefined) return cls;
  }
  return undefined;
}
