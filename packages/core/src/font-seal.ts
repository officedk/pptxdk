import { createHash } from "node:crypto";

import type { Font, FontFace, FontFaceStyle } from "#/theme.ts";

/** Internal read-only view of a sealed or unsealed font face's bytes and cached SHA-256 digest. */
export type FontFaceResourceView = {
  readonly bytes: Uint8Array;
  readonly sha256: string;
};

const immutableBytes = new WeakSet<Uint8Array>();
const sha256ByBytes = new WeakMap<Uint8Array, string>();
const sealedFonts = new WeakSet<object>();
const sealedFaceResources = new WeakMap<FontFace, FontFaceResourceView>();

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
    sealedFaceResources.set(frozenFace, { bytes: privateBytes, sha256 });
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
