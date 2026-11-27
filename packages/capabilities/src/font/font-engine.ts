import {
  bytesSha256,
  CapabilityError,
  peekFontFaceResource,
  type FontEmbeddingMode,
  type FontFaceStyle,
} from "@pptxdk/core/capability";
import * as fontkit from "fontkit";
import * as hb from "harfbuzzjs";

import type {} from "./emscripten.d.ts";
import { subsetFont } from "./font-subset.ts";

/** Normalizes a font family name for case-insensitive, whitespace-collapsed lookup. */
export function fontFamilyKey(value: string): string {
  return value.trim().replace(/\s+/gu, " ").toLowerCase();
}

type FontBinaryFormat = "TTF" | "CFF" | "WOFF" | "WOFF2";

type SfntTableRecord = {
  readonly offset: number;
  readonly length: number;
};

type SfntDirectory = {
  readonly signature: string;
  readonly tables: ReadonlyMap<string, SfntTableRecord>;
};

type FontkitExtras = {
  readonly post?: { readonly isFixedPitch?: number };
  readonly directory?: { readonly tables?: Record<string, unknown> };
  readonly name?: {
    readonly records?: {
      readonly fontFamily?: Readonly<Record<string, unknown>>;
      readonly preferredFamily?: Readonly<Record<string, unknown>>;
    };
  };
};

type UnicodeFontNames = {
  readonly legacyFamilies: ReadonlySet<string>;
  readonly legacyFamilyKeys: ReadonlySet<string>;
  readonly preferredFamilies: ReadonlySet<string>;
  readonly preferredFamilyKeys: ReadonlySet<string>;
  readonly primaryLegacyFamily: string | undefined;
  readonly eotUtf16LeFields: ReadonlyMap<number, Uint8Array>;
  readonly nameTableError: string | undefined;
};

type CachedHarfBuzzFace = {
  readonly blob: hb.Blob;
  readonly face: hb.Face;
  readonly font: hb.Font;
};

type ParsedFontBinary = {
  readonly identity: string;
  readonly bytes: Uint8Array;
  readonly format: FontBinaryFormat;
  readonly font: fontkit.Font;
  readonly sfnt: SfntDirectory | undefined;
  readonly names: UnicodeFontNames;
  readonly hasGlyf: boolean;
  readonly hasFvar: boolean;
  readonly hasColorOrBitmapTable: boolean;
  readonly hasOs2: boolean;
  readonly weightClass: number | undefined;
  readonly style: FontFaceStyle;
  readonly os2Bold: boolean;
  readonly isItalic: boolean;
  readonly fsType: number | undefined;
  readonly unitsPerEm: number;
  readonly ascent: number;
  readonly descent: number;
  readonly lineGap: number;
  readonly underlineThickness: number;
  readonly underlinePosition: number;
  readonly yStrikeoutSize: number;
  readonly yStrikeoutPosition: number;
  readonly pitchFamily: number;
  readonly charset: number;
  hbEntry: CachedHarfBuzzFace | undefined;
};

/** Publicly exposed measurement & outline handle returned by the font deep module. */
export type FontMeasureFace = {
  readonly identity: string;
  readonly unitsPerEm: number;
  readonly ascent: number;
  readonly descent: number;
  readonly lineGap: number;
  readonly underlineThickness: number;
  readonly underlinePosition: number;
  readonly yStrikeoutSize: number;
  readonly yStrikeoutPosition: number;
  readonly os2Bold: boolean;
  readonly hasColorOrBitmapTable: boolean;
};

export type FontResourceDescriptor = {
  readonly family: string;
  readonly weight: number;
  readonly style: FontFaceStyle;
  readonly substitute: boolean;
  readonly data: Uint8Array;
  readonly identity: string;
};

export type ShapedOutlineGlyph = {
  readonly path: string;
  readonly xAdvance: number;
  readonly xOffset: number;
  readonly yOffset: number;
  readonly cluster: number;
  readonly isSpace: boolean;
};

export type CheckedEmbeddableTrueType = {
  readonly identity: string;
  readonly primaryLegacyFamily: string | undefined;
  readonly weightClass: number;
  readonly fsType: number;
  readonly os2Bold: boolean;
  readonly isItalic: boolean;
  readonly pitchFamily: number;
  readonly charset: number;
};

export type FontEmbeddingContext = {
  readonly family: string;
  readonly weight: number;
  readonly style: FontFaceStyle;
  readonly path?: string;
};

export type ImportedEmbeddedFontInfo = {
  readonly familyMatches: boolean;
  readonly weight: number | undefined;
  readonly style: FontFaceStyle | undefined;
};

const unsupportedColorFontTables = [
  "CBDT",
  "CBLC",
  "COLR",
  "CPAL",
  "SVG ",
  "sbix",
] as const;

const parsedFontCache = new Map<string, ParsedFontBinary>();
const measureFaceCache = new WeakMap<FontMeasureFace, ParsedFontBinary>();

let sharedHbBuffer: hb.Buffer | undefined;
let sharedHbBufferBusy = false;

function withHbBuffer<T>(fn: (buffer: hb.Buffer) => T): T {
  if (!sharedHbBufferBusy) {
    sharedHbBufferBusy = true;
    sharedHbBuffer ??= new hb.Buffer();
    try {
      return fn(sharedHbBuffer);
    } finally {
      sharedHbBuffer.reset();
      sharedHbBufferBusy = false;
    }
  }
  const buffer = new hb.Buffer();
  try {
    return fn(buffer);
  } finally {
    buffer.reset();
  }
}

function readSfntDirectory(bytes: Uint8Array): SfntDirectory {
  if (bytes.byteLength < 12) {
    throw new Error("Font data is shorter than the 12-byte SFNT header");
  }
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  if (
    signature !== "\0\x01\0\0" &&
    signature !== "true" &&
    signature !== "OTTO"
  ) {
    throw new Error("Font data is not a valid SFNT file");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const count = view.getUint16(4);
  if (12 + count * 16 > bytes.byteLength) {
    throw new Error("Font SFNT table directory is incomplete");
  }
  const tables = new Map<string, SfntTableRecord>();
  for (let index = 0; index < count; index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...bytes.subarray(record, record + 4));
    const offset = view.getUint32(record + 8);
    const length = view.getUint32(record + 12);
    if (offset > bytes.byteLength || length > bytes.byteLength - offset) {
      throw new Error(`Font table ${tag} exceeds file boundaries`);
    }
    tables.set(tag, { offset, length });
  }
  return { signature, tables };
}

function tryReadSfntDirectory(bytes: Uint8Array): SfntDirectory | undefined {
  if (bytes.byteLength < 12) return undefined;
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  if (
    signature !== "\0\x01\0\0" &&
    signature !== "true" &&
    signature !== "OTTO"
  ) {
    return undefined;
  }
  return readSfntDirectory(bytes);
}

function isItalicFontkit(font: fontkit.Font): boolean {
  const os2 = font["OS/2"];
  return (
    Boolean(os2?.fsSelection?.italic) ||
    Boolean(os2?.fsSelection?.oblique) ||
    font.italicAngle !== 0
  );
}

const utf16BeDecoder = new TextDecoder("utf-16be");
const eotRequiredNameIds = new Set([1, 2, 4, 5]);

function unicodeNamePriority(platformID: number, languageID: number): number {
  if (platformID === 3 && languageID === 0x0409) return 0;
  if (platformID === 0) return 1;
  return 2;
}

function emptyUnicodeFontNames(nameTableError?: string): UnicodeFontNames {
  return {
    legacyFamilies: new Set(),
    legacyFamilyKeys: new Set(),
    preferredFamilies: new Set(),
    preferredFamilyKeys: new Set(),
    primaryLegacyFamily: undefined,
    eotUtf16LeFields: new Map(),
    nameTableError,
  };
}

/**
 * Extracts all Unicode `Name ID 1` / `Name ID 16` family names and EOT UTF-16LE fields
 * from an SFNT `name` table in a single pass (`platformID === 0` or `platformID === 3`
 * with `encodingID === 1 || 10`, across all language IDs).
 */
function parseSfntUnicodeNames(
  bytes: Uint8Array,
  table: SfntTableRecord | undefined,
): UnicodeFontNames {
  if (table === undefined) {
    return emptyUnicodeFontNames("Font name table is missing");
  }
  if (table.length < 6) {
    return emptyUnicodeFontNames("Font name table is incomplete");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const count = view.getUint16(table.offset + 2);
  const stringsOffset = view.getUint16(table.offset + 4);
  if (6 + count * 12 > table.length || stringsOffset > table.length) {
    return emptyUnicodeFontNames("Font name table record out of bounds");
  }

  const legacyFamilies = new Set<string>();
  const legacyFamilyKeys = new Set<string>();
  const preferredFamilies = new Set<string>();
  const preferredFamilyKeys = new Set<string>();
  let primaryLegacyFamily: string | undefined;
  let primaryLegacyPriority = Number.POSITIVE_INFINITY;
  const eotUtf16LeFields = new Map<number, Uint8Array>();
  const eotPriorities = new Map<number, number>();
  let nameTableError: string | undefined;

  for (let index = 0; index < count; index += 1) {
    const record = table.offset + 6 + index * 12;
    const platformID = view.getUint16(record);
    const encodingID = view.getUint16(record + 2);
    const languageID = view.getUint16(record + 4);
    const nameID = view.getUint16(record + 6);
    const length = view.getUint16(record + 8);
    const offset = stringsOffset + view.getUint16(record + 10);
    if (offset > table.length || length > table.length - offset) {
      nameTableError ??= "Font name string out of bounds";
      continue;
    }
    const isUnicode =
      platformID === 0 ||
      (platformID === 3 && (encodingID === 1 || encodingID === 10));
    if (!isUnicode) continue;

    const raw = bytes.subarray(
      table.offset + offset,
      table.offset + offset + length,
    );
    if (raw.byteLength % 2 !== 0) {
      nameTableError ??= "Font name string is not UTF-16 encoded";
      continue;
    }

    const priority = unicodeNamePriority(platformID, languageID);
    if (eotRequiredNameIds.has(nameID)) {
      const existingPriority =
        eotPriorities.get(nameID) ?? Number.POSITIVE_INFINITY;
      if (priority < existingPriority) {
        const littleEndian = new Uint8Array(raw.byteLength);
        for (let at = 0; at < raw.byteLength; at += 2) {
          littleEndian[at] = raw[at + 1]!;
          littleEndian[at + 1] = raw[at]!;
        }
        eotUtf16LeFields.set(nameID, littleEndian);
        eotPriorities.set(nameID, priority);
      }
    }

    if (nameID !== 1 && nameID !== 16) continue;
    const decoded = utf16BeDecoder.decode(raw);
    if (decoded.trim() === "") continue;
    if (nameID === 1) {
      legacyFamilies.add(decoded);
      legacyFamilyKeys.add(fontFamilyKey(decoded));
      if (priority < primaryLegacyPriority) {
        primaryLegacyFamily = decoded;
        primaryLegacyPriority = priority;
      }
    } else {
      preferredFamilies.add(decoded);
      preferredFamilyKeys.add(fontFamilyKey(decoded));
    }
  }

  return {
    legacyFamilies,
    legacyFamilyKeys,
    preferredFamilies,
    preferredFamilyKeys,
    primaryLegacyFamily,
    eotUtf16LeFields,
    nameTableError,
  };
}

/**
 * Fallback Unicode family extraction for WOFF/WOFF2 containers via `fontkit`'s
 * decompressed `name.records` across all language entries.
 */
function extractWoffUnicodeNames(font: fontkit.Font): UnicodeFontNames {
  const records = (font as unknown as FontkitExtras).name?.records;
  const legacyFamilies = new Set<string>();
  const legacyFamilyKeys = new Set<string>();
  const preferredFamilies = new Set<string>();
  const preferredFamilyKeys = new Set<string>();

  const fontFamilyRecords = records?.fontFamily ?? {};
  for (const value of Object.values(fontFamilyRecords)) {
    if (typeof value === "string" && value.trim() !== "") {
      legacyFamilies.add(value);
      legacyFamilyKeys.add(fontFamilyKey(value));
    }
  }

  const preferredFamilyRecords = records?.preferredFamily ?? {};
  for (const value of Object.values(preferredFamilyRecords)) {
    if (typeof value === "string" && value.trim() !== "") {
      preferredFamilies.add(value);
      preferredFamilyKeys.add(fontFamilyKey(value));
    }
  }

  const englishLegacy = fontFamilyRecords["en"];
  const firstLegacy = [...legacyFamilies][0];
  const fallbackFamily =
    typeof font.familyName === "string" && font.familyName.trim() !== ""
      ? font.familyName
      : undefined;
  const primaryLegacyFamily =
    typeof englishLegacy === "string" && englishLegacy.trim() !== ""
      ? englishLegacy
      : (firstLegacy ?? fallbackFamily);
  if (primaryLegacyFamily !== undefined && legacyFamilies.size === 0) {
    legacyFamilies.add(primaryLegacyFamily);
    legacyFamilyKeys.add(fontFamilyKey(primaryLegacyFamily));
  }

  return {
    legacyFamilies,
    legacyFamilyKeys,
    preferredFamilies,
    preferredFamilyKeys,
    primaryLegacyFamily,
    eotUtf16LeFields: new Map(),
    nameTableError: undefined,
  };
}

function matchesAnyUnicodeFamily(
  parsed: ParsedFontBinary,
  family: string,
): boolean {
  const targetKey = fontFamilyKey(family);
  return (
    parsed.names.legacyFamilyKeys.has(targetKey) ||
    parsed.names.preferredFamilyKeys.has(targetKey)
  );
}

function matchesLegacyUnicodeFamily(
  parsed: ParsedFontBinary,
  family: string,
): boolean {
  return parsed.names.legacyFamilyKeys.has(fontFamilyKey(family));
}

function computeFontCategory(font: fontkit.Font): {
  readonly pitchFamily: number;
  readonly charset: number;
} {
  const panose = font["OS/2"]?.panose ?? [];
  const serifStyle = panose[1] ?? 0;
  const proportion = panose[3] ?? 0;
  const fixedPitch =
    ((font as unknown as FontkitExtras).post?.isFixedPitch ?? 0) !== 0 ||
    proportion === 9;
  const sansSerifStyles = new Set([11, 12, 13, 14, 15]);
  const family = fixedPitch
    ? 0x30
    : panose[0] === 2
      ? sansSerifStyles.has(serifStyle)
        ? 0x20
        : serifStyle >= 2
          ? 0x10
          : 0
      : panose[0] === 3
        ? 0x40
        : panose[0] === 4 || panose[0] === 5
          ? 0x50
          : 0;
  return { pitchFamily: family | (fixedPitch ? 1 : 2), charset: 0 };
}

function isValidWeightClass(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 1000
  );
}

function parseFontBinaryUncached(
  bytes: Uint8Array,
  identity: string,
): ParsedFontBinary {
  if (bytes.byteLength < 12) {
    throw new Error("Font data is empty or truncated");
  }
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  if (signature === "ttcf") {
    throw new Error("Font collections (TTC/OTC) are not supported");
  }
  const safeBytes = new Uint8Array(bytes);
  const sfnt = tryReadSfntDirectory(safeBytes);
  const opened = fontkit.create(
    Buffer.from(safeBytes.buffer, safeBytes.byteOffset, safeBytes.byteLength),
  );
  if (
    !("unitsPerEm" in opened) ||
    typeof opened.hasGlyphForCodePoint !== "function"
  ) {
    throw new Error("Font collection is not a single font file");
  }

  let format: FontBinaryFormat;
  if (opened.type === "WOFF") {
    format = "WOFF";
  } else if (opened.type === "WOFF2") {
    format = "WOFF2";
  } else if (opened.type === "TTF") {
    if (sfnt === undefined) {
      throw new Error("SFNT header is invalid for uncompressed font");
    }
    const isCff =
      sfnt.signature === "OTTO" ||
      sfnt.tables.has("CFF ") ||
      sfnt.tables.has("CFF2");
    format = isCff ? "CFF" : "TTF";
  } else {
    throw new Error(
      `Unsupported font container type: ${String((opened as { readonly type?: unknown }).type)}`,
    );
  }

  const directoryTables = (opened as unknown as FontkitExtras).directory
    ?.tables;
  const hasTable = (tag: string): boolean =>
    Boolean(sfnt?.tables.has(tag)) ||
    (directoryTables !== undefined && tag in directoryTables) ||
    tag in opened;

  const hasFvar =
    hasTable("fvar") ||
    (opened.variationAxes !== undefined &&
      Object.keys(opened.variationAxes).length > 0);
  const hasColorOrBitmapTable = unsupportedColorFontTables.some((tag) =>
    hasTable(tag),
  );
  const hasGlyf =
    format === "TTF" &&
    Boolean(sfnt?.tables.has("glyf")) &&
    !hasTable("CFF ") &&
    !hasTable("CFF2");

  const os2 = opened["OS/2"];
  const hasOs2 = os2 !== undefined;
  let weightClass: number | undefined;
  let fsType: number | undefined;
  const rawOs2 = sfnt?.tables.get("OS/2");
  if (rawOs2 !== undefined && rawOs2.length >= 10) {
    const view = new DataView(
      safeBytes.buffer,
      safeBytes.byteOffset,
      safeBytes.byteLength,
    );
    const rawWeight = view.getUint16(rawOs2.offset + 4);
    weightClass = isValidWeightClass(rawWeight) ? rawWeight : undefined;
    fsType = view.getUint16(rawOs2.offset + 8);
  } else if (hasOs2 && isValidWeightClass(os2.usWeightClass)) {
    weightClass = os2.usWeightClass;
  }

  const isItalic = isItalicFontkit(opened);
  const os2Bold = Boolean(os2?.fsSelection?.bold);
  const { pitchFamily, charset } = computeFontCategory(opened);
  const names =
    sfnt !== undefined
      ? parseSfntUnicodeNames(safeBytes, sfnt.tables.get("name"))
      : extractWoffUnicodeNames(opened);

  return {
    identity,
    bytes: safeBytes,
    format,
    font: opened,
    sfnt,
    names,
    hasGlyf,
    hasFvar,
    hasColorOrBitmapTable,
    hasOs2,
    weightClass,
    style: isItalic ? "italic" : "normal",
    os2Bold,
    isItalic,
    fsType,
    unitsPerEm: opened.unitsPerEm,
    ascent: opened.ascent,
    descent: opened.descent,
    lineGap: opened.lineGap,
    underlineThickness: opened.underlineThickness,
    underlinePosition: opened.underlinePosition,
    yStrikeoutSize: os2?.yStrikeoutSize ?? 0,
    yStrikeoutPosition: os2?.yStrikeoutPosition ?? 0,
    pitchFamily,
    charset,
    hbEntry: undefined,
  };
}

function getOrParseFontBinary(
  bytes: Uint8Array,
  knownIdentity?: string,
): ParsedFontBinary {
  const identity = knownIdentity ?? bytesSha256(bytes);
  let parsed = parsedFontCache.get(identity);
  if (parsed === undefined) {
    parsed = parseFontBinaryUncached(bytes, identity);
    parsedFontCache.set(identity, parsed);
  }
  return parsed;
}

function tryGetOrParseFontBinary(
  bytes: Uint8Array,
  knownIdentity?: string,
): ParsedFontBinary | undefined {
  try {
    return getOrParseFontBinary(bytes, knownIdentity);
  } catch {
    return undefined;
  }
}

function ensureHarfBuzzFont(parsed: ParsedFontBinary): hb.Font | undefined {
  if (parsed.format !== "TTF" && parsed.format !== "CFF") {
    return undefined;
  }
  if (parsed.hbEntry === undefined) {
    const blob = new hb.Blob(parsed.bytes);
    const face = new hb.Face(blob, 0);
    const font = new hb.Font(face);
    font.setScale(parsed.unitsPerEm, parsed.unitsPerEm);
    parsed.hbEntry = { blob, face, font };
  }
  return parsed.hbEntry.font;
}

function toMeasureFace(parsed: ParsedFontBinary): FontMeasureFace {
  const face: FontMeasureFace = {
    identity: parsed.identity,
    unitsPerEm: parsed.unitsPerEm,
    ascent: parsed.ascent,
    descent: parsed.descent,
    lineGap: parsed.lineGap,
    underlineThickness: parsed.underlineThickness,
    underlinePosition: parsed.underlinePosition,
    yStrikeoutSize: parsed.yStrikeoutSize,
    yStrikeoutPosition: parsed.yStrikeoutPosition,
    os2Bold: parsed.os2Bold,
    hasColorOrBitmapTable: parsed.hasColorOrBitmapTable,
  };
  measureFaceCache.set(face, parsed);
  return face;
}

function parsedOfMeasureFace(face: FontMeasureFace): ParsedFontBinary {
  const cached =
    measureFaceCache.get(face) ?? parsedFontCache.get(face.identity);
  if (cached === undefined) {
    throw new Error(`Missing parsed font binary for ${face.identity}`);
  }
  return cached;
}

function fontErrorDetails(
  source: FontEmbeddingContext,
  extra?: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  return {
    font: {
      family: source.family,
      weight: source.weight,
      style: source.style,
    },
    ...extra,
  };
}

function failFont(
  source: FontEmbeddingContext,
  code: string,
  message: string,
  extra?: Readonly<Record<string, unknown>>,
): never {
  throw new CapabilityError(
    code,
    message,
    source.path,
    fontErrorDetails(source, extra),
  );
}

/**
 * Attempts to read static single-font weight (`OS/2.usWeightClass`) and glyph style
 * (`italic || oblique || italicAngle !== 0`) for `font()` construction.
 * Returns `undefined` for corrupt bytes, font collections, or variable (`fvar`) fonts
 * so `resolve`/`emit` can report a node-path-attributed `BuildError`.
 */
export function tryInferStaticFontFaceMetadata(
  bytes: Uint8Array,
  knownIdentity?: string,
):
  | {
      readonly weight: number;
      readonly style: FontFaceStyle;
    }
  | undefined {
  const parsed = tryGetOrParseFontBinary(bytes, knownIdentity);
  if (
    parsed === undefined ||
    parsed.hasFvar ||
    parsed.weightClass === undefined
  ) {
    return undefined;
  }
  return {
    weight: parsed.weightClass,
    style: parsed.style,
  };
}

/**
 * Resolves the effective `(weight, style)` slot of a declared `FontFace` for style matching.
 * Supports unconstructed literal `{ data }` faces and ensures `fvar`/corrupt omitted-weight
 * faces are still selected so `resolve` rejects them with full path context.
 */
export function resolveDeclaredFaceSlot(face: {
  readonly weight?: number | undefined;
  readonly style?: FontFaceStyle | undefined;
  readonly data?: Uint8Array | undefined;
}): {
  readonly weight: number;
  readonly style: FontFaceStyle;
  readonly matchAnyWeightOnFvarOrCorrupt: boolean;
} {
  if (face.weight !== undefined && face.style !== undefined) {
    return {
      weight: face.weight,
      style: face.style,
      matchAnyWeightOnFvarOrCorrupt: false,
    };
  }
  const resource = peekFontFaceResource(face);
  if (resource === undefined) {
    return {
      weight: face.weight ?? 400,
      style: face.style ?? "normal",
      matchAnyWeightOnFvarOrCorrupt: false,
    };
  }
  const parsed = tryGetOrParseFontBinary(resource.bytes, resource.sha256);
  if (parsed === undefined || parsed.hasFvar) {
    return {
      weight: face.weight ?? parsed?.weightClass ?? 400,
      style: face.style ?? parsed?.style ?? "normal",
      matchAnyWeightOnFvarOrCorrupt: face.weight === undefined,
    };
  }
  return {
    weight: face.weight ?? parsed.weightClass ?? 400,
    style: face.style ?? parsed.style,
    matchAnyWeightOnFvarOrCorrupt: false,
  };
}

/**
 * Validates that a font face resource is a single non-variable font whose actual `OS/2`
 * weight and italic metadata match the declared `weight` and `style`.
 */
export function validateFontFaceResource(
  source: FontResourceDescriptor,
): FontMeasureFace {
  let parsed: ParsedFontBinary;
  try {
    parsed = getOrParseFontBinary(source.data, source.identity);
  } catch (error) {
    return failFont(
      source,
      "unsupported-font-format",
      `Cannot parse font file for ${source.family}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  if (parsed.hasFvar) {
    return failFont(
      source,
      "unsupported-font-format",
      `Variable font "${source.family}" contains an OpenType "fvar" axis table and is not supported; supply static per-weight font files instead`,
    );
  }
  if (parsed.weightClass === undefined) {
    return failFont(
      source,
      "unsupported-font-format",
      `Font file for ${source.family} is missing a valid OS/2 usWeightClass (1..1000)`,
    );
  }
  if (parsed.weightClass !== source.weight) {
    return failFont(
      source,
      "font-weight-mismatch",
      `Font value for ${source.family} declared weight ${source.weight}, but file OS/2 usWeightClass is ${parsed.weightClass}`,
    );
  }
  if (parsed.style !== source.style) {
    return failFont(
      source,
      "font-style-mismatch",
      `Font value for ${source.family} declared style "${source.style}", but file metadata indicates "${parsed.style}"`,
    );
  }
  return toMeasureFace(parsed);
}

/**
 * Opens a validated font face for text measurement and preview outline extraction,
 * verifying family match (unless `substitute` is true), line metrics, and initializing
 * HarfBuzz shaping for static `.ttf` and static CFF `.otf` fonts.
 */
export function openMeasureFontFace(
  source: FontResourceDescriptor,
): FontMeasureFace {
  const face = validateFontFaceResource(source);
  const parsed = parsedOfMeasureFace(face);
  if (!source.substitute && !matchesAnyUnicodeFamily(parsed, source.family)) {
    const declaredFamily =
      parsed.names.primaryLegacyFamily ??
      [...parsed.names.preferredFamilies][0] ??
      parsed.font.familyName;
    return failFont(
      source,
      "font-substitution-required",
      `Metrics file declares ${declaredFamily}; retaining output font name ${source.family} requires setting Font.substitute explicitly`,
      { reason: "font-substitution-required" },
    );
  }
  if (
    !Number.isFinite(parsed.unitsPerEm) ||
    parsed.unitsPerEm <= 0 ||
    !Number.isFinite(parsed.ascent) ||
    !Number.isFinite(parsed.descent) ||
    !Number.isFinite(parsed.lineGap)
  ) {
    return failFont(
      source,
      "font-metrics-invalid",
      `Font line metrics for ${source.family} are invalid`,
      { reason: "font-metrics-invalid" },
    );
  }
  ensureHarfBuzzFont(parsed);
  return face;
}

/** Returns unique visible characters in `text` that lack a non-.notdef glyph in `face`. */
export function findMissingFontGlyphs(
  face: FontMeasureFace,
  text: string,
  requiresVisibleGlyph: (character: string) => boolean,
): string[] {
  const parsed = parsedOfMeasureFace(face);
  const missing = new Set<string>();
  for (const character of text) {
    if (!requiresVisibleGlyph(character)) continue;
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      (!parsed.font.hasGlyphForCodePoint(codePoint) ||
        parsed.font.glyphForCodePoint(codePoint).id === 0)
    ) {
      missing.add(character);
    }
  }
  return [...missing];
}

/** Shapes `visibleText` and returns per-glyph horizontal advances in font design units. */
export function shapeFontAdvances(
  face: FontMeasureFace,
  visibleText: string,
): readonly number[] {
  if (visibleText === "") return [];
  const parsed = parsedOfMeasureFace(face);
  const hbFont = ensureHarfBuzzFont(parsed);
  if (hbFont !== undefined) {
    return withHbBuffer((buffer) => {
      buffer.addText(visibleText);
      buffer.guessSegmentProperties();
      hb.shape(hbFont, buffer);
      return buffer
        .getGlyphInfosAndPositions()
        .map((item) => item.xAdvance ?? 0);
    });
  }
  const shaped = parsed.font.layout(visibleText);
  return shaped.positions.map((position) => position.xAdvance);
}

/** Shapes `visibleText` with direction/language and returns SVG glyph outline paths and offsets. */
export function shapeFontOutlines(
  face: FontMeasureFace,
  visibleText: string,
  isRtl: boolean,
  lang: string | undefined,
): readonly ShapedOutlineGlyph[] {
  if (visibleText === "") return [];
  const parsed = parsedOfMeasureFace(face);
  const hbFont = ensureHarfBuzzFont(parsed);
  if (hbFont !== undefined) {
    return withHbBuffer((buffer) => {
      buffer.addText(visibleText);
      buffer.guessSegmentProperties();
      buffer.setDirection(isRtl ? hb.Direction.RTL : hb.Direction.LTR);
      if (lang !== undefined) {
        buffer.setLanguage(lang);
      }
      hb.shape(hbFont, buffer);
      return buffer.getGlyphInfosAndPositions().map((item) => ({
        path: hbFont.glyphToPath(item.codepoint),
        xAdvance: item.xAdvance ?? 0,
        xOffset: item.xOffset ?? 0,
        yOffset: item.yOffset ?? 0,
        cluster: item.cluster,
        isSpace: /\p{White_Space}/u.test(visibleText[item.cluster] ?? ""),
      }));
    });
  }
  const fkRun = parsed.font.layout(
    visibleText,
    undefined,
    undefined,
    undefined,
    isRtl ? "rtl" : "ltr",
  );
  return fkRun.glyphs.map((glyph, index) => {
    const pos = fkRun.positions[index]!;
    const isSpace = glyph.codePoints.some((cp) =>
      /\p{White_Space}/u.test(String.fromCodePoint(cp)),
    );
    return {
      path: glyph.path.toSVG(),
      xAdvance: pos.xAdvance,
      xOffset: pos.xOffset,
      yOffset: pos.yOffset,
      cluster: index,
      isSpace,
    };
  });
}

function checkFsTypeEmbeddability(fsType: number): boolean {
  const usagePermission = fsType & 0x000f;
  if (![0, 2, 4, 8].includes(usagePermission) || (fsType & ~0x030f) !== 0) {
    return false;
  }
  if (
    usagePermission === 2 ||
    usagePermission === 4 ||
    (fsType & 0x0200) !== 0
  ) {
    return false;
  }
  return true;
}

/**
 * Checks whether raw font bytes, declared `family` (`Name ID 1`), and author policy allow full
 * TrueType font embedding.
 */
export function isTrueTypeFontEmbeddable(
  bytes: Uint8Array,
  family: string,
  substitute: boolean,
  declaredEmbeddable: boolean,
  knownIdentity?: string,
): boolean {
  if (substitute || !declaredEmbeddable) return false;
  const parsed = tryGetOrParseFontBinary(bytes, knownIdentity);
  if (
    parsed === undefined ||
    !parsed.hasGlyf ||
    parsed.hasFvar ||
    parsed.hasColorOrBitmapTable ||
    parsed.sfnt === undefined ||
    !parsed.sfnt.tables.has("head") ||
    !parsed.sfnt.tables.has("name") ||
    parsed.fsType === undefined
  ) {
    return false;
  }
  if (!matchesLegacyUnicodeFamily(parsed, family)) {
    return false;
  }
  const os2 = parsed.sfnt.tables.get("OS/2");
  if (os2 === undefined || os2.length < 64) {
    return false;
  }
  return checkFsTypeEmbeddability(parsed.fsType);
}

/** Validates that a font face satisfies OOXML static TrueType EOT embedding requirements. */
export function validateEmbeddableTrueTypeFont(
  input: FontEmbeddingContext & {
    readonly identity: string;
    readonly bytes?: Uint8Array;
    readonly substitute: boolean;
    readonly embeddable: boolean;
    readonly bold: boolean;
    readonly italic: boolean;
  },
  mode: FontEmbeddingMode,
): CheckedEmbeddableTrueType {
  if (!input.embeddable) {
    return failFont(
      input,
      "font-not-embeddable",
      `${input.family} explicitly declares embeddable: false and cannot satisfy font embedding request`,
    );
  }
  if (input.substitute) {
    const parsed =
      input.bytes !== undefined
        ? tryGetOrParseFontBinary(input.bytes, input.identity)
        : undefined;
    const nativeLegacy = parsed?.names.primaryLegacyFamily;
    if (
      parsed !== undefined &&
      matchesLegacyUnicodeFamily(parsed, input.family)
    ) {
      const nativeName = nativeLegacy ?? input.family;
      return failFont(
        input,
        "font-substitute-not-embeddable",
        `Font "${input.family}" sets substitute: true, but the font file's native Name ID 1 ("${nativeName}") already matches "${input.family}"; remove redundant substitute: true to enable embedding.`,
      );
    }
    if (nativeLegacy !== undefined) {
      return failFont(
        input,
        "font-substitute-not-embeddable",
        `Font "${input.family}" sets substitute: true with a font file whose native Name ID 1 is "${nativeLegacy}"; substitute fonts cannot be embedded under another family name.`,
      );
    }
    return failFont(
      input,
      "font-substitute-not-embeddable",
      `Font "${input.family}" uses substitute font metrics and cannot satisfy embedding request for the declared family`,
    );
  }
  if (input.bytes === undefined || input.identity === "missing") {
    return failFont(
      input,
      "font-embedding-resource-missing",
      `${input.family} ${input.weight} ${input.style} is missing embeddable font resource`,
    );
  }

  const actualIdentity = bytesSha256(input.bytes);
  if (actualIdentity !== input.identity) {
    return failFont(
      input,
      "invalid-font-face",
      `Font resource identity for ${input.family} does not match actual bytes`,
    );
  }

  let parsed: ParsedFontBinary;
  try {
    parsed = getOrParseFontBinary(input.bytes, actualIdentity);
  } catch (error) {
    return failFont(
      input,
      "unsupported-font-format",
      `Failed to parse TrueType font tables for ${input.family}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  if (parsed.hasFvar) {
    return failFont(
      input,
      "unsupported-font-format",
      `Variable font "${input.family}" contains an OpenType "fvar" axis table and cannot be embedded; supply static per-weight TrueType font files instead`,
    );
  }
  if (parsed.hasColorOrBitmapTable) {
    return failFont(
      input,
      "unsupported-font-format",
      `Font embedding only supports monochrome static TrueType outlines; color or bitmap font "${input.family}" is not supported`,
    );
  }
  if (
    parsed.format !== "TTF" ||
    !parsed.hasGlyf ||
    parsed.sfnt === undefined ||
    !parsed.sfnt.tables.has("OS/2") ||
    !parsed.sfnt.tables.has("head") ||
    !parsed.sfnt.tables.has("name")
  ) {
    return failFont(
      input,
      "unsupported-font-format",
      "Font embedding only supports static TrueType (glyf) outlines; CFF OpenType (.otf), WOFF, WOFF2, and font collections are not supported",
    );
  }

  const os2 = parsed.sfnt.tables.get("OS/2")!;
  if (
    os2.length < 64 ||
    parsed.fsType === undefined ||
    parsed.weightClass === undefined
  ) {
    return failFont(
      input,
      "unsupported-font-format",
      "Font OS/2 table is missing weight, fsType, or style selection fields",
    );
  }

  const fsType = parsed.fsType;
  const usagePermission = fsType & 0x000f;
  if (![0, 2, 4, 8].includes(usagePermission) || (fsType & ~0x030f) !== 0) {
    return failFont(
      input,
      "invalid-font-permissions",
      `OS/2 fsType of ${input.family} contains invalid or reserved permission flags`,
    );
  }
  if (usagePermission === 2) {
    return failFont(
      input,
      "font-embedding-restricted",
      `OS/2 fsType of ${input.family} restricts embedding`,
    );
  }
  if (usagePermission === 4) {
    return failFont(
      input,
      "font-embedding-preview-print",
      `OS/2 fsType of ${input.family} permits only preview and print, which cannot satisfy font embedding for editable presentations`,
    );
  }
  if (mode === "subset" && (fsType & 0x0100) !== 0) {
    return failFont(
      input,
      "font-subsetting-forbidden",
      `OS/2 fsType of ${input.family} forbids subsetting; switch to full embedding or omit embedding request`,
    );
  }
  if ((fsType & 0x0200) !== 0) {
    return failFont(
      input,
      "font-embedding-bitmap-only",
      `OS/2 fsType of ${input.family} allows bitmap embedding only and cannot embed TrueType outlines`,
    );
  }

  if (!matchesLegacyUnicodeFamily(parsed, input.family)) {
    const nativeLegacy = parsed.names.primaryLegacyFamily;
    if (parsed.names.preferredFamilyKeys.has(fontFamilyKey(input.family))) {
      const nativeLabel =
        nativeLegacy ?? "an unspecified Name ID 1 legacy family";
      return failFont(
        input,
        "font-family-mismatch",
        `Font "${input.family}" matches typographic Name ID 16 only, whereas embedding requires matching the font file's native Name ID 1 legacy family "${nativeLabel}". Register an independent font("${nativeLabel}", ...) for embedding, or disable embedFonts.`,
      );
    }
    const nativeLabel =
      nativeLegacy ??
      [...parsed.names.preferredFamilies][0] ??
      parsed.font.familyName ??
      "unknown";
    return failFont(
      input,
      "font-family-mismatch",
      `Font file's native Name ID 1 family "${nativeLabel}" does not match declared family "${input.family}". Register font("${nativeLabel}", ...) for embedding, or set substitute: true for unembedded measurement only.`,
    );
  }
  if (parsed.weightClass !== input.weight) {
    return failFont(
      input,
      "font-weight-mismatch",
      `${input.family} requested ${input.weight} weight, but OS/2 usWeightClass is ${parsed.weightClass}`,
    );
  }
  if (
    parsed.os2Bold !== input.bold ||
    parsed.isItalic !== input.italic ||
    parsed.isItalic !== (input.style === "italic")
  ) {
    return failFont(
      input,
      "font-style-mismatch",
      `OS/2 fsSelection of ${input.family} does not match bold or italic styles emitted for text`,
    );
  }

  return {
    identity: parsed.identity,
    primaryLegacyFamily: parsed.names.primaryLegacyFamily,
    weightClass: parsed.weightClass,
    fsType,
    os2Bold: parsed.os2Bold,
    isItalic: parsed.isItalic,
    pitchFamily: parsed.pitchFamily,
    charset: parsed.charset,
  };
}

function getEotNameUtf16Le(
  names: UnicodeFontNames,
  id: number,
  context: FontEmbeddingContext,
): Uint8Array {
  if (names.nameTableError !== undefined) {
    return failFont(context, "unsupported-font-format", names.nameTableError);
  }
  const field = names.eotUtf16LeFields.get(id);
  if (field === undefined) {
    return failFont(
      context,
      "unsupported-font-format",
      `Font name table is missing Unicode name field ${id} required for EOT`,
    );
  }
  return field;
}

function subsetTrueTypeAndVerify(
  parsed: ParsedFontBinary,
  checked: CheckedEmbeddableTrueType,
  characters: string,
  context: FontEmbeddingContext,
  requiresVisibleGlyph: (character: string) => boolean,
): Uint8Array {
  const requiredCharacters = [...new Set(characters)];
  const missingCharacters = requiredCharacters.filter((character) => {
    const codePoint = character.codePointAt(0);
    return (
      codePoint !== undefined &&
      requiresVisibleGlyph(character) &&
      (!parsed.font.hasGlyphForCodePoint(codePoint) ||
        parsed.font.glyphForCodePoint(codePoint).id === 0)
    );
  });
  if (missingCharacters.length > 0) {
    return failFont(
      context,
      "font-missing-glyph",
      `${context.family} ${context.weight} ${context.style} font is missing glyph for emitted character ${missingCharacters[0]!}`,
      { characters: missingCharacters },
    );
  }

  let subsetBytes: Uint8Array;
  try {
    // Coverage check ignores non-glyph codepoints, but the original character stream is passed to HarfBuzz to retain layout closure.
    subsetBytes = subsetFont(parsed.bytes, characters);
  } catch (error) {
    return failFont(
      context,
      "font-subsetting-failed",
      `Failed to generate character subset for ${context.family}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  let subsetParsed: ParsedFontBinary;
  try {
    subsetParsed = parseFontBinaryUncached(
      subsetBytes,
      bytesSha256(subsetBytes),
    );
    if (
      subsetParsed.format !== "TTF" ||
      !subsetParsed.hasGlyf ||
      subsetParsed.hasFvar ||
      subsetParsed.sfnt === undefined ||
      !subsetParsed.sfnt.tables.has("OS/2") ||
      !subsetParsed.sfnt.tables.has("head") ||
      !subsetParsed.sfnt.tables.has("name")
    ) {
      throw new Error("Subset is not a static TrueType font");
    }
  } catch (error) {
    return failFont(
      context,
      "font-subsetting-failed",
      `Failed to parse character subset for ${context.family} independently: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  if (
    subsetParsed.weightClass !== checked.weightClass ||
    subsetParsed.font["OS/2"].fsType.editable !==
      parsed.font["OS/2"].fsType.editable ||
    subsetParsed.font["OS/2"].fsType.noSubsetting !==
      parsed.font["OS/2"].fsType.noSubsetting ||
    subsetParsed.os2Bold !== parsed.os2Bold ||
    subsetParsed.isItalic !== checked.isItalic
  ) {
    return failFont(
      context,
      "font-subsetting-failed",
      `Character subset for ${context.family} did not retain embedding permissions and styles of the original font`,
    );
  }
  for (const character of requiredCharacters) {
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      requiresVisibleGlyph(character) &&
      !subsetParsed.font.hasGlyphForCodePoint(codePoint)
    ) {
      return failFont(
        context,
        "font-subsetting-failed",
        `Character subset for ${context.family} did not retain emitted character ${character}`,
      );
    }
  }
  return subsetBytes;
}

function packUncompressedEot(
  parsed: ParsedFontBinary,
  checked: CheckedEmbeddableTrueType,
  payloadBytes: Uint8Array,
  context: FontEmbeddingContext,
): Uint8Array {
  const sourceBytes = parsed.bytes;
  const sourceSfnt = parsed.sfnt!;
  const os2 = sourceSfnt.tables.get("OS/2")!;
  const view = new DataView(
    sourceBytes.buffer,
    sourceBytes.byteOffset,
    sourceBytes.byteLength,
  );
  let payloadSfnt: SfntDirectory;
  try {
    payloadSfnt = readSfntDirectory(payloadBytes);
  } catch (error) {
    return failFont(
      context,
      "unsupported-font-format",
      error instanceof Error ? error.message : String(error),
    );
  }
  const payloadHead = payloadSfnt.tables.get("head");
  if (payloadHead === undefined) {
    return failFont(
      context,
      "unsupported-font-format",
      "Font payload is missing head table",
    );
  }
  const payloadView = new DataView(
    payloadBytes.buffer,
    payloadBytes.byteOffset,
    payloadBytes.byteLength,
  );
  const os2Version = view.getUint16(os2.offset);
  const field32 = (offset: number) =>
    offset + 4 <= os2.length ? view.getUint32(os2.offset + offset) : 0;
  const panose = sourceBytes.slice(os2.offset + 32, os2.offset + 42);
  const chunks: Uint8Array[] = [];
  const u8 = (value: number) => chunks.push(Uint8Array.of(value));
  const u16 = (value: number) => {
    const chunk = new Uint8Array(2);
    new DataView(chunk.buffer).setUint16(0, value, true);
    chunks.push(chunk);
  };
  const u32 = (value: number) => {
    const chunk = new Uint8Array(4);
    new DataView(chunk.buffer).setUint32(0, value >>> 0, true);
    chunks.push(chunk);
  };
  u32(0);
  u32(payloadBytes.byteLength);
  u32(0x00020002);
  u32(0);
  chunks.push(panose);
  u8(1);
  u8(checked.isItalic ? 1 : 0);
  u32(checked.weightClass);
  u16(checked.fsType);
  u16(0x504c);
  for (const offset of [42, 46, 50, 54]) u32(field32(offset));
  for (const offset of [78, 82]) {
    u32(os2Version >= 1 && os2.length >= offset + 4 ? field32(offset) : 0);
  }
  u32(payloadView.getUint32(payloadHead.offset + 8));
  for (let index = 0; index < 4; index += 1) u32(0);
  for (const id of [1, 2, 5, 4]) {
    const name = getEotNameUtf16Le(parsed.names, id, context);
    u16(0);
    u16(name.byteLength);
    chunks.push(name);
  }
  u16(0);
  u16(0);
  u32(0x50475342);
  u32(0);
  u16(0);
  u16(0);
  u32(0);
  u32(0);
  chunks.push(payloadBytes);
  const size = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  new DataView(output.buffer).setUint32(0, size, true);
  return output;
}

/**
 * Generates an uncompressed EOT (`.fntdata`) container for an already-validated
 * static TrueType font in either `"full"` or `"subset"` embedding mode.
 */
export function buildEmbeddedEotPayload(
  sourceBytes: Uint8Array,
  checked: CheckedEmbeddableTrueType,
  mode: FontEmbeddingMode,
  characters: string | undefined,
  context: FontEmbeddingContext,
  requiresVisibleGlyph: (character: string) => boolean,
): Uint8Array {
  const parsed = getOrParseFontBinary(sourceBytes, checked.identity);
  if (parsed.sfnt === undefined) {
    return failFont(
      context,
      "unsupported-font-format",
      `Font data for ${context.family} is not a valid SFNT file`,
    );
  }
  const payloadBytes =
    mode === "subset"
      ? subsetTrueTypeAndVerify(
          parsed,
          checked,
          characters ?? "",
          context,
          requiresVisibleGlyph,
        )
      : parsed.bytes;
  return packUncompressedEot(parsed, checked, payloadBytes, context);
}

/**
 * Inspects unpacked SFNT bytes from an imported PPTX embedded font part, returning
 * family match status and `1..1000` weight / unified italic metadata if valid TrueType.
 */
export function inspectImportedEmbeddedFont(
  ttfBytes: Uint8Array,
  declaredTypeface: string,
): ImportedEmbeddedFontInfo | undefined {
  const parsed = tryGetOrParseFontBinary(ttfBytes);
  if (parsed === undefined || parsed.format !== "TTF" || parsed.hasFvar) {
    return undefined;
  }
  return {
    familyMatches: matchesAnyUnicodeFamily(parsed, declaredTypeface),
    weight: parsed.weightClass,
    style: parsed.hasOs2 ? parsed.style : undefined,
  };
}
