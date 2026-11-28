function asciiTag(bytes: Uint8Array, offset: number): string {
  return String.fromCharCode(
    bytes[offset]!,
    bytes[offset + 1]!,
    bytes[offset + 2]!,
    bytes[offset + 3]!,
  );
}

function writeAsciiTag(target: Uint8Array, offset: number, tag: string): void {
  for (let index = 0; index < 4; index += 1) {
    target[offset + index] = tag.charCodeAt(index);
  }
}

function sfntChecksum(bytes: Uint8Array): number {
  const paddedLength = Math.ceil(bytes.byteLength / 4) * 4;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let sum = 0;
  for (let offset = 0; offset < paddedLength; offset += 4) {
    if (offset + 4 <= bytes.byteLength) {
      sum = (sum + view.getUint32(offset)) >>> 0;
    } else {
      let word = 0;
      for (let byte = 0; byte < 4; byte += 1) {
        const value =
          offset + byte < bytes.byteLength ? bytes[offset + byte]! : 0;
        word = (word << 8) | value;
      }
      sum = (sum + (word >>> 0)) >>> 0;
    }
  }
  return sum >>> 0;
}

function packSfnt(
  sfntVersion: number,
  tables: ReadonlyMap<string, Uint8Array>,
): Uint8Array {
  const sorted = [...tables.entries()].toSorted(([a], [b]) =>
    a < b ? -1 : a > b ? 1 : 0,
  );
  const numTables = sorted.length;
  const headerSize = 12 + numTables * 16;
  let totalSize = headerSize;
  for (const [, data] of sorted) {
    totalSize += Math.ceil(data.byteLength / 4) * 4;
  }

  const output = new Uint8Array(totalSize);
  const view = new DataView(output.buffer);
  view.setUint32(0, sfntVersion >>> 0);
  view.setUint16(4, numTables);
  const entrySelector = Math.floor(Math.log2(Math.max(1, numTables)));
  const searchRange = (1 << entrySelector) * 16;
  view.setUint16(6, searchRange);
  view.setUint16(8, entrySelector);
  view.setUint16(10, numTables * 16 - searchRange);

  let dataOffset = headerSize;
  let headOffset: number | undefined;
  for (let index = 0; index < sorted.length; index += 1) {
    const [tag, raw] = sorted[index]!;
    const tableBytes =
      tag === "head" && raw.byteLength >= 12 ? new Uint8Array(raw) : raw;
    if (tag === "head" && tableBytes.byteLength >= 12) {
      new DataView(tableBytes.buffer).setUint32(8, 0);
      headOffset = dataOffset;
    }
    output.set(tableBytes, dataOffset);
    const recordOffset = 12 + index * 16;
    writeAsciiTag(output, recordOffset, tag);
    view.setUint32(recordOffset + 4, sfntChecksum(tableBytes));
    view.setUint32(recordOffset + 8, dataOffset);
    view.setUint32(recordOffset + 12, tableBytes.byteLength);
    dataOffset += Math.ceil(tableBytes.byteLength / 4) * 4;
  }

  if (headOffset !== undefined) {
    const adjustment = (0xb1b0afba - sfntChecksum(output)) >>> 0;
    view.setUint32(headOffset + 8, adjustment);
  }
  return output;
}

const variableFontTables = new Set([
  "fvar",
  "gvar",
  "avar",
  "STAT",
  "HVAR",
  "VVAR",
  "MVAR",
]);

/**
 * Strips OpenType variable-font axis and variation delta tables so static default-instance
 * TrueType fixtures can be exercised without triggering variable-font rejection.
 */
export function stripVariableFontTables(source: Uint8Array): Uint8Array {
  const view = new DataView(
    source.buffer,
    source.byteOffset,
    source.byteLength,
  );
  const sfntVersion = view.getUint32(0);
  const numTables = view.getUint16(4);
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < numTables; index += 1) {
    const recordOffset = 12 + index * 16;
    const tag = asciiTag(source, recordOffset);
    if (variableFontTables.has(tag)) continue;
    const offset = view.getUint32(recordOffset + 8);
    const length = view.getUint32(recordOffset + 12);
    tables.set(tag, source.slice(offset, offset + length));
  }
  return packSfnt(sfntVersion, tables);
}

export type SfntNameRecordInput = {
  readonly platformID: number;
  readonly encodingID: number;
  readonly languageID: number;
  readonly nameID: number;
  readonly value: string;
};

/**
 * Replaces an SFNT font's `name` table with spec-sorted UTF-16BE records and recomputes table checksums.
 */
export function replaceSfntNameTable(
  source: Uint8Array,
  records: readonly SfntNameRecordInput[],
): Uint8Array {
  const view = new DataView(
    source.buffer,
    source.byteOffset,
    source.byteLength,
  );
  const sfntVersion = view.getUint32(0);
  const numTables = view.getUint16(4);
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < numTables; index += 1) {
    const recordOffset = 12 + index * 16;
    const tag = asciiTag(source, recordOffset);
    const offset = view.getUint32(recordOffset + 8);
    const length = view.getUint32(recordOffset + 12);
    tables.set(tag, source.slice(offset, offset + length));
  }

  const sortedRecords = [...records]
    .map((entry) => ({
      ...entry,
      bytes: utf16BeBytes(entry.value),
    }))
    .toSorted(
      (a, b) =>
        a.platformID - b.platformID ||
        a.encodingID - b.encodingID ||
        a.languageID - b.languageID ||
        a.nameID - b.nameID,
    );
  const headerSize = 6 + sortedRecords.length * 12;
  const totalStringBytes = sortedRecords.reduce(
    (sum, item) => sum + item.bytes.byteLength,
    0,
  );
  const nameTable = new Uint8Array(headerSize + totalStringBytes);
  const nameView = new DataView(nameTable.buffer);
  nameView.setUint16(0, 0);
  nameView.setUint16(2, sortedRecords.length);
  nameView.setUint16(4, headerSize);

  let strOffset = 0;
  for (let index = 0; index < sortedRecords.length; index += 1) {
    const entry = sortedRecords[index]!;
    const recOffset = 6 + index * 12;
    nameView.setUint16(recOffset, entry.platformID);
    nameView.setUint16(recOffset + 2, entry.encodingID);
    nameView.setUint16(recOffset + 4, entry.languageID);
    nameView.setUint16(recOffset + 6, entry.nameID);
    nameView.setUint16(recOffset + 8, entry.bytes.byteLength);
    nameView.setUint16(recOffset + 10, strOffset);
    nameTable.set(entry.bytes, headerSize + strOffset);
    strOffset += entry.bytes.byteLength;
  }

  tables.set("name", nameTable);
  return packSfnt(sfntVersion, tables);
}

/**
 * Patches `OS/2` (`usWeightClass`, `fsSelection`) and `post` (`italicAngle`) fields
 * on an SFNT font and recomputes SFNT checksums.
 */
export function patchSfntOs2AndPost(
  source: Uint8Array,
  patch: {
    readonly weightClass?: number;
    readonly italic?: boolean;
    readonly oblique?: boolean;
    readonly bold?: boolean;
    readonly italicAngle?: number;
  },
): Uint8Array {
  const view = new DataView(
    source.buffer,
    source.byteOffset,
    source.byteLength,
  );
  const sfntVersion = view.getUint32(0);
  const numTables = view.getUint16(4);
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < numTables; index += 1) {
    const recordOffset = 12 + index * 16;
    const tag = asciiTag(source, recordOffset);
    const offset = view.getUint32(recordOffset + 8);
    const length = view.getUint32(recordOffset + 12);
    const tableBytes = source.slice(offset, offset + length);
    const tableView = new DataView(tableBytes.buffer);
    if (tag === "OS/2") {
      if (patch.weightClass !== undefined) {
        tableView.setUint16(4, patch.weightClass);
      }
      let fsSelection = tableView.getUint16(62);
      if (patch.italic !== undefined) {
        fsSelection = patch.italic
          ? fsSelection | 0x0001
          : fsSelection & ~0x0001;
      }
      if (patch.bold !== undefined) {
        fsSelection = patch.bold ? fsSelection | 0x0020 : fsSelection & ~0x0020;
      }
      if (patch.oblique !== undefined) {
        fsSelection = patch.oblique
          ? fsSelection | 0x0200
          : fsSelection & ~0x0200;
      }
      tableView.setUint16(62, fsSelection);
    } else if (tag === "post" && patch.italicAngle !== undefined) {
      tableView.setInt32(4, Math.round(patch.italicAngle * 65536));
    }
    tables.set(tag, tableBytes);
  }
  return packSfnt(sfntVersion, tables);
}

function encodeCffInt(value: number): Uint8Array {
  if (value >= -107 && value <= 107) {
    return Uint8Array.of(value + 139);
  }
  if (value >= 108 && value <= 1131) {
    const v = value - 108;
    return Uint8Array.of((v >> 8) + 247, v & 0xff);
  }
  if (value >= -1131 && value <= -108) {
    const v = -value - 108;
    return Uint8Array.of((v >> 8) + 251, v & 0xff);
  }
  const bytes = new Uint8Array(5);
  bytes[0] = 29;
  new DataView(bytes.buffer).setInt32(1, value);
  return bytes;
}

function encodeCffInt32(value: number): Uint8Array {
  const bytes = new Uint8Array(5);
  bytes[0] = 29;
  new DataView(bytes.buffer).setInt32(1, value);
  return bytes;
}

function encodeCffIndex(entries: readonly Uint8Array[]): Uint8Array {
  if (entries.length === 0) {
    return Uint8Array.of(0, 0);
  }
  const totalData = entries.reduce((sum, entry) => sum + entry.byteLength, 0);
  const offSize = totalData + 1 < 0x100 ? 1 : totalData + 1 < 0x10000 ? 2 : 4;
  const headerSize = 3 + (entries.length + 1) * offSize;
  const output = new Uint8Array(headerSize + totalData);
  const view = new DataView(output.buffer);
  view.setUint16(0, entries.length);
  output[2] = offSize;

  let offset = 1;
  let dataPos = headerSize;
  for (let index = 0; index <= entries.length; index += 1) {
    const pos = 3 + index * offSize;
    if (offSize === 1) {
      output[pos] = offset;
    } else if (offSize === 2) {
      view.setUint16(pos, offset);
    } else {
      view.setUint32(pos, offset);
    }
    if (index < entries.length) {
      const entry = entries[index]!;
      output.set(entry, dataPos);
      offset += entry.byteLength;
      dataPos += entry.byteLength;
    }
  }
  return output;
}

function concatBytes(chunks: readonly Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const output = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

function utf16BeBytes(value: string): Uint8Array {
  const output = new Uint8Array(value.length * 2);
  const view = new DataView(output.buffer);
  for (let index = 0; index < value.length; index += 1) {
    view.setUint16(index * 2, value.charCodeAt(index));
  }
  return output;
}

export type MinimalCffOtfOptions = {
  readonly familyName?: string;
  readonly preferredFamily?: string;
  readonly weightClass?: number;
  readonly italic?: boolean;
  readonly oblique?: boolean;
  readonly bold?: boolean;
  readonly italicAngle?: number;
  readonly fsType?: number;
};

/**
 * Synthesizes a minimal monochrome static CFF OpenType (`.otf`) font covering basic ASCII
 * code points (`0x20..0x7E`), suitable for HarfBuzz shaping, SVG path extraction, and PNG preview.
 */
export function createMinimalCffOtf(
  options: MinimalCffOtfOptions = {},
): Uint8Array {
  const familyName = options.familyName ?? "Synthetic CFF";
  const weightClass = options.weightClass ?? 400;
  const italic = options.italic ?? false;
  const oblique = options.oblique ?? false;
  const bold = options.bold ?? weightClass >= 600;
  const italicAngle = options.italicAngle ?? (italic || oblique ? -12 : 0);
  const fsType = options.fsType ?? 0;

  // Glyph 0: .notdef; glyphs 1..95: ASCII U+0020..U+007E (CFF Standard Strings SIDs 1..95).
  const firstCodePoint = 0x20;
  const lastCodePoint = 0x7e;
  const glyphCount = 1 + (lastCodePoint - firstCodePoint + 1);
  const advanceWidth = 560;

  const charStrings: Uint8Array[] = [
    // .notdef
    Uint8Array.of(14),
  ];
  const sids: number[] = [];
  for (let code = firstCodePoint; code <= lastCodePoint; code += 1) {
    sids.push(code - 0x20 + 1);
    if (code === 0x20) {
      charStrings.push(concatBytes([encodeCffInt(300), Uint8Array.of(14)]));
    } else {
      // Simple 440x680 rectangle starting at (60, 0) in 1000-UPM space.
      charStrings.push(
        concatBytes([
          encodeCffInt(advanceWidth),
          encodeCffInt(60),
          encodeCffInt(0),
          Uint8Array.of(21), // rmoveto
          encodeCffInt(440),
          Uint8Array.of(6), // hlineto
          encodeCffInt(680),
          Uint8Array.of(7), // vlineto
          encodeCffInt(-440),
          Uint8Array.of(6), // hlineto
          Uint8Array.of(14), // endchar
        ]),
      );
    }
  }

  const charsetBytes = new Uint8Array(1 + sids.length * 2);
  const charsetView = new DataView(charsetBytes.buffer);
  charsetBytes[0] = 0;
  for (let index = 0; index < sids.length; index += 1) {
    charsetView.setUint16(1 + index * 2, sids[index]!);
  }

  const headerBytes = Uint8Array.of(1, 0, 4, 1);
  const nameIndex = encodeCffIndex([
    new TextEncoder().encode(
      familyName.replace(/[^A-Za-z0-9_-]/gu, "") || "SyntheticCFF",
    ),
  ]);
  const stringIndex = encodeCffIndex([]);
  const globalSubrIndex = encodeCffIndex([]);
  const charStringsIndex = encodeCffIndex(charStrings);
  const privateDict = concatBytes([
    encodeCffInt(0),
    Uint8Array.of(20), // defaultWidthX
  ]);

  const buildTopDict = (
    charsetOffset: number,
    charStringsOffset: number,
    privateOffset: number,
  ): Uint8Array =>
    concatBytes([
      encodeCffInt32(charsetOffset),
      Uint8Array.of(15), // charset
      encodeCffInt32(charStringsOffset),
      Uint8Array.of(17), // CharStrings
      encodeCffInt32(privateDict.byteLength),
      encodeCffInt32(privateOffset),
      Uint8Array.of(18), // Private
    ]);

  const placeholderTopDictIndex = encodeCffIndex([buildTopDict(0, 0, 0)]);
  const prefixLength =
    headerBytes.byteLength +
    nameIndex.byteLength +
    placeholderTopDictIndex.byteLength +
    stringIndex.byteLength +
    globalSubrIndex.byteLength;
  const charsetOffset = prefixLength;
  const charStringsOffset = charsetOffset + charsetBytes.byteLength;
  const privateOffset = charStringsOffset + charStringsIndex.byteLength;
  const topDictIndex = encodeCffIndex([
    buildTopDict(charsetOffset, charStringsOffset, privateOffset),
  ]);

  const cffTable = concatBytes([
    headerBytes,
    nameIndex,
    topDictIndex,
    stringIndex,
    globalSubrIndex,
    charsetBytes,
    charStringsIndex,
    privateDict,
  ]);

  // head (54 bytes)
  const head = new Uint8Array(54);
  const headView = new DataView(head.buffer);
  headView.setUint32(0, 0x00010000);
  headView.setUint32(4, 0x00010000);
  headView.setUint32(12, 0x5f0f3cf5); // magicNumber
  headView.setUint16(16, 0x0003); // flags
  headView.setUint16(18, 1000); // unitsPerEm
  headView.setInt16(36, 0); // xMin
  headView.setInt16(38, -200); // yMin
  headView.setInt16(40, 560); // xMax
  headView.setInt16(42, 800); // yMax
  headView.setUint16(44, (bold ? 1 : 0) | (italic ? 2 : 0));
  headView.setUint16(46, 8); // lowestRecPPEM

  // hhea (36 bytes)
  const hhea = new Uint8Array(36);
  const hheaView = new DataView(hhea.buffer);
  hheaView.setUint32(0, 0x00010000);
  hheaView.setInt16(4, 800); // ascent
  hheaView.setInt16(6, -200); // descent
  hheaView.setInt16(8, 0); // lineGap
  hheaView.setUint16(10, advanceWidth); // advanceWidthMax
  hheaView.setInt16(12, 0); // minLeftSideBearing
  hheaView.setInt16(14, 0); // minRightSideBearing
  hheaView.setInt16(16, 500); // xMaxExtent
  hheaView.setInt16(18, 1); // caretSlopeRise
  hheaView.setUint16(34, glyphCount); // numberOfHMetrics

  // maxp v0.5 (6 bytes for CFF)
  const maxp = new Uint8Array(6);
  const maxpView = new DataView(maxp.buffer);
  maxpView.setUint32(0, 0x00005000);
  maxpView.setUint16(4, glyphCount);

  // hmtx
  const hmtx = new Uint8Array(glyphCount * 4);
  const hmtxView = new DataView(hmtx.buffer);
  for (let gid = 0; gid < glyphCount; gid += 1) {
    const adv = gid === 1 ? 300 : advanceWidth;
    hmtxView.setUint16(gid * 4, adv);
    hmtxView.setInt16(gid * 4 + 2, gid === 0 || gid === 1 ? 0 : 60);
  }

  // OS/2 v3 (96 bytes)
  const os2 = new Uint8Array(96);
  const os2View = new DataView(os2.buffer);
  os2View.setUint16(0, 3); // version
  os2View.setInt16(2, advanceWidth); // xAvgCharWidth
  os2View.setUint16(4, weightClass);
  os2View.setUint16(6, 5); // usWidthClass = Medium
  os2View.setUint16(8, fsType);
  os2View.setInt16(26, 50); // yStrikeoutSize
  os2View.setInt16(28, 300); // yStrikeoutPosition
  os2[32] = 2; // PANOSE family = Latin Text
  os2[33] = 11; // PANOSE serif = Normal Sans
  os2View.setUint32(42, 1); // ulUnicodeRange1 Basic Latin
  writeAsciiTag(os2, 58, " synthetic".slice(0, 4));
  let fsSelection = 0;
  if (italic) fsSelection |= 0x0001;
  if (bold) fsSelection |= 0x0020;
  if (!italic && !bold && !oblique) fsSelection |= 0x0040;
  if (oblique) fsSelection |= 0x0200;
  os2View.setUint16(62, fsSelection);
  os2View.setUint16(64, firstCodePoint);
  os2View.setUint16(66, lastCodePoint);
  os2View.setInt16(68, 800); // sTypoAscender
  os2View.setInt16(70, -200); // sTypoDescender
  os2View.setInt16(72, 0); // sTypoLineGap
  os2View.setUint16(74, 800); // usWinAscent
  os2View.setUint16(76, 200); // usWinDescent

  // post v3.0 (32 bytes)
  const post = new Uint8Array(32);
  const postView = new DataView(post.buffer);
  postView.setUint32(0, 0x00030000);
  // Fixed 16.16 italicAngle
  const fixedAngle = Math.round(italicAngle * 65536);
  postView.setInt32(4, fixedAngle);
  postView.setInt16(8, -100); // underlinePosition
  postView.setInt16(10, 50); // underlineThickness

  // cmap format 4 (single contiguous range 0x20..0x7E -> GID 1..95 + 0xFFFF sentinel)
  const cmap = new Uint8Array(44);
  const cmapView = new DataView(cmap.buffer);
  cmapView.setUint16(0, 0); // version
  cmapView.setUint16(2, 1); // numTables
  cmapView.setUint16(4, 3); // platformID Windows
  cmapView.setUint16(6, 1); // encodingID Unicode BMP
  cmapView.setUint32(8, 12); // offset to subtable
  // Format 4 subtable at offset 12 (32 bytes)
  cmapView.setUint16(12, 4); // format
  cmapView.setUint16(14, 32); // length
  cmapView.setUint16(16, 0); // language
  cmapView.setUint16(18, 4); // segCountX2 (2 segments)
  cmapView.setUint16(20, 4); // searchRange
  cmapView.setUint16(22, 1); // entrySelector
  cmapView.setUint16(24, 0); // rangeShift
  cmapView.setUint16(26, lastCodePoint); // endCode[0]
  cmapView.setUint16(28, 0xffff); // endCode[1]
  cmapView.setUint16(30, 0); // reservedPad
  cmapView.setUint16(32, firstCodePoint); // startCode[0]
  cmapView.setUint16(34, 0xffff); // startCode[1]
  cmapView.setInt16(36, 1 - firstCodePoint); // idDelta[0]
  cmapView.setInt16(38, 1); // idDelta[1]
  cmapView.setUint16(40, 0); // idRangeOffset[0]
  cmapView.setUint16(42, 0); // idRangeOffset[1]

  // name table (Platform 3 Encoding 1 Language 0x0409)
  const subfamily =
    bold && italic
      ? "Bold Italic"
      : bold
        ? "Bold"
        : italic
          ? "Italic"
          : "Regular";
  const nameEntries: readonly (readonly [number, string])[] = [
    [1, familyName],
    [2, subfamily],
    [4, `${familyName} ${subfamily}`],
    [5, "Version 1.000"],
    [6, `${familyName.replace(/\s+/gu, "")}-${subfamily.replace(/\s+/gu, "")}`],
    ...(options.preferredFamily === undefined
      ? []
      : ([[16, options.preferredFamily]] as const)),
  ];
  const encodedNames = nameEntries.map(([id, text]) => ({
    id,
    bytes: utf16BeBytes(text),
  }));
  const nameHeaderSize = 6 + encodedNames.length * 12;
  const totalStringBytes = encodedNames.reduce(
    (sum, item) => sum + item.bytes.byteLength,
    0,
  );
  const nameTable = new Uint8Array(nameHeaderSize + totalStringBytes);
  const nameView = new DataView(nameTable.buffer);
  nameView.setUint16(0, 0); // format
  nameView.setUint16(2, encodedNames.length);
  nameView.setUint16(4, nameHeaderSize);
  let strOffset = 0;
  for (let index = 0; index < encodedNames.length; index += 1) {
    const entry = encodedNames[index]!;
    const recOffset = 6 + index * 12;
    nameView.setUint16(recOffset, 3); // platformID Windows
    nameView.setUint16(recOffset + 2, 1); // encodingID Unicode BMP
    nameView.setUint16(recOffset + 4, 0x0409); // en-US
    nameView.setUint16(recOffset + 6, entry.id);
    nameView.setUint16(recOffset + 8, entry.bytes.byteLength);
    nameView.setUint16(recOffset + 10, strOffset);
    nameTable.set(entry.bytes, nameHeaderSize + strOffset);
    strOffset += entry.bytes.byteLength;
  }

  return packSfnt(
    0x4f54544f, // "OTTO"
    new Map<string, Uint8Array>([
      ["CFF ", cffTable],
      ["OS/2", os2],
      ["cmap", cmap],
      ["head", head],
      ["hhea", hhea],
      ["hmtx", hmtx],
      ["maxp", maxp],
      ["name", nameTable],
      ["post", post],
    ]),
  );
}

/**
 * Packs standalone SFNT binaries into a TrueType Collection (`ttcf` v1 or v2) container.
 */
export function buildTestTtc(
  sfnts: readonly Uint8Array[],
  version: 1 | 2 = 1,
): Uint8Array {
  const parsed = sfnts.map((bytes) => {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const sfntVersion = view.getUint32(0);
    const numTables = view.getUint16(4);
    const tables: { tag: Uint8Array; checksum: number; data: Uint8Array }[] =
      [];
    for (let i = 0; i < numTables; i += 1) {
      const rec = 12 + i * 16;
      const offset = view.getUint32(rec + 8);
      const length = view.getUint32(rec + 12);
      tables.push({
        tag: bytes.slice(rec, rec + 4),
        checksum: view.getUint32(rec + 4),
        data: bytes.slice(offset, offset + length),
      });
    }
    return { sfntVersion, tables };
  });

  const ttcHeaderSize = (version === 2 ? 24 : 12) + parsed.length * 4;
  let cursor = ttcHeaderSize;
  const dirOffsets: number[] = [];
  for (const fontEntry of parsed) {
    dirOffsets.push(cursor);
    cursor += 12 + fontEntry.tables.length * 16;
  }
  const tableOffsets: number[][] = [];
  for (const fontEntry of parsed) {
    const offsets: number[] = [];
    for (const tbl of fontEntry.tables) {
      offsets.push(cursor);
      cursor += Math.ceil(tbl.data.byteLength / 4) * 4;
    }
    tableOffsets.push(offsets);
  }

  const out = new Uint8Array(cursor);
  const view = new DataView(out.buffer);
  out.set([0x74, 0x74, 0x63, 0x66], 0);
  view.setUint16(4, version);
  view.setUint16(6, 0);
  view.setUint32(8, parsed.length);
  for (let i = 0; i < parsed.length; i += 1) {
    view.setUint32(12 + i * 4, dirOffsets[i]!);
  }
  for (let f = 0; f < parsed.length; f += 1) {
    const fontEntry = parsed[f]!;
    const dirOff = dirOffsets[f]!;
    const numTables = fontEntry.tables.length;
    view.setUint32(dirOff, fontEntry.sfntVersion);
    view.setUint16(dirOff + 4, numTables);
    const entrySelector = Math.floor(Math.log2(Math.max(1, numTables)));
    const searchRange = (1 << entrySelector) * 16;
    view.setUint16(dirOff + 6, searchRange);
    view.setUint16(dirOff + 8, entrySelector);
    view.setUint16(dirOff + 10, numTables * 16 - searchRange);
    for (let t = 0; t < numTables; t += 1) {
      const tbl = fontEntry.tables[t]!;
      const recOff = dirOff + 12 + t * 16;
      const dataOff = tableOffsets[f]![t]!;
      out.set(tbl.tag, recOff);
      view.setUint32(recOff + 4, tbl.checksum);
      view.setUint32(recOff + 8, dataOff);
      view.setUint32(recOff + 12, tbl.data.byteLength);
      out.set(tbl.data, dataOff);
    }
  }
  return out;
}
