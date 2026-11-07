import { createHash } from "node:crypto";

import {
  CapabilityError,
  xmlEscape,
  type CapabilityPresentationContribution,
} from "@react-pptx/core";
import * as fontkit from "fontkit";

import { fontFamilyKey, matchesFontFamily } from "./font-family.ts";
import { subsetFont } from "./font-subset.ts";
import type { FontFaceStyle } from "./font.ts";
import { requiresFontGlyph } from "./text-service.ts";

const fontRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const fontContentType = "application/x-fontdata";
type FontEmbeddingMode = "full" | "subset";

/** 文字序列化最终选择的字体文件和样式；缺少文件也会保留为实际用字要求。 */
export type FontEmbeddingUse = {
  readonly family: string;
  readonly weight: number;
  readonly style: FontFaceStyle;
  readonly claimKey: string;
  readonly identity: string;
  readonly bytes?: Uint8Array;
  readonly substitute: boolean;
  readonly embeddable: boolean;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly path?: string;
};

type FontTable = { readonly offset: number; readonly length: number };
type FontTables = ReadonlyMap<string, FontTable>;
type FontkitExtras = {
  readonly post?: { readonly isFixedPitch: number };
};

function fail(use: FontEmbeddingUse, code: string, message: string): never {
  throw new CapabilityError(code, message, use.path);
}

function tablesOf(bytes: Uint8Array, use: FontEmbeddingUse): FontTables {
  if (bytes.byteLength < 12) {
    return fail(
      use,
      "unsupported-font-format",
      "字体数据不是有效的 TrueType 文件",
    );
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  if (signature !== "\0\x01\0\0" && signature !== "true") {
    return fail(
      use,
      "unsupported-font-format",
      "字体嵌入目前只支持静态 TrueType 轮廓；WOFF、WOFF2、字体集合和 CFF 不受支持",
    );
  }
  const count = view.getUint16(4);
  if (12 + count * 16 > bytes.byteLength) {
    return fail(use, "unsupported-font-format", "字体的 SFNT 表目录不完整");
  }
  const tables = new Map<string, FontTable>();
  for (let index = 0; index < count; index += 1) {
    const record = 12 + index * 16;
    const name = String.fromCharCode(...bytes.subarray(record, record + 4));
    const offset = view.getUint32(record + 8);
    const length = view.getUint32(record + 12);
    if (offset > bytes.byteLength || length > bytes.byteLength - offset) {
      return fail(
        use,
        "unsupported-font-format",
        `字体表 ${name} 超出文件范围`,
      );
    }
    tables.set(name, { offset, length });
  }
  if (
    !tables.has("glyf") ||
    tables.has("fvar") ||
    !tables.has("OS/2") ||
    !tables.has("head") ||
    !tables.has("name")
  ) {
    return fail(
      use,
      "unsupported-font-format",
      "字体嵌入需要包含 OS/2、head、name、glyf 表的静态 TrueType 字体",
    );
  }
  const os2 = tables.get("OS/2")!;
  if (os2.length < 64) {
    return fail(
      use,
      "unsupported-font-format",
      "字体的 OS/2 表缺少字重、fsType 或样式选择字段",
    );
  }
  return tables;
}

function checkedFont(
  use: FontEmbeddingUse,
  mode: FontEmbeddingMode,
): {
  readonly font: fontkit.Font;
  readonly tables: FontTables;
  readonly fsType: number;
  readonly weightClass: number;
} {
  if (!use.embeddable) {
    return fail(
      use,
      "font-not-embeddable",
      `${use.family} 明确声明 embeddable: false，无法满足字体嵌入请求`,
    );
  }
  if (use.substitute) {
    return fail(
      use,
      "font-substitute-not-embeddable",
      `${use.family} 使用替代字体度量，不能满足声明字族的嵌入请求`,
    );
  }
  if (use.bytes === undefined || use.identity === "missing") {
    return fail(
      use,
      "font-embedding-resource-missing",
      `${use.family} ${use.weight} ${use.style} 缺少可嵌入的实际字体资源`,
    );
  }

  const identity = createHash("sha256").update(use.bytes).digest("hex");
  if (identity !== use.identity) {
    return fail(
      use,
      "invalid-font-face",
      `${use.family} 的字体资源身份与实际字节不一致`,
    );
  }
  const tables = tablesOf(use.bytes, use);
  const os2 = tables.get("OS/2")!;
  const view = new DataView(
    use.bytes.buffer,
    use.bytes.byteOffset,
    use.bytes.byteLength,
  );
  const fsType = view.getUint16(os2.offset + 8);
  const weightClass = view.getUint16(os2.offset + 4);
  const usagePermission = fsType & 0x000f;
  if (![0, 2, 4, 8].includes(usagePermission) || (fsType & ~0x030f) !== 0) {
    return fail(
      use,
      "invalid-font-permissions",
      `${use.family} 的 OS/2 fsType 含无效或保留许可标志`,
    );
  }
  if (usagePermission === 2) {
    return fail(
      use,
      "font-embedding-restricted",
      `${use.family} 的 OS/2 fsType 禁止嵌入`,
    );
  }
  if (usagePermission === 4) {
    return fail(
      use,
      "font-embedding-preview-print",
      `${use.family} 的 OS/2 fsType 仅允许预览和打印，不能满足可编辑文稿的字体嵌入请求`,
    );
  }
  if (mode === "subset" && (fsType & 0x0100) !== 0) {
    return fail(
      use,
      "font-subsetting-forbidden",
      `${use.family} 的 OS/2 fsType 禁止对子集化；请改为完整嵌入或取消嵌入请求`,
    );
  }
  if ((fsType & 0x0200) !== 0) {
    return fail(
      use,
      "font-embedding-bitmap-only",
      `${use.family} 的 OS/2 fsType 只允许嵌入位图，不能嵌入 TrueType 轮廓`,
    );
  }

  let opened: fontkit.Font | fontkit.FontCollection;
  try {
    opened = fontkit.create(
      Buffer.from(use.bytes.buffer, use.bytes.byteOffset, use.bytes.byteLength),
    );
  } catch {
    return fail(
      use,
      "unsupported-font-format",
      `${use.family} 的 TrueType 字体表无法解析`,
    );
  }
  if (opened.type !== "TTF") {
    return fail(
      use,
      "unsupported-font-format",
      `${use.family} 的字体数据不是受支持的静态 TrueType 字体`,
    );
  }
  const font = opened;
  if (!matchesFontFamily(font, use.family)) {
    return fail(
      use,
      "font-family-mismatch",
      `字体文件声明的字族 ${font.familyName} 与输出字族 ${use.family} 不一致；如用于度量请显式声明 substitute`,
    );
  }
  if (font.variationAxes && Object.keys(font.variationAxes).length > 0) {
    return fail(
      use,
      "unsupported-font-format",
      `${use.family} 是可变字体，当前无法嵌入`,
    );
  }
  if (weightClass !== use.weight) {
    return fail(
      use,
      "font-style-mismatch",
      `${use.family} 请求 ${use.weight} 字重，但 OS/2 usWeightClass 为 ${weightClass}`,
    );
  }
  if (
    font["OS/2"].fsSelection.bold !== use.bold ||
    font["OS/2"].fsSelection.italic !== use.italic ||
    font["OS/2"].fsSelection.italic !== (use.style === "italic")
  ) {
    return fail(
      use,
      "font-style-mismatch",
      `${use.family} 的 OS/2 fsSelection 与文字实际写出的粗体或斜体样式不一致`,
    );
  }
  return { font, tables, fsType, weightClass };
}

function subsetBytes(
  use: FontEmbeddingUse,
  checked: ReturnType<typeof checkedFont>,
  characters: string,
): Uint8Array {
  const requiredCharacters = [...new Set(characters)];
  for (const character of requiredCharacters) {
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      requiresFontGlyph(character) &&
      !checked.font.hasGlyphForCodePoint(codePoint)
    ) {
      return fail(
        use,
        "font-missing-glyph",
        `${use.family} 的 ${use.weight} ${use.style} 字体缺少实际输出字符 ${character}`,
      );
    }
  }

  let bytes: Uint8Array;
  try {
    // 覆盖检查忽略非字形码点，但原字符流仍交给 HarfBuzz 保留布局闭包。
    bytes = subsetFont(use.bytes!, characters);
  } catch (error) {
    return fail(
      use,
      "font-subsetting-failed",
      `${use.family} 的字符子集生成失败：${error instanceof Error ? error.message : String(error)}`,
    );
  }

  let font: fontkit.Font;
  try {
    const opened = fontkit.create(
      Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength),
    );
    if (opened.type !== "TTF") throw new Error("子集不是静态 TrueType 字体");
    font = opened;
    tablesOf(bytes, use);
  } catch (error) {
    return fail(
      use,
      "font-subsetting-failed",
      `${use.family} 的字符子集无法独立解析：${error instanceof Error ? error.message : String(error)}`,
    );
  }

  if (
    font["OS/2"].usWeightClass !== checked.weightClass ||
    font["OS/2"].fsType.editable !== checked.font["OS/2"].fsType.editable ||
    font["OS/2"].fsType.noSubsetting !==
      checked.font["OS/2"].fsType.noSubsetting ||
    font["OS/2"].fsSelection.bold !== checked.font["OS/2"].fsSelection.bold ||
    font["OS/2"].fsSelection.italic !== checked.font["OS/2"].fsSelection.italic
  ) {
    return fail(
      use,
      "font-subsetting-failed",
      `${use.family} 的字符子集未保留原字体的嵌入许可和样式`,
    );
  }
  for (const character of requiredCharacters) {
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      requiresFontGlyph(character) &&
      !font.hasGlyphForCodePoint(codePoint)
    ) {
      return fail(
        use,
        "font-subsetting-failed",
        `${use.family} 的字符子集未保留实际输出字符 ${character}`,
      );
    }
  }
  return bytes;
}

function nameString(
  bytes: Uint8Array,
  table: FontTable,
  id: number,
  use: FontEmbeddingUse,
): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (table.length < 6) {
    return fail(use, "unsupported-font-format", "字体 name 表不完整");
  }
  const count = view.getUint16(table.offset + 2);
  const stringsOffset = view.getUint16(table.offset + 4);
  if (6 + count * 12 > table.length || stringsOffset > table.length) {
    return fail(use, "unsupported-font-format", "字体 name 表记录越界");
  }
  let fallback: Uint8Array | undefined;
  for (let index = 0; index < count; index += 1) {
    const record = table.offset + 6 + index * 12;
    if (view.getUint16(record + 6) !== id) continue;
    const platform = view.getUint16(record);
    const language = view.getUint16(record + 4);
    const length = view.getUint16(record + 8);
    const offset = stringsOffset + view.getUint16(record + 10);
    if (offset > table.length || length > table.length - offset) {
      return fail(use, "unsupported-font-format", "字体 name 字符串越界");
    }
    if (platform !== 0 && platform !== 3) continue;
    const raw = bytes.subarray(
      table.offset + offset,
      table.offset + offset + length,
    );
    if (raw.byteLength % 2 !== 0) {
      return fail(
        use,
        "unsupported-font-format",
        "字体 name 字符串不是 UTF-16 编码",
      );
    }
    const littleEndian = new Uint8Array(raw.byteLength);
    for (let at = 0; at < raw.byteLength; at += 2) {
      littleEndian[at] = raw[at + 1]!;
      littleEndian[at + 1] = raw[at]!;
    }
    if (platform === 3 && language === 0x0409) return littleEndian;
    fallback ??= littleEndian;
  }
  if (fallback === undefined) {
    return fail(
      use,
      "unsupported-font-format",
      `字体 name 表缺少 EOT 所需的 Unicode 名称字段 ${id}`,
    );
  }
  return fallback;
}

function eot(
  sourceBytes: Uint8Array,
  use: FontEmbeddingUse,
  checked: ReturnType<typeof checkedFont>,
  payloadBytes: Uint8Array = sourceBytes,
): Uint8Array {
  const { tables, fsType, weightClass, font } = checked;
  const os2 = tables.get("OS/2")!;
  const names = tables.get("name")!;
  const view = new DataView(
    sourceBytes.buffer,
    sourceBytes.byteOffset,
    sourceBytes.byteLength,
  );
  const payloadTables = tablesOf(payloadBytes, use);
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
  u8(font["OS/2"].fsSelection.italic ? 1 : 0);
  u32(weightClass);
  u16(fsType);
  u16(0x504c);
  for (const offset of [42, 46, 50, 54]) u32(field32(offset));
  for (const offset of [78, 82]) {
    u32(os2Version >= 1 && os2.length >= offset + 4 ? field32(offset) : 0);
  }
  const payloadHead = payloadTables.get("head")!;
  u32(payloadView.getUint32(payloadHead.offset + 8));
  for (let index = 0; index < 4; index += 1) u32(0);
  for (const id of [1, 2, 5, 4]) {
    const name = nameString(sourceBytes, names, id, use);
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

function fontCategory(font: fontkit.Font): {
  readonly pitchFamily: number;
  readonly charset: number;
} {
  const panose = font["OS/2"].panose;
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

type FontSlot = "regular" | "bold" | "italic" | "boldItalic";

function slot(bold: boolean, italic: boolean): FontSlot {
  if (bold) return italic ? "boldItalic" : "bold";
  return italic ? "italic" : "regular";
}

function slotOrder(value: FontSlot): number {
  return { regular: 0, bold: 1, italic: 2, boldItalic: 3 }[value];
}

/** 按实际输出字族和样式生成完整字体或文稿字符子集。 */
export function fontPresentationContributions(
  uses: readonly FontEmbeddingUse[],
  requested: FontEmbeddingMode | undefined,
): CapabilityPresentationContribution[] {
  if (requested === undefined) return [];
  const contributions: CapabilityPresentationContribution[] = [];
  const slots = new Map<string, string>();
  for (const use of uses) {
    const checked = checkedFont(use, requested);
    const styleSlot = slot(use.bold, use.italic);
    const family = xmlEscape(use.family);
    const category = fontCategory(checked.font);
    const recordContent = `<p:font typeface="${family}" pitchFamily="${category.pitchFamily}" charset="${category.charset}"/>`;
    const recordKey = fontFamilyKey(use.family);
    const recordIdentity = `${recordContent}\u0000${category.pitchFamily}\u0000${category.charset}`;
    const itemKey = `${recordKey}\u0000${styleSlot}`;
    const existing = slots.get(itemKey);
    if (existing !== undefined && existing !== use.identity) {
      return fail(
        use,
        "conflicting-resource",
        `${use.family} 的 ${styleSlot} 样式对应多个字体文件`,
      );
    }
    if (existing !== undefined) continue;
    slots.set(itemKey, use.identity);
    const resourceClaim = { kind: "font-face", key: use.claimKey };
    const resolveBytes =
      requested === "subset"
        ? (characters: string) =>
            eot(use.bytes!, use, checked, subsetBytes(use, checked, characters))
        : undefined;
    contributions.push({
      collection: "p:embeddedFontLst",
      presentationAttributes: { embedTrueTypeFonts: "1" },
      record: {
        key: recordKey,
        identity: recordIdentity,
        tagName: "p:embeddedFont",
        content: recordContent,
      },
      item: {
        key: styleSlot,
        identity: use.identity,
        order: slotOrder(styleSlot),
        relationshipType: fontRelationship,
        pathTemplate: "ppt/fonts/font{n}.fntdata",
        contentType: fontContentType,
        ...(resolveBytes === undefined
          ? { bytes: eot(use.bytes!, use, checked) }
          : { resourceClaim, resolveBytes }),
        content: (relationshipId) =>
          `<p:${styleSlot} r:id="${xmlEscape(relationshipId)}"/>`,
      },
    });
  }
  return contributions;
}
