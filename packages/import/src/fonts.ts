import { z } from "zod";

import type { ImportDiagnostic } from "#/diagnostics.ts";
import {
  A_NS,
  P_NS,
  REL_FONT,
  REL_SLIDE_MASTER,
  REL_THEME,
  R_NS,
} from "#/namespaces.ts";
import {
  resolvePartPath,
  type ParsedRelationshipEntry,
  type PptxPackage,
} from "#/opc.ts";
import { OoxmlBoolSchema, OoxmlIntSchema } from "#/xml-fragments.ts";
import { XmlElementCursor, parseXmlPart } from "#/xml-reader.ts";

export type EmbeddedFontSlot = "regular" | "bold" | "italic" | "boldItalic";

export type EmbeddedFontSlotRecord = {
  readonly slot: EmbeddedFontSlot;
  readonly relationshipId: string;
  readonly partPath: string;
  readonly bytes: Uint8Array;
  readonly xpath: string;
};

export type EmbeddedFontRecord = {
  readonly typeface: string;
  readonly pitchFamily?: number;
  readonly charset?: number;
  readonly slots: readonly EmbeddedFontSlotRecord[];
  readonly xpath: string;
};

export type PresentationFontSchemeInfo = {
  readonly themeName?: string;
  readonly majorTypeface?: string;
  readonly minorTypeface?: string;
};

export type UnpackedEotResult =
  | {
      readonly status: "ok";
      readonly ttfBytes: Uint8Array;
      readonly eotWeight?: number;
      readonly eotItalic?: boolean;
    }
  | {
      readonly status: "unsupported";
      readonly reason: "compressed-eot" | "obfuscated-eot" | "eudc-eot";
    }
  | {
      readonly status: "corrupt";
      readonly reason: string;
    };

const fontSlotNames: readonly EmbeddedFontSlot[] = [
  "regular",
  "bold",
  "italic",
  "boldItalic",
];

const PresentationFontAttrSchema = z.object({
  embedTrueTypeFonts: OoxmlBoolSchema.optional(),
  saveSubsetFonts: OoxmlBoolSchema.optional(),
});

const EmbeddedFontTypefaceXmlSchema = z.object({
  typeface: z.string().min(1),
  pitchFamily: OoxmlIntSchema.optional(),
  charset: OoxmlIntSchema.optional(),
  panose: z.string().optional(),
});

const EmbeddedFontSlotRefXmlSchema = z.object({
  id: z.string().min(1),
});

const ThemeTypefaceXmlSchema = z.object({
  typeface: z.string().min(1),
});

function isTrueTypeSfntHeader(bytes: Uint8Array): boolean {
  if (bytes.byteLength < 12) return false;
  const sig = String.fromCharCode(bytes[0]!, bytes[1]!, bytes[2]!, bytes[3]!);
  return sig === "\0\x01\0\0" || sig === "true";
}

/**
 * Unpacks an uncompressed EOT (`.fntdata`) container or raw SFNT buffer into a detached
 * TrueType `Uint8Array`, isolating compressed/obfuscated or corrupt payloads cleanly.
 */
export function unpackEotFontData(bytes: Uint8Array): UnpackedEotResult {
  if (isTrueTypeSfntHeader(bytes)) {
    return { status: "ok", ttfBytes: new Uint8Array(bytes) };
  }
  if (bytes.byteLength < 82) {
    return { status: "corrupt", reason: "EOT header shorter than 82 bytes" };
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eotSize = view.getUint32(0, true);
  const fontDataSize = view.getUint32(4, true);
  const version = view.getUint32(8, true);
  const flags = view.getUint32(12, true);
  const italic = bytes[27]! !== 0;
  const weight = view.getUint32(28, true);
  const magic = view.getUint16(34, true);

  if (magic !== 0x504c) {
    return {
      status: "corrupt",
      reason: `Invalid EOT magic number 0x${magic.toString(16)}`,
    };
  }
  if (
    version !== 0x00010000 &&
    version !== 0x00020001 &&
    version !== 0x00020002
  ) {
    return {
      status: "corrupt",
      reason: `Unsupported EOT version 0x${version.toString(16)}`,
    };
  }
  if ((flags & 0x00000004) !== 0) {
    return { status: "unsupported", reason: "compressed-eot" };
  }
  if ((flags & 0x10000000) !== 0) {
    return { status: "unsupported", reason: "obfuscated-eot" };
  }
  if ((flags & 0x00000001) !== 0) {
    return { status: "unsupported", reason: "eudc-eot" };
  }
  if (
    fontDataSize === 0 ||
    eotSize > bytes.byteLength ||
    fontDataSize > eotSize - 82
  ) {
    return {
      status: "corrupt",
      reason: "EOT font payload bounds exceed container length",
    };
  }

  const payloadStart = eotSize - fontDataSize;
  const ttfBytes = bytes.slice(payloadStart, eotSize);
  if (!isTrueTypeSfntHeader(ttfBytes)) {
    return {
      status: "corrupt",
      reason: "Unpacked EOT payload is not a static TrueType SFNT file",
    };
  }
  return {
    status: "ok",
    ttfBytes,
    ...(weight > 0 ? { eotWeight: weight } : {}),
    eotItalic: italic,
  };
}

/**
 * Reads presentation-level `<p:embeddedFontLst>` declarations and resolves each font slot
 * relationship into detached `.fntdata` bytes while isolating missing or broken targets.
 */
export function readEmbeddedFontList(
  presentationCursor: XmlElementCursor,
  presentationRels: readonly ParsedRelationshipEntry[],
  pkg: PptxPackage,
  diagnostics: ImportDiagnostic[],
): readonly EmbeddedFontRecord[] {
  const rawPresAttrs = presentationCursor.attrsObject([
    "embedTrueTypeFonts",
    "saveSubsetFonts",
  ]);
  const parsedPresAttrs = PresentationFontAttrSchema.safeParse(rawPresAttrs);
  if (!parsedPresAttrs.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: {
        part: pkg.presentationPath,
        xpath: presentationCursor.xpath,
      },
      attributeOrElement: "p:presentation",
      rawValue: JSON.stringify(rawPresAttrs),
    });
  }

  const embeddedFontLst = presentationCursor.child(P_NS, "embeddedFontLst");
  if (embeddedFontLst === undefined) return [];

  const relById = new Map<string, ParsedRelationshipEntry>();
  for (const rel of presentationRels) {
    relById.set(rel.Id, rel);
  }

  const records: EmbeddedFontRecord[] = [];
  for (const entryCursor of embeddedFontLst.children(P_NS, "embeddedFont")) {
    const fontEl = entryCursor.child(P_NS, "font");
    if (fontEl === undefined) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: entryCursor.xpath },
      });
      continue;
    }

    const rawFont = fontEl.attrsObject([
      "typeface",
      "pitchFamily",
      "charset",
      "panose",
    ]);
    const parsedFont = EmbeddedFontTypefaceXmlSchema.safeParse(rawFont);
    if (!parsedFont.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: fontEl.xpath },
        attributeOrElement: "p:font",
        rawValue: JSON.stringify(rawFont),
      });
      continue;
    }

    const slots: EmbeddedFontSlotRecord[] = [];
    for (const slotName of fontSlotNames) {
      const slotCursor = entryCursor.child(P_NS, slotName);
      if (slotCursor === undefined) continue;
      const rId = slotCursor.attr("id", R_NS);
      const parsedRef = EmbeddedFontSlotRefXmlSchema.safeParse({ id: rId });
      if (!parsedRef.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: pkg.presentationPath, xpath: slotCursor.xpath },
          attributeOrElement: `p:${slotName}/@r:id`,
          rawValue: rId ?? "",
        });
        continue;
      }

      const rel = relById.get(parsedRef.data.id);
      if (
        rel === undefined ||
        rel.Type !== REL_FONT ||
        rel.TargetMode === "External"
      ) {
        diagnostics.push({
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: pkg.presentationPath, xpath: slotCursor.xpath },
          relationshipId: parsedRef.data.id,
        });
        continue;
      }

      const targetPart = resolvePartPath(pkg.presentationPath, rel.Target);
      const bytes =
        targetPart === undefined
          ? undefined
          : pkg.readDetachedBytes(targetPart);
      if (targetPart === undefined || bytes === undefined) {
        diagnostics.push({
          code: "missing-target-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: pkg.presentationPath, xpath: slotCursor.xpath },
          relationshipId: parsedRef.data.id,
          targetPart: targetPart ?? rel.Target,
        });
        continue;
      }

      slots.push({
        slot: slotName,
        relationshipId: parsedRef.data.id,
        partPath: targetPart,
        bytes,
        xpath: slotCursor.xpath,
      });
    }

    records.push({
      typeface: parsedFont.data.typeface,
      ...(parsedFont.data.pitchFamily === undefined
        ? {}
        : { pitchFamily: parsedFont.data.pitchFamily }),
      ...(parsedFont.data.charset === undefined
        ? {}
        : { charset: parsedFont.data.charset }),
      slots,
      xpath: entryCursor.xpath,
    });
  }

  return records;
}

/**
 * Reads the primary theme `<a:fontScheme>` (`majorFont`/`minorFont` Latin typeface) following
 * presentation -> slideMaster -> theme relationships so `+mj-lt` / `+mn-lt` tokens can link
 * to embedded font resources.
 */
export function readPresentationThemeFontScheme(
  pkg: PptxPackage,
): PresentationFontSchemeInfo | undefined {
  const masterRel = pkg.presentationRels.find(
    (rel) => rel.Type === REL_SLIDE_MASTER && rel.TargetMode !== "External",
  );
  if (masterRel === undefined) return undefined;
  const masterPath = resolvePartPath(pkg.presentationPath, masterRel.Target);
  if (masterPath === undefined) return undefined;

  const masterRels = pkg.readRelationships(masterPath).entries;
  const themeRel = masterRels.find(
    (rel) => rel.Type === REL_THEME && rel.TargetMode !== "External",
  );
  if (themeRel === undefined) return undefined;
  const themePath = resolvePartPath(masterPath, themeRel.Target);
  if (themePath === undefined) return undefined;

  const themeXmlText = pkg.readTextPart(themePath);
  if (themeXmlText === undefined) return undefined;

  let themeEl: ReturnType<typeof parseXmlPart>;
  try {
    themeEl = parseXmlPart(themePath, themeXmlText);
  } catch {
    return undefined;
  }
  if (
    (themeEl.namespaceURI ?? "") !== A_NS ||
    (themeEl.localName ?? themeEl.tagName) !== "theme"
  ) {
    return undefined;
  }

  const themeCursor = XmlElementCursor.createRoot(themeEl, themePath);
  const themeName = themeCursor.attr("name");
  const themeElements = themeCursor.child(A_NS, "themeElements");
  const fontScheme = themeElements?.child(A_NS, "fontScheme");
  if (fontScheme === undefined) return undefined;

  const majorLatin = fontScheme
    .child(A_NS, "majorFont")
    ?.child(A_NS, "latin")
    ?.attrsObject(["typeface"]);
  const minorLatin = fontScheme
    .child(A_NS, "minorFont")
    ?.child(A_NS, "latin")
    ?.attrsObject(["typeface"]);

  const parsedMajor = ThemeTypefaceXmlSchema.safeParse(majorLatin);
  const parsedMinor = ThemeTypefaceXmlSchema.safeParse(minorLatin);

  return {
    ...(themeName !== undefined && themeName.trim() !== ""
      ? { themeName }
      : {}),
    ...(parsedMajor.success
      ? { majorTypeface: parsedMajor.data.typeface }
      : {}),
    ...(parsedMinor.success
      ? { minorTypeface: parsedMinor.data.typeface }
      : {}),
  };
}
