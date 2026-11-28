import {
  CapabilityError,
  type CapabilityPresentationContribution,
  type FontEmbeddingMode,
} from "@pptxdk/core/capability";
import { xmlEscape } from "@pptxdk/emit";

import {
  buildEmbeddedEotPayload,
  fontFamilyKey,
  validateEmbeddableTrueTypeFont,
} from "./font-engine.ts";
import type { FontFaceStyle } from "./font.ts";

/** Retains control characters and default ignorable code points in text stream without requiring standalone cmap glyphs. */
export function requiresFontGlyph(character: string): boolean {
  return !/^[\p{Cc}\p{Default_Ignorable_Code_Point}]$/u.test(character);
}

/** Checks whether a character is a visible non-whitespace character requiring a non-.notdef glyph. */
export function requiresVisibleFontGlyph(character: string): boolean {
  return !/^[\p{Cc}\p{Default_Ignorable_Code_Point}\p{White_Space}]$/u.test(
    character,
  );
}

const fontRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const fontContentType = "application/x-fontdata";
export type { FontEmbeddingMode };

/** Final font file and style chosen during text serialization; missing files are preserved as actual font requirements. */
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

function fail(
  use: FontEmbeddingUse,
  code: string,
  message: string,
  extraDetails?: Readonly<Record<string, unknown>>,
): never {
  throw new CapabilityError(code, message, use.path, {
    font: {
      family: use.family,
      weight: use.weight,
      style: use.style,
    },
    ...extraDetails,
  });
}

type FontSlot = "regular" | "bold" | "italic" | "boldItalic";

function slot(bold: boolean, italic: boolean): FontSlot {
  if (bold) return italic ? "boldItalic" : "bold";
  return italic ? "italic" : "regular";
}

function slotOrder(value: FontSlot): number {
  return { regular: 0, bold: 1, italic: 2, boldItalic: 3 }[value];
}

function formatEmbeddedFontSlotConflictMessage(
  family: string,
  styleSlot: FontSlot,
  nativeLegacyFamily: string | undefined,
): string {
  if (
    nativeLegacyFamily !== undefined &&
    fontFamilyKey(nativeLegacyFamily) !== fontFamilyKey(family)
  ) {
    return `Font family "${family}" maps multiple font files to the "${styleSlot}" embedding slot (conflicting file native Name ID 1: "${nativeLegacyFamily}"). OOXML (<p:embeddedFont>) supports only four slots per font family (regular, bold, italic, boldItalic); register "${nativeLegacyFamily}" as a separate font family.`;
  }
  const nativePart =
    nativeLegacyFamily === undefined
      ? ""
      : ` (file native Name ID 1: "${nativeLegacyFamily}")`;
  return `Font family "${family}" maps multiple font files to the "${styleSlot}" embedding slot${nativePart}. OOXML (<p:embeddedFont>) supports only four slots per font family (regular, bold, italic, boldItalic); register additional weights (such as 500 Medium or 600 SemiBold) under a separate font family name (for example "${family} Medium").`;
}

function embeddedFontSlotConflictDetails(
  use: FontEmbeddingUse,
): Readonly<Record<string, unknown>> {
  return {
    font: {
      family: use.family,
      weight: use.weight,
      style: use.style,
    },
    reason: "not-embeddable",
  };
}

/** Generates full fonts or presentation character subsets according to emitted families and styles. */
export function fontPresentationContributions(
  uses: readonly FontEmbeddingUse[],
  requested: FontEmbeddingMode | undefined,
): CapabilityPresentationContribution[] {
  if (requested === undefined) return [];
  const contributions: CapabilityPresentationContribution[] = [];
  const slots = new Map<
    string,
    {
      readonly identity: string;
      readonly primaryLegacyFamily: string | undefined;
    }
  >();
  const recordContents = new Map<string, string>();
  for (const use of uses) {
    const checked = validateEmbeddableTrueTypeFont(use, requested);
    const styleSlot = slot(checked.os2Bold, checked.isItalic);
    const family = xmlEscape(use.family);
    const recordKey = fontFamilyKey(use.family);
    // Keep one <p:font> header per case-insensitive family so PANOSE or casing differences across weights do not fail record registration.
    let recordContent = recordContents.get(recordKey);
    if (recordContent === undefined) {
      const panoseAttr =
        checked.panose === undefined ? "" : ` panose="${checked.panose}"`;
      recordContent = `<p:font typeface="${family}"${panoseAttr} pitchFamily="${checked.pitchFamily}" charset="${checked.charset}"/>`;
      recordContents.set(recordKey, recordContent);
    }
    const recordIdentity = recordKey;
    const itemKey = `${recordKey}\u0000${styleSlot}`;
    const existing = slots.get(itemKey);
    const conflictDetails = embeddedFontSlotConflictDetails(use);
    if (existing !== undefined && existing.identity !== use.identity) {
      const conflictNative =
        checked.primaryLegacyFamily !== undefined &&
        fontFamilyKey(checked.primaryLegacyFamily) !== fontFamilyKey(use.family)
          ? checked.primaryLegacyFamily
          : (existing.primaryLegacyFamily ?? checked.primaryLegacyFamily);
      return fail(
        use,
        "font-unavailable",
        formatEmbeddedFontSlotConflictMessage(
          use.family,
          styleSlot,
          conflictNative,
        ),
        {
          reason: "not-embeddable",
        },
      );
    }
    if (existing !== undefined) continue;
    slots.set(itemKey, {
      identity: use.identity,
      primaryLegacyFamily: checked.primaryLegacyFamily,
    });
    const conflictMessage = formatEmbeddedFontSlotConflictMessage(
      use.family,
      styleSlot,
      checked.primaryLegacyFamily,
    );
    const resourceClaim = { kind: "font-face", key: use.claimKey };
    const resolveBytes =
      requested === "subset"
        ? (characters: string) =>
            buildEmbeddedEotPayload(
              use.bytes!,
              checked,
              "subset",
              characters,
              use,
              requiresVisibleFontGlyph,
            )
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
          ? {
              bytes: buildEmbeddedEotPayload(
                use.bytes!,
                checked,
                "full",
                undefined,
                use,
                requiresVisibleFontGlyph,
              ),
            }
          : { resourceClaim, resolveBytes }),
        content: (relationshipId) =>
          `<p:${styleSlot} r:id="${xmlEscape(relationshipId)}"/>`,
        onConflict: {
          code: "font-unavailable",
          message: conflictMessage,
          details: conflictDetails,
        },
      },
    });
  }
  return contributions;
}
