import {
  defaultTheme,
  defineTheme,
  type Font,
  type FontFace,
} from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import {
  P_NS,
  readEmbeddedFontList,
  readPresentationThemeFontScheme,
  unpackEotFontData,
  type DeckImportContext,
  type DeckImportHandler,
  type DeckImportOutcome,
  type EmbeddedFontSlot,
  type ImportDiagnostic,
} from "@pptxdk/import";
import * as fontkit from "fontkit";

import { fontFamilyKey, matchesFontFamily } from "./font-family.ts";
import { font } from "./font.ts";

function fallbackWeightForSlot(slot: EmbeddedFontSlot): number {
  return slot === "bold" || slot === "boldItalic" ? 700 : 400;
}

function fallbackStyleForSlot(slot: EmbeddedFontSlot): "normal" | "italic" {
  return slot === "italic" || slot === "boldItalic" ? "italic" : "normal";
}

function attachEmbeddedFontsToSubtree(
  node: CommonNode,
  fontsByKey: ReadonlyMap<string, Font>,
): void {
  const style = node.props["style"];
  if (typeof style === "object" && style !== null && !Array.isArray(style)) {
    const family = (style as Readonly<Record<string, unknown>>)["fontFamily"];
    if (
      typeof family === "string" &&
      family !== "major" &&
      family !== "minor"
    ) {
      const matched = fontsByKey.get(fontFamilyKey(family));
      if (matched !== undefined) {
        (
          node as CommonNode<
            string,
            Record<string, unknown>,
            Record<string, unknown>
          >
        ).style({
          fontFamily: matched,
        });
      }
    }
  }
  for (const child of node.children) {
    if (child instanceof CommonNode) {
      attachEmbeddedFontsToSubtree(child, fontsByKey);
    }
  }
}

/**
 * Reads `<p:embeddedFontLst>` from `presentation.xml`, unpacks uncompressed EOT font parts
 * into detached `Font` resources, attaches them to theme font slots and text nodes, and
 * isolates compressed, obfuscated, or corrupt `.fntdata` gracefully.
 */
export function readEmbeddedFontsIntoDeck(
  context: DeckImportContext,
): DeckImportOutcome {
  const { presentationCursor, deckNode, package: pkg } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const records = readEmbeddedFontList(
    presentationCursor,
    pkg.presentationRels,
    pkg,
    diagnostics,
  );

  const fontsByKey = new Map<string, Font>();
  for (const record of records) {
    const faces: FontFace[] = [];
    let requiresSubstitute = false;

    for (const slot of record.slots) {
      const unpacked = unpackEotFontData(slot.bytes);
      if (unpacked.status === "unsupported") {
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: { part: slot.partPath, xpath: slot.xpath },
          namespaceUri: P_NS,
          localName: slot.slot,
          mechanismHint: unpacked.reason,
        });
        continue;
      }
      if (unpacked.status === "corrupt") {
        diagnostics.push({
          code: "corrupt-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: slot.partPath, xpath: slot.xpath },
        });
        continue;
      }

      let opened: fontkit.Font | fontkit.FontCollection;
      try {
        opened = fontkit.create(
          Buffer.from(
            unpacked.ttfBytes.buffer,
            unpacked.ttfBytes.byteOffset,
            unpacked.ttfBytes.byteLength,
          ),
        );
      } catch {
        diagnostics.push({
          code: "corrupt-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: slot.partPath, xpath: slot.xpath },
        });
        continue;
      }

      if (opened.type !== "TTF") {
        diagnostics.push({
          code: "corrupt-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: slot.partPath, xpath: slot.xpath },
        });
        continue;
      }

      if (!matchesFontFamily(opened, record.typeface)) {
        requiresSubstitute = true;
      }

      const rawWeight = opened["OS/2"]?.usWeightClass;
      const weight =
        typeof rawWeight === "number" &&
        rawWeight >= 100 &&
        rawWeight <= 900 &&
        rawWeight % 100 === 0
          ? rawWeight
          : fallbackWeightForSlot(slot.slot);
      const isItalic =
        Boolean(opened["OS/2"]?.fsSelection?.italic) ||
        Boolean(opened["OS/2"]?.fsSelection?.oblique) ||
        opened.italicAngle !== 0;
      const style: "normal" | "italic" =
        opened["OS/2"] !== undefined
          ? isItalic
            ? "italic"
            : "normal"
          : fallbackStyleForSlot(slot.slot);

      faces.push({
        weight,
        style,
        data: new Uint8Array(unpacked.ttfBytes),
      });
    }

    if (faces.length > 0) {
      fontsByKey.set(
        fontFamilyKey(record.typeface),
        font(
          record.typeface,
          faces,
          requiresSubstitute ? { substitute: true } : {},
        ),
      );
    }
  }

  const scheme = readPresentationThemeFontScheme(pkg);
  if (scheme !== undefined) {
    const resolvedMajor =
      scheme.majorTypeface === undefined
        ? undefined
        : (fontsByKey.get(fontFamilyKey(scheme.majorTypeface)) ??
          scheme.majorTypeface);
    const resolvedMinor =
      scheme.minorTypeface === undefined
        ? undefined
        : (fontsByKey.get(fontFamilyKey(scheme.minorTypeface)) ??
          scheme.minorTypeface);

    const hasCustomMajor =
      resolvedMajor !== undefined && resolvedMajor !== defaultTheme.fonts.major;
    const hasCustomMinor =
      resolvedMinor !== undefined && resolvedMinor !== defaultTheme.fonts.minor;

    if (hasCustomMajor || hasCustomMinor) {
      deckNode.set({
        theme: defineTheme({
          ...(scheme.themeName === undefined ? {} : { name: scheme.themeName }),
          fonts: {
            ...(resolvedMajor === undefined ? {} : { major: resolvedMajor }),
            ...(resolvedMinor === undefined ? {} : { minor: resolvedMinor }),
          },
        }),
      });
    }
  }

  if (fontsByKey.size > 0) {
    attachEmbeddedFontsToSubtree(deckNode, fontsByKey);
  }

  return { diagnostics };
}

/** Deck-level import handler wired into `embeddedFontMechanism`. */
export const embeddedFontImportHandler: DeckImportHandler = {
  target: "deck",
  read: readEmbeddedFontsIntoDeck,
};
