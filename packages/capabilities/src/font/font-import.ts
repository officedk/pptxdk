import {
  defaultTheme,
  defineTheme,
  isTheme,
  type Font,
  type FontFace,
  type PlaceholderListStyles,
  type PlaceholderMap,
  type Theme,
  type ThemeInput,
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

function upgradeFontToken(
  family: unknown,
  fontsByKey: ReadonlyMap<string, Font>,
): Font | undefined {
  if (typeof family !== "string" || family === "major" || family === "minor") {
    return undefined;
  }
  return fontsByKey.get(fontFamilyKey(family));
}

function upgradeStyleRecordFont(
  style: unknown,
  fontsByKey: ReadonlyMap<string, Font>,
): {
  readonly value?: Readonly<Record<string, unknown>>;
  readonly changed: boolean;
} {
  if (typeof style !== "object" || style === null || Array.isArray(style)) {
    return { changed: false };
  }
  const rec = style as Readonly<Record<string, unknown>>;
  const matched = upgradeFontToken(rec["fontFamily"], fontsByKey);
  if (matched === undefined) return { changed: false };
  return { changed: true, value: { ...rec, fontFamily: matched } };
}

function upgradePlaceholderListStylesFont(
  listStyles: PlaceholderListStyles | undefined,
  fontsByKey: ReadonlyMap<string, Font>,
): {
  readonly value: PlaceholderListStyles | undefined;
  readonly changed: boolean;
} {
  if (listStyles === undefined) return { value: undefined, changed: false };
  let changed = false;
  const next: Record<string, unknown> = {};
  for (const [levelKey, levelStyle] of Object.entries(listStyles)) {
    if (levelStyle === undefined) continue;
    const upgraded = upgradeStyleRecordFont(levelStyle, fontsByKey);
    if (upgraded.changed && upgraded.value !== undefined) {
      next[levelKey] = upgraded.value;
      changed = true;
    } else {
      next[levelKey] = levelStyle;
    }
  }
  return changed
    ? { value: next, changed: true }
    : { value: listStyles, changed: false };
}

function upgradePlaceholdersFont(
  placeholders: PlaceholderMap | undefined,
  fontsByKey: ReadonlyMap<string, Font>,
): {
  readonly value: PlaceholderMap | undefined;
  readonly changed: boolean;
} {
  if (placeholders === undefined) return { value: undefined, changed: false };
  let changed = false;
  const next: Record<string, PlaceholderMap[string]> = {};
  for (const [key, spec] of Object.entries(placeholders)) {
    const styleUp = upgradeStyleRecordFont(spec.style, fontsByKey);
    const listUp = upgradePlaceholderListStylesFont(
      spec.listStyles,
      fontsByKey,
    );
    if (styleUp.changed || listUp.changed) {
      changed = true;
      next[key] = {
        ...spec,
        ...(styleUp.changed && styleUp.value !== undefined
          ? { style: styleUp.value }
          : {}),
        ...(listUp.changed && listUp.value !== undefined
          ? { listStyles: listUp.value }
          : {}),
      };
    } else {
      next[key] = spec;
    }
  }
  return changed
    ? { value: next, changed: true }
    : { value: placeholders, changed: false };
}

function attachEmbeddedFontsToSubtree(
  node: CommonNode,
  fontsByKey: ReadonlyMap<string, Font>,
): void {
  const mutableNode = node as CommonNode<
    string,
    Record<string, unknown>,
    Record<string, unknown>
  >;
  const styleUp = upgradeStyleRecordFont(node.props["style"], fontsByKey);
  if (styleUp.changed && styleUp.value !== undefined) {
    mutableNode.style({ fontFamily: styleUp.value["fontFamily"] });
  }
  if (node.type === "slide-master" || node.type === "slide-layout") {
    const rawPlaceholders = node.props["placeholders"] as
      | PlaceholderMap
      | undefined;
    const upgradedPh = upgradePlaceholdersFont(rawPlaceholders, fontsByKey);
    if (upgradedPh.changed && upgradedPh.value !== undefined) {
      mutableNode.set({ placeholders: upgradedPh.value });
    }
    if (node.type === "slide-master") {
      const rawTextStyles = node.props["textStyles"] as
        | {
            readonly titleStyle?: PlaceholderListStyles;
            readonly bodyStyle?: PlaceholderListStyles;
            readonly otherStyle?: PlaceholderListStyles;
          }
        | undefined;
      if (rawTextStyles !== undefined) {
        const titleUp = upgradePlaceholderListStylesFont(
          rawTextStyles.titleStyle,
          fontsByKey,
        );
        const bodyUp = upgradePlaceholderListStylesFont(
          rawTextStyles.bodyStyle,
          fontsByKey,
        );
        const otherUp = upgradePlaceholderListStylesFont(
          rawTextStyles.otherStyle,
          fontsByKey,
        );
        if (titleUp.changed || bodyUp.changed || otherUp.changed) {
          mutableNode.set({
            textStyles: {
              ...(titleUp.value === undefined
                ? {}
                : { titleStyle: titleUp.value }),
              ...(bodyUp.value === undefined
                ? {}
                : { bodyStyle: bodyUp.value }),
              ...(otherUp.value === undefined
                ? {}
                : { otherStyle: otherUp.value }),
            },
          });
        }
      }
    }
  }
  const rawFooterPlaceholders = node.props["footerPlaceholders"];
  if (
    typeof rawFooterPlaceholders === "object" &&
    rawFooterPlaceholders !== null &&
    !Array.isArray(rawFooterPlaceholders)
  ) {
    let changed = false;
    const nextPlaceholders: Record<string, unknown> = {
      ...(rawFooterPlaceholders as Readonly<Record<string, unknown>>),
    };
    for (const [slot, spec] of Object.entries(nextPlaceholders)) {
      if (typeof spec !== "object" || spec === null || Array.isArray(spec)) {
        continue;
      }
      const nextStyle = upgradeStyleRecordFont(
        (spec as Readonly<Record<string, unknown>>)["style"],
        fontsByKey,
      );
      if (nextStyle.changed && nextStyle.value !== undefined) {
        nextPlaceholders[slot] = {
          ...(spec as Readonly<Record<string, unknown>>),
          style: nextStyle.value,
        };
        changed = true;
      }
    }
    if (changed) {
      mutableNode.set({ footerPlaceholders: nextPlaceholders });
    }
  }
  for (const slideSlot of ["dateTime", "footer", "slideNumber"] as const) {
    const rawSlot = node.props[slideSlot];
    if (
      typeof rawSlot !== "object" ||
      rawSlot === null ||
      Array.isArray(rawSlot)
    ) {
      continue;
    }
    const nextStyle = upgradeStyleRecordFont(
      (rawSlot as Readonly<Record<string, unknown>>)["style"],
      fontsByKey,
    );
    if (nextStyle.changed && nextStyle.value !== undefined) {
      mutableNode.set({
        [slideSlot]: {
          ...(rawSlot as Readonly<Record<string, unknown>>),
          style: nextStyle.value,
        },
      });
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
  const attachFontsToThemeInput = (
    existing: ThemeInput | undefined,
  ): {
    readonly theme?: ReturnType<typeof defineTheme>;
    readonly changed: boolean;
  } => {
    if (existing === undefined) return { changed: false };
    const existingFonts = existing.fonts ?? {};
    const nextFonts: Record<string, Font | string> = {};
    let changed = false;
    for (const [slot, token] of Object.entries(existingFonts)) {
      if (token === undefined) continue;
      if (typeof token === "string") {
        const matched = fontsByKey.get(fontFamilyKey(token));
        if (matched !== undefined) {
          nextFonts[slot] = matched;
          changed = true;
          continue;
        }
      }
      nextFonts[slot] = token;
    }
    if (!changed) return { changed: false };
    return {
      changed: true,
      theme: defineTheme({
        ...existing,
        fonts: nextFonts,
      }),
    };
  };

  let masterFallbackTheme: ThemeInput | undefined;
  const mutatedThemes = new Set<Theme>();
  for (const child of deckNode.children) {
    if (child instanceof CommonNode && child.type === "slide-master") {
      const masterTheme = child.props["theme"] as
        | Theme
        | ThemeInput
        | undefined;
      masterFallbackTheme ??= masterTheme;
      if (isTheme(masterTheme) && masterTheme !== defaultTheme) {
        if (!mutatedThemes.has(masterTheme)) {
          mutatedThemes.add(masterTheme);
          const updatedMasterTheme = attachFontsToThemeInput(masterTheme);
          if (
            updatedMasterTheme.changed &&
            updatedMasterTheme.theme !== undefined
          ) {
            masterTheme.set({ fonts: updatedMasterTheme.theme.fonts });
          }
        }
      } else {
        const updatedMasterTheme = attachFontsToThemeInput(masterTheme);
        if (
          updatedMasterTheme.changed &&
          updatedMasterTheme.theme !== undefined
        ) {
          child.set({ theme: updatedMasterTheme.theme });
        }
      }
    }
  }

  const existingDeckTheme = deckNode.props["theme"] as ThemeInput | undefined;
  const baseDeckTheme = existingDeckTheme ?? masterFallbackTheme;

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
      const upgradedBase =
        attachFontsToThemeInput(baseDeckTheme).theme ?? baseDeckTheme;
      const themeName = scheme.themeName ?? upgradedBase?.name;
      const nextFonts: Record<string, Font | string> = {};
      if (upgradedBase?.fonts !== undefined) {
        for (const [slot, token] of Object.entries(upgradedBase.fonts)) {
          if (token !== undefined) {
            nextFonts[slot] = token;
          }
        }
      }
      if (resolvedMajor !== undefined) {
        nextFonts["major"] = resolvedMajor;
      }
      if (resolvedMinor !== undefined) {
        nextFonts["minor"] = resolvedMinor;
      }
      deckNode.set({
        theme: defineTheme({
          ...upgradedBase,
          ...(themeName === undefined ? {} : { name: themeName }),
          fonts: nextFonts,
        }),
      });
    } else if (existingDeckTheme !== undefined) {
      const updated = attachFontsToThemeInput(existingDeckTheme);
      if (updated.changed && updated.theme !== undefined) {
        deckNode.set({ theme: updated.theme });
      }
    }
  } else if (existingDeckTheme !== undefined) {
    const updated = attachFontsToThemeInput(existingDeckTheme);
    if (updated.changed && updated.theme !== undefined) {
      deckNode.set({ theme: updated.theme });
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
