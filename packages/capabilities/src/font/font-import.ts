import {
  defaultTheme,
  defineTheme,
  isFont,
  isTheme,
  masterTextStyleCategoryForRole,
  matchMasterPlaceholderSpec,
  type Font,
  type FontFace,
  type MasterPlaceholderMap,
  type NativePlaceholderRef,
  type PlaceholderListStyles,
  type PlaceholderMap,
  type Theme,
  type ThemeInput,
  type ThemeOverride,
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

import { fontFamilyKey, inspectImportedEmbeddedFont } from "./font-engine.ts";
import { font } from "./font.ts";

type EmbeddedSlotWeights = Partial<Record<EmbeddedFontSlot, number>>;

type EmbeddedFontLookup = {
  readonly fontsByKey: ReadonlyMap<string, Font>;
  readonly slotWeightsByKey: ReadonlyMap<string, EmbeddedSlotWeights>;
};

type ThemeFontMap = {
  readonly major?: Font | string | undefined;
  readonly minor?: Font | string | undefined;
};

type FontCascadeContext = {
  readonly themeFonts: ThemeFontMap;
  readonly fontFamily: unknown;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly effectiveWeight: number;
};

type MasterTextStylesRecord = {
  readonly titleStyle?: PlaceholderListStyles;
  readonly bodyStyle?: PlaceholderListStyles;
  readonly otherStyle?: PlaceholderListStyles;
};

type MasterCascadeLookup = {
  readonly themeFonts: ThemeFontMap;
  readonly textStyles?: MasterTextStylesRecord | undefined;
  readonly placeholders?: MasterPlaceholderMap | undefined;
  readonly footerPlaceholders?: Readonly<Record<string, unknown>> | undefined;
};

type LayoutCascadeLookup = {
  readonly themeFonts: ThemeFontMap;
  readonly master?: MasterCascadeLookup | undefined;
  readonly placeholders?: PlaceholderMap | undefined;
  readonly footerPlaceholders?: Readonly<Record<string, unknown>> | undefined;
};

const reservedThemeFontPattern = /^(?:major|minor|\+(?:mj|mn)-)/iu;

function isPlainRecord(
  value: unknown,
): value is Readonly<Record<string, unknown>> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function fallbackWeightForSlot(slot: EmbeddedFontSlot): number {
  return slot === "bold" || slot === "boldItalic" ? 700 : 400;
}

function fallbackStyleForSlot(slot: EmbeddedFontSlot): "normal" | "italic" {
  return slot === "italic" || slot === "boldItalic" ? "italic" : "normal";
}

function isBoldWeight(weight: unknown): boolean | undefined {
  if (weight === "bold") return true;
  if (weight === "normal") return false;
  if (typeof weight === "number") return weight >= 600;
  return undefined;
}

function numericWeightOrUndefined(weight: unknown): number | undefined {
  if (weight === "normal") return 400;
  if (weight === "bold") return 700;
  if (typeof weight === "number") return weight;
  return undefined;
}

function isItalicFontStyle(style: unknown): boolean | undefined {
  if (style === "italic") return true;
  if (style === "normal") return false;
  return undefined;
}

function slotForBoldItalic(bold: boolean, italic: boolean): EmbeddedFontSlot {
  if (bold) return italic ? "boldItalic" : "bold";
  return italic ? "italic" : "regular";
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

function resolveEffectiveFamilyKey(
  fontFamily: unknown,
  themeFonts: ThemeFontMap,
): string | undefined {
  if (isFont(fontFamily)) {
    return fontFamilyKey(fontFamily.family);
  }
  if (typeof fontFamily !== "string" || fontFamily.trim() === "") {
    return undefined;
  }
  if (fontFamily === "major" || fontFamily === "minor") {
    const mapped = themeFonts[fontFamily];
    if (isFont(mapped)) return fontFamilyKey(mapped.family);
    if (typeof mapped === "string" && mapped.trim() !== "") {
      return fontFamilyKey(mapped);
    }
  }
  return fontFamilyKey(fontFamily);
}

function applyStyleContextOnly(
  inherited: FontCascadeContext,
  style: unknown,
): FontCascadeContext {
  if (!isPlainRecord(style)) return inherited;
  const nextFontFamily =
    style["fontFamily"] !== undefined
      ? style["fontFamily"]
      : inherited.fontFamily;
  const nextBold = isBoldWeight(style["fontWeight"]) ?? inherited.bold;
  const nextItalic = isItalicFontStyle(style["fontStyle"]) ?? inherited.italic;
  const nextWeight =
    numericWeightOrUndefined(style["fontWeight"]) ?? inherited.effectiveWeight;
  return {
    themeFonts: inherited.themeFonts,
    fontFamily: nextFontFamily,
    bold: nextBold,
    italic: nextItalic,
    effectiveWeight: nextWeight,
  };
}

function foldStyleLayers(
  base: FontCascadeContext,
  layers: readonly unknown[],
): FontCascadeContext {
  let current = base;
  for (const layer of layers) {
    current = applyStyleContextOnly(current, layer);
  }
  return current;
}

function upgradeStyleRecordFontAndWeight(
  style: unknown,
  inherited: FontCascadeContext,
  lookup: EmbeddedFontLookup,
  options?: { readonly backfillInheritedLeaf?: boolean },
): {
  readonly value?: Readonly<Record<string, unknown>>;
  readonly changed: boolean;
  readonly nextContext: FontCascadeContext;
} {
  if (!isPlainRecord(style)) {
    if (options?.backfillInheritedLeaf === true) {
      const familyKey = resolveEffectiveFamilyKey(
        inherited.fontFamily,
        inherited.themeFonts,
      );
      const slot = slotForBoldItalic(inherited.bold, inherited.italic);
      const slotWeight =
        familyKey === undefined
          ? undefined
          : lookup.slotWeightsByKey.get(familyKey)?.[slot];
      if (
        slotWeight !== undefined &&
        slotWeight !== inherited.effectiveWeight
      ) {
        return {
          changed: true,
          value: { fontWeight: slotWeight },
          nextContext: {
            ...inherited,
            effectiveWeight: slotWeight,
          },
        };
      }
    }
    return { changed: false, nextContext: inherited };
  }

  const matched = upgradeFontToken(style["fontFamily"], lookup.fontsByKey);
  const nextRec: Record<string, unknown> =
    matched === undefined ? { ...style } : { ...style, fontFamily: matched };
  let changed = matched !== undefined;

  const hasLocalFontFamily = style["fontFamily"] !== undefined;
  const hasLocalFontWeight = style["fontWeight"] !== undefined;
  const hasLocalFontStyle = style["fontStyle"] !== undefined;
  const hasAnyLocalFontDecl =
    hasLocalFontFamily || hasLocalFontWeight || hasLocalFontStyle;

  const nextFontFamily = hasLocalFontFamily
    ? nextRec["fontFamily"]
    : inherited.fontFamily;
  const nextBold = isBoldWeight(style["fontWeight"]) ?? inherited.bold;
  const nextItalic = isItalicFontStyle(style["fontStyle"]) ?? inherited.italic;
  const currentEffectiveWeight =
    numericWeightOrUndefined(style["fontWeight"]) ?? inherited.effectiveWeight;

  const familyKey = resolveEffectiveFamilyKey(
    nextFontFamily,
    inherited.themeFonts,
  );
  const slot = slotForBoldItalic(nextBold, nextItalic);
  const slotWeights =
    familyKey === undefined
      ? undefined
      : lookup.slotWeightsByKey.get(familyKey);
  let slotWeight = slotWeights?.[slot];
  if (slotWeight === undefined && nextBold && slotWeights !== undefined) {
    // Unlinked >=600 faces (e.g. SemiBold with fsSelection.bold=false) embed in regular/italic
    // while levelDefRPrXml still writes b="1" for numeric fontWeight >= 600.
    const unlinkedSlot: EmbeddedFontSlot = nextItalic ? "italic" : "regular";
    const candidate = slotWeights[unlinkedSlot];
    if (candidate !== undefined && candidate >= 600) {
      slotWeight = candidate;
    }
  }

  let nextEffectiveWeight = currentEffectiveWeight;
  if (
    slotWeight !== undefined &&
    slotWeight !== currentEffectiveWeight &&
    (hasAnyLocalFontDecl || options?.backfillInheritedLeaf === true)
  ) {
    nextRec["fontWeight"] = slotWeight;
    changed = true;
    nextEffectiveWeight = slotWeight;
  }

  const nextContext: FontCascadeContext = {
    themeFonts: inherited.themeFonts,
    fontFamily: nextFontFamily,
    bold: nextBold,
    italic: nextItalic,
    effectiveWeight: nextEffectiveWeight,
  };

  return changed
    ? { changed: true, value: nextRec, nextContext }
    : { changed: false, nextContext };
}

function hasUnstyledTextLeaf(node: CommonNode): boolean {
  if (node.type === "span" || node.type === "slide-number") {
    return true;
  }
  if (
    node.type === "text" ||
    node.type === "cell" ||
    node.type === "notes" ||
    node.type === "paragraph"
  ) {
    return (
      node.children.length === 0 ||
      node.children.some((child) => !(child instanceof CommonNode))
    );
  }
  if (node.type === "shape") {
    return node.children.some((child) => !(child instanceof CommonNode));
  }
  return false;
}

function upgradePlaceholderListStylesFont(
  listStyles: PlaceholderListStyles | undefined,
  contextForLevel: (levelNum: number) => FontCascadeContext,
  lookup: EmbeddedFontLookup,
): {
  readonly value: PlaceholderListStyles | undefined;
  readonly changed: boolean;
} {
  if (listStyles === undefined) return { value: undefined, changed: false };
  let changed = false;
  const next: Record<string, unknown> = {};
  for (const [levelKey, levelStyle] of Object.entries(listStyles)) {
    if (levelStyle === undefined) continue;
    const levelNum = Number(levelKey);
    const inherited = contextForLevel(
      Number.isInteger(levelNum) && levelNum >= 1 ? levelNum : 1,
    );
    const upgraded = upgradeStyleRecordFontAndWeight(
      levelStyle,
      inherited,
      lookup,
    );
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
  baseContext: FontCascadeContext,
  masterTextStyles: MasterTextStylesRecord | undefined,
  masterPlaceholders: MasterPlaceholderMap | undefined,
  lookup: EmbeddedFontLookup,
): {
  readonly value: PlaceholderMap | undefined;
  readonly changed: boolean;
} {
  if (placeholders === undefined) return { value: undefined, changed: false };
  let changed = false;
  const next: Record<string, PlaceholderMap[string]> = {};
  for (const [key, spec] of Object.entries(placeholders)) {
    const isFooterRole =
      spec.type === "dt" || spec.type === "ftr" || spec.type === "sldNum";
    const category = masterTextStyleCategoryForRole(spec.type);
    const masterTx = isFooterRole ? undefined : masterTextStyles?.[category];
    const masterPh =
      masterPlaceholders === undefined
        ? undefined
        : matchMasterPlaceholderSpec(masterPlaceholders, spec.type);

    const styleInherited = foldStyleLayers(baseContext, [
      masterTx?.[1],
      masterPh?.style,
      masterPh?.listStyles?.[1],
    ]);
    const styleUp = upgradeStyleRecordFontAndWeight(
      spec.style,
      styleInherited,
      lookup,
    );
    const effectiveSpecStyle = styleUp.changed ? styleUp.value : spec.style;

    const listUp = upgradePlaceholderListStylesFont(
      spec.listStyles,
      (levelNum) => {
        const key1to9 = Math.min(9, Math.max(1, levelNum)) as
          | 1
          | 2
          | 3
          | 4
          | 5
          | 6
          | 7
          | 8
          | 9;
        return foldStyleLayers(baseContext, [
          masterTx?.[key1to9],
          masterPh?.style,
          masterPh?.listStyles?.[key1to9],
          effectiveSpecStyle,
        ]);
      },
      lookup,
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

function upgradeFooterPlaceholdersRecord(
  rawFooterPlaceholders: unknown,
  baseContext: FontCascadeContext,
  inheritedFooterPlaceholders: Readonly<Record<string, unknown>> | undefined,
  lookup: EmbeddedFontLookup,
): {
  readonly value?: Readonly<Record<string, unknown>>;
  readonly changed: boolean;
} {
  if (!isPlainRecord(rawFooterPlaceholders)) {
    return { changed: false };
  }
  let changed = false;
  const nextPlaceholders: Record<string, unknown> = {
    ...rawFooterPlaceholders,
  };
  for (const [slot, spec] of Object.entries(nextPlaceholders)) {
    if (!isPlainRecord(spec)) continue;
    const inheritedSpec = inheritedFooterPlaceholders?.[slot];
    const inheritedContext = isPlainRecord(inheritedSpec)
      ? applyStyleContextOnly(baseContext, inheritedSpec["style"])
      : baseContext;
    const nextStyle = upgradeStyleRecordFontAndWeight(
      spec["style"],
      inheritedContext,
      lookup,
    );
    if (nextStyle.changed && nextStyle.value !== undefined) {
      nextPlaceholders[slot] = {
        ...spec,
        style: nextStyle.value,
      };
      changed = true;
    }
  }
  return changed
    ? { value: nextPlaceholders, changed: true }
    : { value: rawFooterPlaceholders, changed: false };
}

function upgradeThemeFontMap(
  fonts: Readonly<Record<string, unknown>> | undefined,
  fontsByKey: ReadonlyMap<string, Font>,
): {
  readonly fonts?: {
    readonly major?: Font | string;
    readonly minor?: Font | string;
  };
  readonly changed: boolean;
} {
  if (fonts === undefined) return { changed: false };
  const nextFonts: { major?: Font | string; minor?: Font | string } = {};
  let changed = false;
  for (const slot of ["major", "minor"] as const) {
    const token = fonts[slot];
    if (token === undefined) continue;
    if (typeof token === "string") {
      const matched = fontsByKey.get(fontFamilyKey(token));
      if (matched !== undefined) {
        nextFonts[slot] = matched;
        changed = true;
        continue;
      }
      nextFonts[slot] = token;
    } else if (isFont(token)) {
      nextFonts[slot] = token;
    }
  }
  return changed ? { fonts: nextFonts, changed: true } : { changed: false };
}

function extractThemeFonts(themeValue: unknown): ThemeFontMap {
  if (isTheme(themeValue)) {
    return themeValue.fonts;
  }
  if (isPlainRecord(themeValue) && isPlainRecord(themeValue["fonts"])) {
    return themeValue["fonts"] as ThemeFontMap;
  }
  return {};
}

function upgradeThemeOverrideOnNode(
  node: CommonNode<string, Record<string, unknown>, Record<string, unknown>>,
  fontsByKey: ReadonlyMap<string, Font>,
): ThemeFontMap | undefined {
  const rawOverride = node.props["themeOverride"];
  if (!isPlainRecord(rawOverride)) return undefined;
  const rawFonts = isPlainRecord(rawOverride["fonts"])
    ? rawOverride["fonts"]
    : undefined;
  const upgraded = upgradeThemeFontMap(rawFonts, fontsByKey);
  if (upgraded.changed && upgraded.fonts !== undefined) {
    node.set({
      themeOverride: {
        ...(rawOverride as ThemeOverride),
        fonts: upgraded.fonts,
      },
    });
    return upgraded.fonts;
  }
  return rawFonts as ThemeFontMap | undefined;
}

function upgradeChartTextStyleContainer(
  container: unknown,
  inherited: FontCascadeContext,
  lookup: EmbeddedFontLookup,
): {
  readonly value?: Readonly<Record<string, unknown>>;
  readonly changed: boolean;
  readonly nextContext: FontCascadeContext;
} {
  if (!isPlainRecord(container)) {
    return { changed: false, nextContext: inherited };
  }
  const styleUp = upgradeStyleRecordFontAndWeight(
    container["style"],
    inherited,
    lookup,
  );
  if (!styleUp.changed || styleUp.value === undefined) {
    return { changed: false, nextContext: styleUp.nextContext };
  }
  return {
    changed: true,
    value: { ...container, style: styleUp.value },
    nextContext: styleUp.nextContext,
  };
}

function upgradeChartPropsFonts(
  node: CommonNode<string, Record<string, unknown>, Record<string, unknown>>,
  chartContext: FontCascadeContext,
  lookup: EmbeddedFontLookup,
): void {
  const patch: Record<string, unknown> = {};

  const titleUp = upgradeChartTextStyleContainer(
    node.props["title"],
    chartContext,
    lookup,
  );
  if (titleUp.changed && titleUp.value !== undefined) {
    patch["title"] = titleUp.value;
  }

  const rawAxes = node.props["axes"];
  if (isPlainRecord(rawAxes)) {
    let axesChanged = false;
    const nextAxes: Record<string, unknown> = { ...rawAxes };
    for (const [axisKey, rawAxis] of Object.entries(rawAxes)) {
      if (!isPlainRecord(rawAxis)) continue;
      const axisUp = upgradeChartTextStyleContainer(
        rawAxis,
        chartContext,
        lookup,
      );
      let nextAxis =
        axisUp.changed && axisUp.value !== undefined
          ? { ...axisUp.value }
          : { ...rawAxis };
      let axisModified = axisUp.changed;

      const axisTitleUp = upgradeChartTextStyleContainer(
        nextAxis["title"],
        axisUp.nextContext,
        lookup,
      );
      if (axisTitleUp.changed && axisTitleUp.value !== undefined) {
        nextAxis = { ...nextAxis, title: axisTitleUp.value };
        axisModified = true;
      }

      if (axisModified) {
        nextAxes[axisKey] = nextAxis;
        axesChanged = true;
      }
    }
    if (axesChanged) {
      patch["axes"] = nextAxes;
    }
  }

  const rawGroups = node.props["groups"];
  if (Array.isArray(rawGroups)) {
    let groupsChanged = false;
    const nextGroups = rawGroups.map((group) => {
      if (!isPlainRecord(group)) return group;
      const labelsUp = upgradeChartTextStyleContainer(
        group["dataLabels"],
        chartContext,
        lookup,
      );
      if (labelsUp.changed && labelsUp.value !== undefined) {
        groupsChanged = true;
        return { ...group, dataLabels: labelsUp.value };
      }
      return group;
    });
    if (groupsChanged) {
      patch["groups"] = nextGroups;
    }
  }

  if (Object.keys(patch).length > 0) {
    node.set(patch);
  }
}

const footerSlotToRole = {
  dateTime: "dt",
  footer: "ftr",
  slideNumber: "sldNum",
} as const;

function resolveFooterInheritanceLayers(
  slideSlot: "dateTime" | "footer" | "slideNumber",
  layout: LayoutCascadeLookup | undefined,
): readonly unknown[] {
  const role = footerSlotToRole[slideSlot];
  const masterFooterStyle = isPlainRecord(
    layout?.master?.footerPlaceholders?.[slideSlot],
  )
    ? layout.master.footerPlaceholders[slideSlot]["style"]
    : undefined;
  const masterPh =
    layout?.master?.placeholders === undefined
      ? undefined
      : matchMasterPlaceholderSpec(layout.master.placeholders, role);
  const layoutFooterStyle = isPlainRecord(
    layout?.footerPlaceholders?.[slideSlot],
  )
    ? layout.footerPlaceholders[slideSlot]["style"]
    : undefined;
  const layoutPh =
    layout?.placeholders === undefined
      ? undefined
      : matchMasterPlaceholderSpec(layout.placeholders, role);
  return [
    masterPh?.style,
    masterPh?.listStyles?.[1],
    masterFooterStyle,
    layoutPh?.style,
    layoutPh?.listStyles?.[1],
    layoutFooterStyle,
  ];
}
function buildMasterCascadeLookup(
  masterNode: CommonNode,
  deckThemeFonts: ThemeFontMap,
): MasterCascadeLookup {
  const masterFonts = extractThemeFonts(masterNode.props["theme"]);
  const themeFonts =
    Object.keys(masterFonts).length > 0
      ? { ...deckThemeFonts, ...masterFonts }
      : deckThemeFonts;
  return {
    themeFonts,
    textStyles: masterNode.props["textStyles"] as
      | MasterTextStylesRecord
      | undefined,
    placeholders: masterNode.props["placeholders"] as
      | MasterPlaceholderMap
      | undefined,
    footerPlaceholders: isPlainRecord(masterNode.props["footerPlaceholders"])
      ? masterNode.props["footerPlaceholders"]
      : undefined,
  };
}

function resolveSlideLayoutLookup(
  slideNode: CommonNode,
  deckThemeFonts: ThemeFontMap,
  masterLookups: ReadonlyMap<CommonNode, MasterCascadeLookup>,
  layoutLookups: ReadonlyMap<CommonNode, LayoutCascadeLookup>,
): LayoutCascadeLookup | undefined {
  const layoutProp = slideNode.props["layout"];
  if (layoutProp === undefined) return undefined;
  if (layoutProp instanceof CommonNode) {
    const cached = layoutLookups.get(layoutProp);
    if (cached !== undefined) return cached;
    const parentMaster =
      layoutProp.parent instanceof CommonNode &&
      layoutProp.parent.type === "slide-master"
        ? layoutProp.parent
        : undefined;
    const masterLookup =
      parentMaster === undefined
        ? undefined
        : (masterLookups.get(parentMaster) ??
          buildMasterCascadeLookup(parentMaster, deckThemeFonts));
    const baseFonts = masterLookup?.themeFonts ?? deckThemeFonts;
    const overrideFonts = extractThemeFonts(layoutProp.props["themeOverride"]);
    return {
      themeFonts:
        Object.keys(overrideFonts).length > 0
          ? { ...baseFonts, ...overrideFonts }
          : baseFonts,
      ...(masterLookup === undefined ? {} : { master: masterLookup }),
      placeholders: layoutProp.props["placeholders"] as
        | PlaceholderMap
        | undefined,
      footerPlaceholders: isPlainRecord(layoutProp.props["footerPlaceholders"])
        ? layoutProp.props["footerPlaceholders"]
        : undefined,
    };
  }
  if (isPlainRecord(layoutProp)) {
    const overrideFonts = extractThemeFonts(layoutProp["themeOverride"]);
    return {
      themeFonts:
        Object.keys(overrideFonts).length > 0
          ? { ...deckThemeFonts, ...overrideFonts }
          : deckThemeFonts,
      placeholders: layoutProp["placeholders"] as PlaceholderMap | undefined,
      footerPlaceholders: isPlainRecord(layoutProp["footerPlaceholders"])
        ? layoutProp["footerPlaceholders"]
        : undefined,
    };
  }
  return undefined;
}

function attachEmbeddedFontsToSubtree(
  node: CommonNode,
  inherited: FontCascadeContext,
  lookup: EmbeddedFontLookup,
  masterLookups: Map<CommonNode, MasterCascadeLookup>,
  layoutLookups: Map<CommonNode, LayoutCascadeLookup>,
  activeLayout: LayoutCascadeLookup | undefined,
  placeholderContext?: {
    readonly layoutPh: PlaceholderMap[string] | undefined;
    readonly masterPh: MasterPlaceholderMap[string] | undefined;
    readonly masterTx: PlaceholderListStyles | undefined;
    readonly slideTextNodeStyle: unknown;
  },
): void {
  const mutableNode = node as CommonNode<
    string,
    Record<string, unknown>,
    Record<string, unknown>
  >;

  let currentContext = inherited;
  let nextLayout = activeLayout;

  if (node.type === "slide-master") {
    const masterThemeFonts = extractThemeFonts(node.props["theme"]);
    const effectiveThemeFonts =
      Object.keys(masterThemeFonts).length > 0
        ? { ...inherited.themeFonts, ...masterThemeFonts }
        : inherited.themeFonts;
    currentContext = {
      themeFonts: effectiveThemeFonts,
      fontFamily: "minor",
      bold: false,
      italic: false,
      effectiveWeight: 400,
    };

    const rawTextStyles = node.props["textStyles"] as
      | MasterTextStylesRecord
      | undefined;
    if (rawTextStyles !== undefined) {
      const levelContext = (_levelNum: number): FontCascadeContext =>
        currentContext;
      const titleUp = upgradePlaceholderListStylesFont(
        rawTextStyles.titleStyle,
        levelContext,
        lookup,
      );
      const bodyUp = upgradePlaceholderListStylesFont(
        rawTextStyles.bodyStyle,
        levelContext,
        lookup,
      );
      const otherUp = upgradePlaceholderListStylesFont(
        rawTextStyles.otherStyle,
        levelContext,
        lookup,
      );
      if (titleUp.changed || bodyUp.changed || otherUp.changed) {
        mutableNode.set({
          textStyles: {
            ...(titleUp.value === undefined
              ? {}
              : { titleStyle: titleUp.value }),
            ...(bodyUp.value === undefined ? {} : { bodyStyle: bodyUp.value }),
            ...(otherUp.value === undefined
              ? {}
              : { otherStyle: otherUp.value }),
          },
        });
      }
    }

    const updatedTextStyles = node.props["textStyles"] as
      | MasterTextStylesRecord
      | undefined;
    const rawPlaceholders = node.props["placeholders"] as
      | MasterPlaceholderMap
      | undefined;
    const upgradedPh = upgradePlaceholdersFont(
      rawPlaceholders,
      currentContext,
      updatedTextStyles,
      undefined,
      lookup,
    );
    if (upgradedPh.changed && upgradedPh.value !== undefined) {
      mutableNode.set({ placeholders: upgradedPh.value });
    }

    const footerUp = upgradeFooterPlaceholdersRecord(
      node.props["footerPlaceholders"],
      currentContext,
      undefined,
      lookup,
    );
    if (footerUp.changed && footerUp.value !== undefined) {
      mutableNode.set({ footerPlaceholders: footerUp.value });
    }

    masterLookups.set(
      node,
      buildMasterCascadeLookup(node, inherited.themeFonts),
    );
  } else if (node.type === "slide-layout") {
    const parentMaster =
      node.parent instanceof CommonNode && node.parent.type === "slide-master"
        ? masterLookups.get(node.parent)
        : undefined;
    const overrideFonts = upgradeThemeOverrideOnNode(
      mutableNode,
      lookup.fontsByKey,
    );
    const baseFonts = parentMaster?.themeFonts ?? inherited.themeFonts;
    const effectiveThemeFonts =
      overrideFonts !== undefined && Object.keys(overrideFonts).length > 0
        ? { ...baseFonts, ...overrideFonts }
        : baseFonts;
    currentContext = {
      themeFonts: effectiveThemeFonts,
      fontFamily: "minor",
      bold: false,
      italic: false,
      effectiveWeight: 400,
    };

    const rawPlaceholders = node.props["placeholders"] as
      | PlaceholderMap
      | undefined;
    const upgradedPh = upgradePlaceholdersFont(
      rawPlaceholders,
      currentContext,
      parentMaster?.textStyles,
      parentMaster?.placeholders,
      lookup,
    );
    if (upgradedPh.changed && upgradedPh.value !== undefined) {
      mutableNode.set({ placeholders: upgradedPh.value });
    }

    const footerUp = upgradeFooterPlaceholdersRecord(
      node.props["footerPlaceholders"],
      currentContext,
      parentMaster?.footerPlaceholders,
      lookup,
    );
    if (footerUp.changed && footerUp.value !== undefined) {
      mutableNode.set({ footerPlaceholders: footerUp.value });
    }

    layoutLookups.set(node, {
      themeFonts: effectiveThemeFonts,
      ...(parentMaster === undefined ? {} : { master: parentMaster }),
      placeholders: node.props["placeholders"] as PlaceholderMap | undefined,
      footerPlaceholders: isPlainRecord(node.props["footerPlaceholders"])
        ? node.props["footerPlaceholders"]
        : undefined,
    });
  } else if (node.type === "slide") {
    nextLayout = resolveSlideLayoutLookup(
      node,
      inherited.themeFonts,
      masterLookups,
      layoutLookups,
    );
    const overrideFonts = upgradeThemeOverrideOnNode(
      mutableNode,
      lookup.fontsByKey,
    );
    const baseFonts = nextLayout?.themeFonts ?? inherited.themeFonts;
    const effectiveThemeFonts = {
      ...baseFonts,
      ...overrideFonts,
    };
    currentContext = {
      ...inherited,
      themeFonts: effectiveThemeFonts,
    };
  }

  let childPlaceholderContext = placeholderContext;
  let nodeStyleInherited = currentContext;
  if (
    node.type !== "deck" &&
    node.type !== "slide-master" &&
    node.type !== "slide-layout" &&
    node.type !== "slide"
  ) {
    if (
      placeholderContext !== undefined &&
      node.type === "paragraph" &&
      isPlainRecord(node.props["style"])
    ) {
      const listProp = node.props["style"]["list"];
      const levelIndex0 =
        isPlainRecord(listProp) && typeof listProp["level"] === "number"
          ? listProp["level"]
          : 0;
      const key1to9 = Math.min(9, Math.max(1, levelIndex0 + 1)) as
        | 1
        | 2
        | 3
        | 4
        | 5
        | 6
        | 7
        | 8
        | 9;
      nodeStyleInherited = foldStyleLayers(
        {
          themeFonts: currentContext.themeFonts,
          fontFamily: "minor",
          bold: false,
          italic: false,
          effectiveWeight: 400,
        },
        [
          placeholderContext.masterTx?.[key1to9],
          placeholderContext.masterPh?.style,
          placeholderContext.masterPh?.listStyles?.[key1to9],
          placeholderContext.layoutPh?.style,
          placeholderContext.layoutPh?.listStyles?.[key1to9],
          placeholderContext.slideTextNodeStyle,
        ],
      );
    } else if (
      node.props["placeholder"] !== undefined ||
      node.props["nativePlaceholder"] !== undefined
    ) {
      const placeholderKey =
        typeof node.props["placeholder"] === "string"
          ? node.props["placeholder"]
          : undefined;
      const nativeRef = isPlainRecord(node.props["nativePlaceholder"])
        ? (node.props["nativePlaceholder"] as NativePlaceholderRef)
        : undefined;
      const layoutPh =
        placeholderKey === undefined
          ? undefined
          : nextLayout?.placeholders?.[placeholderKey];
      const role = layoutPh?.type ?? nativeRef?.type;
      if (role !== undefined) {
        const category = masterTextStyleCategoryForRole(role);
        const masterTx = nextLayout?.master?.textStyles?.[category];
        const masterPh =
          nextLayout?.master?.placeholders === undefined
            ? undefined
            : matchMasterPlaceholderSpec(nextLayout.master.placeholders, role);
        nodeStyleInherited = foldStyleLayers(
          {
            themeFonts: currentContext.themeFonts,
            fontFamily: "minor",
            bold: false,
            italic: false,
            effectiveWeight: 400,
          },
          [
            masterTx?.[1],
            masterPh?.style,
            masterPh?.listStyles?.[1],
            layoutPh?.style,
            layoutPh?.listStyles?.[1],
          ],
        );
        childPlaceholderContext = {
          layoutPh,
          masterPh,
          masterTx,
          slideTextNodeStyle: node.props["style"],
        };
      }
    }
  }

  const styleUp = upgradeStyleRecordFontAndWeight(
    node.props["style"],
    nodeStyleInherited,
    lookup,
    { backfillInheritedLeaf: hasUnstyledTextLeaf(node) },
  );
  if (styleUp.changed && styleUp.value !== undefined) {
    const patch: Record<string, unknown> = {};
    const prevStyle = isPlainRecord(node.props["style"])
      ? node.props["style"]
      : undefined;
    if (styleUp.value["fontFamily"] !== prevStyle?.["fontFamily"]) {
      patch["fontFamily"] = styleUp.value["fontFamily"];
    }
    if (styleUp.value["fontWeight"] !== prevStyle?.["fontWeight"]) {
      patch["fontWeight"] = styleUp.value["fontWeight"];
    }
    if (Object.keys(patch).length > 0) {
      mutableNode.style(patch);
    }
    if (
      childPlaceholderContext !== undefined &&
      placeholderContext === undefined
    ) {
      childPlaceholderContext = {
        ...childPlaceholderContext,
        slideTextNodeStyle: node.props["style"],
      };
    }
  }

  if (node.type === "chart") {
    upgradeChartPropsFonts(mutableNode, styleUp.nextContext, lookup);
  }

  if (node.type === "slide") {
    for (const slideSlot of ["dateTime", "footer", "slideNumber"] as const) {
      const rawSlot = node.props[slideSlot];
      if (!isPlainRecord(rawSlot)) continue;
      const footerInherited = foldStyleLayers(
        currentContext,
        resolveFooterInheritanceLayers(slideSlot, nextLayout),
      );
      const nextStyle = upgradeStyleRecordFontAndWeight(
        rawSlot["style"],
        footerInherited,
        lookup,
      );
      if (nextStyle.changed && nextStyle.value !== undefined) {
        mutableNode.set({
          [slideSlot]: {
            ...rawSlot,
            style: nextStyle.value,
          },
        });
      }
    }
  }

  const childInherited =
    node.type === "deck" ||
    node.type === "slide-master" ||
    node.type === "slide-layout" ||
    node.type === "slide"
      ? {
          ...styleUp.nextContext,
          themeFonts: currentContext.themeFonts,
        }
      : styleUp.nextContext;

  for (const child of node.children) {
    if (child instanceof CommonNode) {
      attachEmbeddedFontsToSubtree(
        child,
        childInherited,
        lookup,
        masterLookups,
        layoutLookups,
        nextLayout,
        node.type === "slide-master" ||
          node.type === "slide-layout" ||
          node.type === "slide"
          ? undefined
          : childPlaceholderContext,
      );
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
  const slotWeightsByKey = new Map<string, EmbeddedSlotWeights>();
  for (const record of records) {
    const faces: FontFace[] = [];
    const slotWeights: Partial<Record<EmbeddedFontSlot, number>> = {};
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

      const inspected = inspectImportedEmbeddedFont(
        unpacked.ttfBytes,
        record.typeface,
      );
      if (inspected === undefined) {
        diagnostics.push({
          code: "corrupt-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: slot.partPath, xpath: slot.xpath },
        });
        continue;
      }

      if (!inspected.familyMatches) {
        requiresSubstitute = true;
      }

      const weight = inspected.weight ?? fallbackWeightForSlot(slot.slot);
      const style: "normal" | "italic" =
        inspected.style ?? fallbackStyleForSlot(slot.slot);

      slotWeights[slot.slot] = weight;
      faces.push({
        weight,
        style,
        data: new Uint8Array(unpacked.ttfBytes),
      });
    }

    if (
      faces.length > 0 &&
      record.typeface.trim() !== "" &&
      !reservedThemeFontPattern.test(record.typeface.trim())
    ) {
      const key = fontFamilyKey(record.typeface);
      fontsByKey.set(
        key,
        font(
          record.typeface,
          faces,
          requiresSubstitute ? { substitute: true } : {},
        ),
      );
      slotWeightsByKey.set(key, slotWeights);
    }
  }

  const schemeDiagnostics: ImportDiagnostic[] = [
    ...(context.existingDiagnostics ?? []),
  ];
  const previousDiagnosticsCount = schemeDiagnostics.length;
  const scheme = readPresentationThemeFontScheme(pkg, schemeDiagnostics);
  diagnostics.push(...schemeDiagnostics.slice(previousDiagnosticsCount));
  const attachFontsToThemeInput = (
    existing: ThemeInput | undefined,
  ): {
    readonly theme?: ReturnType<typeof defineTheme>;
    readonly changed: boolean;
  } => {
    if (existing === undefined) return { changed: false };
    const existingFonts = existing.fonts ?? {};
    const nextFonts: { major?: Font | string; minor?: Font | string } = {};
    let changed = false;
    for (const slot of ["major", "minor"] as const) {
      const token = existingFonts[slot];
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
      const nextFonts: { major?: Font | string; minor?: Font | string } = {};
      if (upgradedBase?.fonts !== undefined) {
        for (const slot of ["major", "minor"] as const) {
          const token = upgradedBase.fonts[slot];
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
    const deckThemeFonts = extractThemeFonts(deckNode.props["theme"]);
    const initialContext: FontCascadeContext = {
      themeFonts: deckThemeFonts,
      fontFamily: "minor",
      bold: false,
      italic: false,
      effectiveWeight: 400,
    };
    attachEmbeddedFontsToSubtree(
      deckNode,
      initialContext,
      { fontsByKey, slotWeightsByKey },
      new Map(),
      new Map(),
      undefined,
    );
  }

  return { diagnostics };
}

/** Deck-level import handler wired into `embeddedFontMechanism`. */
export const embeddedFontImportHandler: DeckImportHandler = {
  target: "deck",
  read: readEmbeddedFontsIntoDeck,
};
