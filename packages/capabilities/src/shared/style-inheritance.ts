import {
  masterTextStyleCategoryForRole,
  resolvePlaceholderBoxFromLayers,
  type PaddingStyle,
  type PlaceholderLevelStyle,
  type PlaceholderListStyles,
  type PlaceholderResolutionContext,
  type TextStyle,
  type VisualStyle,
} from "@pptxdk/core/capability";

import { normalizeLegacyFillStyle } from "./visual.ts";

export type SharedStyleProperties = TextStyle & PaddingStyle & VisualStyle;

export type StylePropertyRule<T> = {
  /** Whether Deck.style / Slide.style / Box.style flows down to free descendants or stays local. */
  readonly freeContainer: "inherit" | "localOnly";
  /** Target OOXML container within a native placeholder cascade. */
  readonly placeholderLayer: "levelStyle" | "bodyPr" | "spPr";
  /** Applies a declared override from a higher-precedence layer onto an accumulated value. */
  readonly apply: (
    inherited: T | undefined,
    localValue: T | undefined,
    isDeclaredLocally: boolean,
  ) => T | undefined;
};

/**
 * Compile-time exhaustive rule registry covering every shared style property across
 * TextStyle, PaddingStyle, and VisualStyle. Adding a shared style property without
 * declaring its free-container and placeholder behavior fails TypeScript compilation.
 */
export type ExhaustiveStyleRuleMap<
  Style extends object = SharedStyleProperties,
> = {
  readonly [Key in keyof Style]-?: StylePropertyRule<Style[Key]>;
};

function replaceWhenDeclared<T>(
  inherited: T | undefined,
  localValue: T | undefined,
  isDeclaredLocally: boolean,
): T | undefined {
  return isDeclaredLocally && localValue !== undefined ? localValue : inherited;
}

function levelRule<T>(): StylePropertyRule<T> {
  return {
    freeContainer: "inherit",
    placeholderLayer: "levelStyle",
    apply: replaceWhenDeclared,
  };
}

function bodyRule<T>(
  freeContainer: "inherit" | "localOnly",
): StylePropertyRule<T> {
  return {
    freeContainer,
    placeholderLayer: "bodyPr",
    apply: replaceWhenDeclared,
  };
}

function spPrRule<T>(): StylePropertyRule<T> {
  return {
    freeContainer: "localOnly",
    placeholderLayer: "spPr",
    apply: replaceWhenDeclared,
  };
}

export const sharedStyleRules: ExhaustiveStyleRuleMap = Object.freeze({
  // Text run & paragraph properties (inherited on free containers, level-cascaded on placeholders)
  fontFamily: levelRule<TextStyle["fontFamily"]>(),
  fontSize: levelRule<TextStyle["fontSize"]>(),
  fontWeight: levelRule<TextStyle["fontWeight"]>(),
  fontStyle: levelRule<TextStyle["fontStyle"]>(),
  color: levelRule<TextStyle["color"]>(),
  letterSpacing: levelRule<TextStyle["letterSpacing"]>(),
  textTransform: levelRule<TextStyle["textTransform"]>(),
  textDecoration: levelRule<TextStyle["textDecoration"]>(),
  textShadow: {
    freeContainer: "inherit",
    placeholderLayer: "levelStyle",
    apply(inherited, localValue, isDeclaredLocally) {
      if (!isDeclaredLocally || localValue === undefined) return inherited;
      return localValue === "none" ? undefined : localValue;
    },
  },
  lang: levelRule<TextStyle["lang"]>(),
  lineHeight: levelRule<TextStyle["lineHeight"]>(),
  align: levelRule<TextStyle["align"]>(),
  spacingAfter: levelRule<TextStyle["spacingAfter"]>(),

  // Text body properties (bodyPr)
  verticalAlign: bodyRule<TextStyle["verticalAlign"]>("inherit"),
  textWrap: bodyRule<TextStyle["textWrap"]>("inherit"),
  padding: bodyRule<PaddingStyle["padding"]>("localOnly"),
  paddingTop: bodyRule<PaddingStyle["paddingTop"]>("localOnly"),
  paddingRight: bodyRule<PaddingStyle["paddingRight"]>("localOnly"),
  paddingBottom: bodyRule<PaddingStyle["paddingBottom"]>("localOnly"),
  paddingLeft: bodyRule<PaddingStyle["paddingLeft"]>("localOnly"),

  // Shape visual properties (spPr, never inherited by free children)
  background: spPrRule<VisualStyle["background"]>(),
  backgroundColor: spPrRule<VisualStyle["backgroundColor"]>(),
  backgroundImage: spPrRule<VisualStyle["backgroundImage"]>(),
  border: spPrRule<VisualStyle["border"]>(),
  borderColor: spPrRule<VisualStyle["borderColor"]>(),
  borderWidth: spPrRule<VisualStyle["borderWidth"]>(),
  borderStyle: spPrRule<VisualStyle["borderStyle"]>(),
  borderRadius: spPrRule<VisualStyle["borderRadius"]>(),
  opacity: spPrRule<VisualStyle["opacity"]>(),
  boxShadow: spPrRule<VisualStyle["boxShadow"]>(),
  softEdge: spPrRule<VisualStyle["softEdge"]>(),
  rotate: spPrRule<VisualStyle["rotate"]>(),
  flipH: spPrRule<VisualStyle["flipH"]>(),
  flipV: spPrRule<VisualStyle["flipV"]>(),
});

/** True iff `key` is an own defined declaration on `record`, honoring explicit `undefined` removal. */
export function hasLocalDeclaration(
  record: Readonly<Record<string, unknown>> | undefined,
  key: string,
): boolean {
  return (
    record !== undefined &&
    Object.hasOwn(record, key) &&
    record[key] !== undefined
  );
}

/** Collects the subset of `keys` explicitly declared with defined values on `record`. */
export function localDeclaredKeySet(
  record: Readonly<Record<string, unknown>> | undefined,
  keys: Iterable<string>,
): ReadonlySet<string> {
  const result = new Set<string>();
  if (record === undefined) return result;
  for (const key of keys) {
    if (hasLocalDeclaration(record, key)) result.add(key);
  }
  return result;
}

function levelStyleEntry(
  styles: PlaceholderListStyles | undefined,
  levelIndex0: number,
): PlaceholderLevelStyle | undefined {
  if (styles === undefined) return undefined;
  const oneBased = Math.min(9, Math.max(1, levelIndex0 + 1)) as
    | 1
    | 2
    | 3
    | 4
    | 5
    | 6
    | 7
    | 8
    | 9;
  return styles[oneBased];
}

function applyLayer(
  accumulator: Record<string, unknown>,
  source: Readonly<Record<string, unknown>> | undefined,
  layerFilter: "all" | StylePropertyRule<unknown>["placeholderLayer"],
): void {
  if (source === undefined) return;
  let effectiveSource = source;
  if (layerFilter === "all" || layerFilter === "spPr") {
    const normalized = normalizeLegacyFillStyle(source);
    if (
      normalized !== source ||
      hasLocalDeclaration(normalized, "background")
    ) {
      const cleaned = { ...normalized };
      delete cleaned["backgroundColor"];
      delete cleaned["backgroundImage"];
      effectiveSource = cleaned;
      delete accumulator["backgroundColor"];
      delete accumulator["backgroundImage"];
    }
    if (hasLocalDeclaration(effectiveSource, "border")) {
      if (
        hasLocalDeclaration(effectiveSource, "borderColor") ||
        hasLocalDeclaration(effectiveSource, "borderWidth") ||
        hasLocalDeclaration(effectiveSource, "borderStyle")
      ) {
        // Defer mixed-declaration error to resolveVisualModel on the unmerged layer.
      } else {
        delete accumulator["borderColor"];
        delete accumulator["borderWidth"];
        delete accumulator["borderStyle"];
      }
    } else if (hasLocalDeclaration(effectiveSource, "borderColor")) {
      delete accumulator["border"];
    }
  }
  for (const [key, rawRule] of Object.entries(sharedStyleRules)) {
    const rule = rawRule as StylePropertyRule<unknown>;
    if (layerFilter !== "all" && rule.placeholderLayer !== layerFilter) {
      continue;
    }
    if (!hasLocalDeclaration(effectiveSource, key)) continue;
    const next = rule.apply(accumulator[key], effectiveSource[key], true);
    if (next === undefined) delete accumulator[key];
    else accumulator[key] = next;
  }
}

/**
 * Evaluates Path B placeholder container defaults (`bodyPr` + level-0 base text properties):
 * Master Theme -> Master txStyles -> Master Ph -> Layout Ph -> Slide local style.
 * Note: Deck.style and Slide.style are intentionally excluded.
 */
export function resolvePlaceholderBaseTextStyle(
  placeholder: PlaceholderResolutionContext,
  slideLocalStyle: Readonly<Record<string, unknown>> | undefined,
): Readonly<Record<string, unknown>> {
  const result: Record<string, unknown> = {};
  const role =
    placeholder.layoutPlaceholder?.type ?? placeholder.nativeRef.type;
  const category = masterTextStyleCategoryForRole(role);
  const masterTxLevel1 = levelStyleEntry(
    placeholder.masterTextStyles?.[category],
    0,
  );
  const masterPhStyle = placeholder.masterPlaceholder?.style;
  const masterPhLevel1 = levelStyleEntry(
    placeholder.masterPlaceholder?.listStyles,
    0,
  );
  const layoutPhStyle = placeholder.layoutPlaceholder?.style;
  const layoutPhLevel1 = levelStyleEntry(
    placeholder.layoutPlaceholder?.listStyles,
    0,
  );

  // Precedence: master txStyles -> masterPh.style -> masterPh.listStyles[1]
  //          -> layoutPh.style -> layoutPh.listStyles[1] -> slide local style.
  applyLayer(result, masterTxLevel1, "levelStyle");
  applyLayer(result, masterPhStyle, "all");
  applyLayer(result, masterPhLevel1, "levelStyle");
  applyLayer(result, layoutPhStyle, "all");
  applyLayer(result, layoutPhLevel1, "levelStyle");
  applyLayer(result, slideLocalStyle, "all");
  return result;
}

/**
 * Evaluates level-specific paragraph/run defaults for a bound placeholder at `levelIndex0` (0..8),
 * cascading Master txStyles -> Master Ph -> Layout Ph -> Slide local style -> P/Span local style.
 */
export function resolvePlaceholderLevelTextStyle(
  placeholder: PlaceholderResolutionContext,
  levelIndex0: number,
  slideLocalStyle: Readonly<Record<string, unknown>> | undefined,
): PlaceholderLevelStyle {
  const result: Record<string, unknown> = {};
  const role =
    placeholder.layoutPlaceholder?.type ?? placeholder.nativeRef.type;
  const category = masterTextStyleCategoryForRole(role);
  const masterTxLevel = levelStyleEntry(
    placeholder.masterTextStyles?.[category],
    levelIndex0,
  );
  const masterPhLevel = levelStyleEntry(
    placeholder.masterPlaceholder?.listStyles,
    levelIndex0,
  );
  const layoutPhLevel = levelStyleEntry(
    placeholder.layoutPlaceholder?.listStyles,
    levelIndex0,
  );

  applyLayer(result, masterTxLevel, "levelStyle");
  applyLayer(result, placeholder.masterPlaceholder?.style, "levelStyle");
  applyLayer(result, masterPhLevel, "levelStyle");
  applyLayer(result, placeholder.layoutPlaceholder?.style, "levelStyle");
  applyLayer(result, layoutPhLevel, "levelStyle");
  applyLayer(result, slideLocalStyle, "levelStyle");

  const resolvedList =
    layoutPhLevel?.list ?? masterPhLevel?.list ?? masterTxLevel?.list;
  if (resolvedList !== undefined) {
    result["list"] = resolvedList;
  }
  return result;
}

/**
 * Evaluates visual (`spPr`) style overrides on a bound placeholder along Path B
 * without absorbing Deck.style or Slide.style.
 */
export function resolvePlaceholderVisualStyle(
  placeholder: PlaceholderResolutionContext,
  slideLocalStyle: Readonly<Record<string, unknown>> | undefined,
): Readonly<Record<string, unknown>> {
  const result: Record<string, unknown> = {};
  applyLayer(result, placeholder.masterPlaceholder?.style, "spPr");
  applyLayer(result, placeholder.layoutPlaceholder?.style, "spPr");
  applyLayer(result, slideLocalStyle, "spPr");
  return result;
}

/** Resolves effective box coordinates for a bound placeholder plus slide-local geometry tracking. */
export function resolvePlaceholderGeometry(
  placeholder: PlaceholderResolutionContext,
  props: Readonly<Record<string, unknown>>,
): {
  readonly x: unknown;
  readonly y: unknown;
  readonly width: unknown;
  readonly height: unknown;
  readonly hasLocalBoxOverride: boolean;
} {
  const style =
    typeof props["style"] === "object" &&
    props["style"] !== null &&
    !Array.isArray(props["style"])
      ? (props["style"] as Readonly<Record<string, unknown>>)
      : undefined;
  const hasLocalStyleKey = (key: string) =>
    style !== undefined && hasLocalDeclaration(style, key);
  const hasLocalBoxOverride =
    hasLocalDeclaration(props, "x") ||
    hasLocalDeclaration(props, "y") ||
    hasLocalDeclaration(props, "width") ||
    hasLocalDeclaration(props, "height") ||
    hasLocalStyleKey("left") ||
    hasLocalStyleKey("top") ||
    hasLocalStyleKey("right") ||
    hasLocalStyleKey("bottom") ||
    hasLocalStyleKey("inset") ||
    hasLocalStyleKey("width") ||
    hasLocalStyleKey("height");
  const resolved = resolvePlaceholderBoxFromLayers(
    [placeholder.masterPlaceholder, placeholder.layoutPlaceholder, props],
    placeholder.slideDimensions,
  );
  return {
    ...resolved,
    hasLocalBoxOverride,
  };
}
