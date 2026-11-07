import { CapabilityError } from "#/diagnostics.ts";

export type FontFaceStyle = "normal" | "italic";

/** Font files provided by the caller can be used for measurement or embedded according to build options. */
export type FontFace = {
  readonly weight?: number;
  readonly style?: FontFaceStyle;
  readonly data?: Uint8Array;
};

/** Font family and explicit measurement resources referenceable across deck themes and components. */
export type Font = {
  readonly family: string;
  readonly faces?: readonly FontFace[];
  readonly substitute?: boolean;
  readonly embeddable?: boolean;
};

export type FontOptions = {
  readonly substitute?: boolean;
  readonly embeddable?: boolean;
};

export type FontReference = string | Font;
export type FontSize = number | `${number}px` | `${number}pt`;
export type TextAlignment = "left" | "center" | "right" | "justify";

/** Inheritable text style fields shared across text objects and containers. */
export type TextStyle = {
  readonly fontFamily?: FontReference;
  readonly fontSize?: FontSize;
  readonly fontWeight?: number | "normal" | "bold";
  readonly fontStyle?: FontFaceStyle;
  readonly color?: ColorValue;
  readonly lineHeight?: FontSize;
  readonly verticalAlign?: "top" | "middle" | "bottom";
  /** Paragraph alignment when not explicitly overridden. */
  readonly align?: TextAlignment;
  /** Paragraph space-after when not explicitly overridden, in pixels. */
  readonly spacingAfter?: number;
  /** Character spacing in pixels. */
  readonly letterSpacing?: number;
  readonly textTransform?: "none" | "uppercase";
  readonly textDecoration?: "none" | "underline";
  readonly textShadow?: Shadow | "none";
  /** Whether the text box wraps normally or keeps each line on a single line. */
  readonly textWrap?: "normal" | "nowrap";
  /** Natural language (BCP 47 language tag, e.g. zh-CN, en-US). */
  readonly lang?: string;
};

export type GradientStop = {
  readonly color: ColorValue;
  readonly position: number;
};

export type LinearGradient = {
  readonly type: "linear-gradient";
  readonly angle: number;
  readonly stops: readonly GradientStop[];
};

export type RadialGradient = {
  readonly type: "radial-gradient";
  readonly stops: readonly GradientStop[];
};

export type Shadow = {
  readonly color?: ColorValue;
  readonly blur?: number;
  readonly distance?: number;
  readonly angle?: number;
};

/** Visual fields shared across native DrawingML shapes, images, text boxes, and visible Boxes. */
export type VisualStyle = {
  readonly backgroundColor?: ColorValue;
  readonly backgroundImage?: LinearGradient | RadialGradient;
  readonly borderColor?: ColorValue;
  readonly borderWidth?: number;
  readonly borderStyle?: "solid" | "dashed" | "dotted";
  readonly borderRadius?: number;
  readonly opacity?: number;
  readonly boxShadow?: Shadow;
  readonly softEdge?: number;
  readonly rotate?: number;
};

export type ImageVisualStyle = Pick<
  VisualStyle,
  | "borderColor"
  | "borderWidth"
  | "borderStyle"
  | "opacity"
  | "boxShadow"
  | "softEdge"
  | "rotate"
>;

export const nativeColorSlots = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export type NativeColorSlot = (typeof nativeColorSlots)[number];

/** Native theme color represented with DrawingML luminance modulation. */
export type ThemeShade = {
  readonly slot: NativeColorSlot;
  readonly lumMod: number;
  readonly lumOff?: number;
};

export type ThemeColor = string | ThemeShade;

/** Color entry can directly provide color, theme slot, or theme token, optionally with alpha. */
export type ColorValue =
  | string
  | ThemeShade
  | { readonly color: string | ThemeShade; readonly alpha?: number };

/** Final theme value applied to a color entry. */
export type ResolvedColor = {
  readonly hex: string;
  readonly alpha?: number;
  readonly slot?: NativeColorSlot;
  readonly lumMod?: number;
  readonly lumOff?: number;
};

export type FontToken = FontReference;

/** Deck-level theme input; major/minor fonts are simultaneously written to native PPTX theme. */
export type ThemeInput = {
  readonly name?: string;
  readonly colors?: Partial<Record<NativeColorSlot, string>> &
    Readonly<Record<string, ThemeColor>>;
  readonly fonts?: {
    readonly major?: FontToken;
    readonly minor?: FontToken;
  } & Readonly<Record<string, FontToken>>;
};

export type Theme = {
  readonly name: string;
  readonly colors: Readonly<Record<NativeColorSlot, string>> &
    Readonly<Record<string, ThemeColor>>;
  readonly fonts: {
    readonly major: FontToken;
    readonly minor: FontToken;
  } & Readonly<Record<string, FontToken>>;
};

/** Local theme scopes on slides or containers can only override color and font tokens already present in the deck. */
export type ThemeScope = {
  readonly colors?: Readonly<Record<string, ThemeColor>>;
  readonly fonts?: Readonly<Record<string, FontToken>>;
};

export type ResolvedThemeFont = {
  readonly family: string;
  readonly font?: Font;
  readonly themeFont?: "major" | "minor";
};

/** Theme, local scopes, and accumulated container text defaults in a single build. */
export type ThemeContext = {
  readonly theme: Theme;
  readonly textStyle: TextStyle;
  with(scope?: ThemeScope, style?: TextStyle): ThemeContext;
  resolveColor(value: ColorValue): ResolvedColor;
  resolveFont(value: FontReference): ResolvedThemeFont;
};

const officeColors: Record<NativeColorSlot, string> = {
  dk1: "#000000",
  lt1: "#FFFFFF",
  dk2: "#1F497D",
  lt2: "#EEECE1",
  accent1: "#4F81BD",
  accent2: "#C0504D",
  accent3: "#9BBB59",
  accent4: "#8064A2",
  accent5: "#4BACC6",
  accent6: "#F79646",
  hlink: "#0000FF",
  folHlink: "#800080",
};

const nativeSchemeReference: Record<NativeColorSlot, string> = {
  dk1: "tx1",
  lt1: "bg1",
  dk2: "tx2",
  lt2: "bg2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

const contextRecords = new WeakMap<
  ThemeContext,
  {
    readonly parent?: ThemeContext;
    readonly scope?: ThemeScope;
  }
>();

const textStyleKeys = new Set([
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "color",
  "lineHeight",
  "verticalAlign",
  "align",
  "spacingAfter",
  "letterSpacing",
  "textTransform",
  "textDecoration",
  "textShadow",
  "textWrap",
  "lang",
]);

/**
 * Asserts that a value is a valid BCP 47 language tag.
 * Rejects non-string values, empty or whitespace strings, and malformed tags.
 */
export function assertValidLang(value: unknown): string {
  if (
    typeof value !== "string" ||
    value.trim() === "" ||
    !/^[A-Za-z0-9]+(-[A-Za-z0-9]+)*$/.test(value)
  ) {
    throw new CapabilityError(
      "invalid-text-lang",
      `Text language must be a valid BCP 47 language tag: ${String(value)}`,
    );
  }
  try {
    const canonical = Intl.getCanonicalLocales(value);
    if (!canonical || canonical.length !== 1) {
      throw new CapabilityError(
        "invalid-text-lang",
        `Text language must be a valid BCP 47 language tag: ${value}`,
      );
    }
  } catch {
    throw new CapabilityError(
      "invalid-text-lang",
      `Text language must be a valid BCP 47 language tag: ${value}`,
    );
  }
  return value;
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Extract inheritable text fields from container style containing layout fields. */
export function textStyleFrom(value: unknown): TextStyle {
  if (value === undefined) return {};
  if (!record(value))
    throw new CapabilityError("invalid-style", "style must be an object");
  if (value["lang"] !== undefined) {
    assertValidLang(value["lang"]);
  }
  return Object.fromEntries(
    Object.entries(value).filter(
      ([key, item]) => textStyleKeys.has(key) && item !== undefined,
    ),
  );
}

function nativeSlot(value: string): value is NativeColorSlot {
  return (nativeColorSlots as readonly string[]).includes(value);
}

export function isFont(value: unknown): value is Font {
  return record(value) && typeof value["family"] === "string";
}

function normalizeTheme(input: ThemeInput = {}): Theme {
  if (!record(input)) throw new TypeError("Theme must be an object");
  const colors = { ...officeColors, ...input.colors };
  const fonts = {
    ...input.fonts,
    major: input.fonts?.major ?? "Aptos Display",
    minor: input.fonts?.minor ?? "Aptos",
  };
  return Object.freeze({
    name: input.name ?? "Office Theme",
    colors: Object.freeze(colors),
    fonts: Object.freeze(fonts),
  });
}

/** Fill unsupplied theme items using initial Office palette and fonts. */
export function defineTheme(input: ThemeInput = {}): Theme {
  return normalizeTheme(input);
}

export const defaultTheme = defineTheme();

/** Generate darker variant in PowerPoint native palette. */
export function darker(slot: NativeColorSlot, amount: number): ThemeShade {
  return { slot, lumMod: 1 - amount };
}

/** Generate lighter variant in PowerPoint native palette. */
export function lighter(slot: NativeColorSlot, amount: number): ThemeShade {
  return { slot, lumMod: 1 - amount, lumOff: amount };
}

export function isNativeColorSlot(value: string): value is NativeColorSlot {
  return nativeSlot(value);
}

function scopeValue(
  context: ThemeContext,
  name: string,
): { readonly value: ThemeColor; readonly parent?: ThemeContext } | undefined {
  let current: ThemeContext | undefined = context;
  while (current !== undefined) {
    const scope = contextRecords.get(current)?.scope;
    if (scope?.colors !== undefined && Object.hasOwn(scope.colors, name)) {
      return {
        value: scope.colors[name]!,
        ...(contextRecords.get(current)?.parent === undefined
          ? {}
          : { parent: contextRecords.get(current)!.parent! }),
      };
    }
    current = contextRecords.get(current)?.parent;
  }
  return undefined;
}

function scopeFontValue(
  context: ThemeContext,
  name: string,
): { readonly value: FontToken; readonly parent?: ThemeContext } | undefined {
  let current: ThemeContext | undefined = context;
  while (current !== undefined) {
    const scope = contextRecords.get(current)?.scope;
    if (scope?.fonts !== undefined && Object.hasOwn(scope.fonts, name)) {
      return {
        value: scope.fonts[name]!,
        ...(contextRecords.get(current)?.parent === undefined
          ? {}
          : { parent: contextRecords.get(current)!.parent! }),
      };
    }
    current = contextRecords.get(current)?.parent;
  }
  return undefined;
}

function shadeHex(hex: string, lumMod: number, lumOff: number): string {
  const [r, g, b] = [0, 2, 4].map(
    (start) => Number.parseInt(hex.slice(start, start + 2), 16) / 255,
  ) as [number, number, number];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  const delta = max - min;
  const saturation =
    delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  const hue =
    delta === 0
      ? 0
      : max === r
        ? ((g - b) / delta + 6) % 6
        : max === g
          ? (b - r) / delta + 2
          : (r - g) / delta + 4;
  const shaded = Math.min(1, Math.max(0, lightness * lumMod + lumOff));
  const chroma = (1 - Math.abs(2 * shaded - 1)) * saturation;
  const x = chroma * (1 - Math.abs((hue % 2) - 1));
  const m = shaded - chroma / 2;
  const [r1, g1, b1] =
    hue < 1
      ? [chroma, x, 0]
      : hue < 2
        ? [x, chroma, 0]
        : hue < 3
          ? [0, chroma, x]
          : hue < 4
            ? [0, x, chroma]
            : hue < 5
              ? [x, 0, chroma]
              : [chroma, 0, x];
  return [r1, g1, b1]
    .map((channel) =>
      Math.round((channel + m) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

function checkedFraction(value: unknown, label: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < 0 ||
    value > 1
  ) {
    throw new CapabilityError(
      "invalid-color",
      `${label} must be a number between 0 and 1`,
    );
  }
  return value;
}

function resolveColorName(
  raw: string,
  context: ThemeContext,
  depth: number,
): ResolvedColor {
  if (depth > 16) {
    throw new CapabilityError(
      "invalid-color",
      `Circular color token reference: ${raw}`,
    );
  }
  const value = raw.trim();
  if (value.startsWith("#")) {
    const digits = value.slice(1);
    const expanded =
      digits.length === 3 || digits.length === 4
        ? digits
            .split("")
            .map((digit) => digit + digit)
            .join("")
        : digits;
    if (!/^[0-9a-f]{6}([0-9a-f]{2})?$/iu.test(expanded)) {
      throw new CapabilityError("invalid-color", `Invalid color value: ${raw}`);
    }
    const alpha =
      expanded.length === 8
        ? Number.parseInt(expanded.slice(6), 16) / 255
        : undefined;
    return {
      hex: expanded.slice(0, 6).toUpperCase(),
      ...(alpha === undefined || alpha === 1 ? {} : { alpha }),
    };
  }
  if (nativeSlot(value)) {
    const configured = context.theme.colors[value];
    const resolved = resolveColorName(
      configured,
      rootContext(context),
      depth + 1,
    );
    return { ...resolved, slot: value };
  }
  const local = scopeValue(context, value);
  if (local !== undefined) {
    return resolveColorValue(
      local.value,
      local.parent ?? rootContext(context),
      depth + 1,
    );
  }
  const configured = context.theme.colors[value];
  if (configured === undefined) {
    throw new CapabilityError(
      "invalid-color",
      `Unknown color or theme token: ${raw}`,
    );
  }
  // Deck token aliases are resolved at the point of use so they see overrides from the nearest scope.
  return resolveColorValue(configured, context, depth + 1);
}

function rootContext(context: ThemeContext): ThemeContext {
  let current = context;
  while (contextRecords.get(current)?.parent !== undefined) {
    current = contextRecords.get(current)!.parent!;
  }
  return current;
}

function resolveColorValue(
  value: ThemeColor,
  context: ThemeContext,
  depth: number,
): ResolvedColor {
  if (typeof value === "string") return resolveColorName(value, context, depth);
  if (
    !record(value) ||
    !nativeSlot(String(value["slot"])) ||
    typeof value["lumMod"] !== "number" ||
    !Number.isFinite(value["lumMod"])
  ) {
    throw new CapabilityError("invalid-color", "Invalid theme shade value");
  }
  const lumMod = checkedFraction(value["lumMod"], "lumMod");
  const lumOff =
    value["lumOff"] === undefined
      ? undefined
      : checkedFraction(value["lumOff"], "lumOff");
  const slot = value["slot"];
  const base = resolveColorName(slot, rootContext(context), depth + 1);
  return {
    hex: shadeHex(base.hex, lumMod, lumOff ?? 0),
    slot,
    lumMod,
    ...(lumOff === undefined ? {} : { lumOff }),
  };
}

function createContext(
  theme: Theme,
  parent: ThemeContext | undefined,
  scope: ThemeScope | undefined,
  textStyle: TextStyle,
): ThemeContext {
  let context: ThemeContext;
  context = Object.freeze({
    theme,
    textStyle: Object.freeze(textStyle),
    with(nextScope?: ThemeScope, style?: TextStyle) {
      let checkedScope: ThemeScope | undefined;
      if (nextScope !== undefined) {
        if (!record(nextScope)) {
          throw new CapabilityError(
            "invalid-theme-scope",
            "theme must be an object",
          );
        }
        const unknown = Object.keys(nextScope).find(
          (key) => key !== "colors" && key !== "fonts",
        );
        if (unknown !== undefined) {
          throw new CapabilityError(
            "invalid-theme-scope",
            `Local theme scope contains unknown field ${unknown}`,
          );
        }
        const colors = nextScope.colors;
        if (colors !== undefined && !record(colors)) {
          throw new CapabilityError(
            "invalid-theme-scope",
            "theme.colors must be an object",
          );
        }
        for (const name of Object.keys(colors ?? {})) {
          if (nativeSlot(name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `Local theme scope cannot redefine native color slot ${name}`,
            );
          }
          if (!Object.hasOwn(theme.colors, name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `Theme has no color token ${name}`,
            );
          }
        }
        const fonts = nextScope.fonts;
        if (fonts !== undefined && !record(fonts)) {
          throw new CapabilityError(
            "invalid-theme-scope",
            "theme.fonts must be an object",
          );
        }
        for (const [name, font] of Object.entries(fonts ?? {})) {
          if (!Object.hasOwn(theme.fonts, name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `Theme has no font token ${name}`,
            );
          }
          if (
            !isFont(font) &&
            (typeof font !== "string" || font.trim() === "")
          ) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `Font token ${name} must be replaced with a Font value or non-empty family string`,
            );
          }
        }
        checkedScope = Object.freeze({
          ...(colors === undefined
            ? {}
            : { colors: Object.freeze({ ...colors }) }),
          ...(fonts === undefined
            ? {}
            : { fonts: Object.freeze({ ...fonts }) }),
        });
      }
      const nextStyle = { ...context.textStyle };
      if (style !== undefined) {
        if (!record(style)) {
          throw new CapabilityError("invalid-style", "style must be an object");
        }
        if (style.lang !== undefined) {
          assertValidLang(style.lang);
        }
        for (const [key, value] of Object.entries(style)) {
          if (textStyleKeys.has(key) && value !== undefined) {
            (nextStyle as Record<string, unknown>)[key] = value;
          }
        }
      }
      return createContext(theme, context, checkedScope, nextStyle);
    },
    resolveColor(value: ColorValue) {
      let raw: string | ThemeShade;
      let alpha: number | undefined;
      if (typeof value === "string" || (record(value) && "slot" in value)) {
        raw = value;
      } else if (record(value) && "color" in value) {
        raw = value["color"];
        alpha =
          value["alpha"] === undefined
            ? undefined
            : checkedFraction(value["alpha"], "color alpha");
      } else {
        throw new CapabilityError(
          "invalid-color",
          "Invalid color value format",
        );
      }
      const resolved =
        typeof raw === "string"
          ? resolveColorName(raw, context, 0)
          : resolveColorValue(raw, context, 0);
      if (alpha === undefined) return resolved;
      return {
        ...resolved,
        alpha: (resolved.alpha ?? 1) * alpha,
      };
    },
    resolveFont(value: FontReference) {
      if (isFont(value)) return { family: value.family, font: value };
      if (typeof value !== "string" || value.trim() === "") {
        throw new CapabilityError(
          "invalid-font-family",
          "fontFamily cannot be empty",
        );
      }
      const local = scopeFontValue(context, value);
      const token =
        local?.value ??
        (Object.hasOwn(theme.fonts, value) ? theme.fonts[value]! : value);
      const themeFontName: ResolvedThemeFont["themeFont"] =
        value === "major" || value === "minor" ? value : undefined;
      const globalThemeFont =
        themeFontName === undefined ? undefined : theme.fonts[themeFontName];
      const globalFamily = isFont(globalThemeFont)
        ? globalThemeFont.family
        : globalThemeFont;
      if (isFont(token)) {
        return {
          family: token.family,
          font: token,
          ...(themeFontName === undefined || token.family !== globalFamily
            ? {}
            : { themeFont: themeFontName }),
        };
      }
      return {
        family: token,
        ...(themeFontName === undefined || token !== globalFamily
          ? {}
          : { themeFont: themeFontName }),
      };
    },
  });
  contextRecords.set(context, {
    ...(parent === undefined ? {} : { parent }),
    ...(scope === undefined ? {} : { scope }),
  });
  return context;
}

/** Create a theme chain unique to a single build; local scopes do not mutate the shared theme object. */
export function createThemeContext(
  input?: ThemeInput | Theme,
  style?: TextStyle,
): ThemeContext {
  if (style !== undefined) {
    const unknown = Object.keys(style).find((key) => !textStyleKeys.has(key));
    if (unknown !== undefined) {
      throw new CapabilityError(
        "invalid-style",
        `Deck.style contains unsupported text property ${unknown}`,
      );
    }
  }
  return createContext(
    normalizeTheme(input),
    undefined,
    undefined,
    textStyleFrom(style),
  );
}

/** Write native theme colors as schemeClr and literal colors as srgbClr. */
export function resolvedColorXml(value: ResolvedColor): string {
  const children = [
    value.lumMod === undefined
      ? ""
      : `<a:lumMod val="${Math.round(value.lumMod * 100000)}"/>`,
    value.lumOff === undefined
      ? ""
      : `<a:lumOff val="${Math.round(value.lumOff * 100000)}"/>`,
    value.alpha === undefined
      ? ""
      : `<a:alpha val="${Math.round(value.alpha * 100000)}"/>`,
  ].join("");
  const element = value.slot === undefined ? "srgbClr" : "schemeClr";
  const color =
    value.slot === undefined ? value.hex : nativeSchemeReference[value.slot];
  return children === ""
    ? `<a:${element} val="${color}"/>`
    : `<a:${element} val="${color}">${children}</a:${element}>`;
}
