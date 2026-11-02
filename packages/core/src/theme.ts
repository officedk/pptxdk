import { CapabilityError } from "#/diagnostics.ts";

export type FontFaceStyle = "normal" | "italic";

/** 字体文件由调用方提供，可用于度量或按构建选项嵌入。 */
export type FontFace = {
  readonly weight?: number;
  readonly style?: FontFaceStyle;
  readonly data?: Uint8Array;
};

/** 可在文稿主题和组件之间引用的字体族及其显式度量资源。 */
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

/** 文字对象和容器共用的可继承文字样式字段。 */
export type TextStyle = {
  readonly fontFamily?: FontReference;
  readonly fontSize?: FontSize;
  readonly fontWeight?: number | "normal" | "bold";
  readonly fontStyle?: FontFaceStyle;
  readonly color?: ColorValue;
  readonly lineHeight?: FontSize;
  readonly verticalAlign?: "top" | "middle" | "bottom";
};

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

/** 用 DrawingML 明度变换表示的原生主题颜色。 */
export type ThemeShade = {
  readonly slot: NativeColorSlot;
  readonly lumMod: number;
  readonly lumOff?: number;
};

export type ThemeColor = string | ThemeShade;

/** 颜色入口可直接给颜色、主题槽或主题令牌，也可附加透明度。 */
export type ColorValue =
  | string
  | ThemeShade
  | { readonly color: string | ThemeShade; readonly alpha?: number };

/** 对颜色入口应用的最终主题值。 */
export type ResolvedColor = {
  readonly hex: string;
  readonly alpha?: number;
  readonly slot?: NativeColorSlot;
  readonly lumMod?: number;
  readonly lumOff?: number;
};

export type FontToken = FontReference;

/** 文稿级主题输入；major/minor 字体同时写入 PPTX 原生主题。 */
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

/** 页面或容器上的局部主题只能替换文稿中已有的颜色与字体令牌。 */
export type ThemeScope = {
  readonly colors?: Readonly<Record<string, ThemeColor>>;
  readonly fonts?: Readonly<Record<string, FontToken>>;
};

export type ResolvedThemeFont = {
  readonly family: string;
  readonly font?: Font;
  readonly themeFont?: "major" | "minor";
};

/** 单次构建中的主题、局部作用域与累计容器文字默认值。 */
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
]);

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** 从包含布局字段的容器 style 中取出可继承文字字段。 */
export function textStyleFrom(value: unknown): TextStyle {
  if (value === undefined) return {};
  if (!record(value))
    throw new CapabilityError("invalid-style", "style 必须是对象");
  return Object.fromEntries(
    Object.entries(value).filter(
      ([key, item]) => textStyleKeys.has(key) && item !== undefined,
    ),
  ) as TextStyle;
}

function nativeSlot(value: string): value is NativeColorSlot {
  return (nativeColorSlots as readonly string[]).includes(value);
}

export function isFont(value: unknown): value is Font {
  return record(value) && typeof value["family"] === "string";
}

function normalizeTheme(input: ThemeInput = {}): Theme {
  if (!record(input)) throw new TypeError("Theme 必须是对象");
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
  }) as Theme;
}

/** 使用首版 Office 配色和字体填充未提供的主题项。 */
export function defineTheme(input: ThemeInput = {}): Theme {
  return normalizeTheme(input);
}

export const defaultTheme = defineTheme();

/** 生成 PowerPoint 原生调色板中的较暗变体。 */
export function darker(slot: NativeColorSlot, amount: number): ThemeShade {
  return { slot, lumMod: 1 - amount };
}

/** 生成 PowerPoint 原生调色板中的较亮变体。 */
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
    throw new CapabilityError("invalid-color", `${label} 必须是 0 到 1 的数值`);
  }
  return value;
}

function resolveColorName(
  raw: string,
  context: ThemeContext,
  depth: number,
): ResolvedColor {
  if (depth > 16) {
    throw new CapabilityError("invalid-color", `颜色令牌循环引用：${raw}`);
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
      throw new CapabilityError("invalid-color", `颜色值 ${raw} 无效`);
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
    throw new CapabilityError("invalid-color", `未知颜色或主题令牌 ${raw}`);
  }
  return resolveColorValue(configured, rootContext(context), depth + 1);
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
    throw new CapabilityError("invalid-color", "主题明暗值无效");
  }
  const lumMod = checkedFraction(value["lumMod"], "lumMod");
  const lumOff =
    value["lumOff"] === undefined
      ? undefined
      : checkedFraction(value["lumOff"], "lumOff");
  const slot = value["slot"] as NativeColorSlot;
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
          throw new CapabilityError("invalid-theme-scope", "theme 必须是对象");
        }
        const unknown = Object.keys(nextScope).find(
          (key) => key !== "colors" && key !== "fonts",
        );
        if (unknown !== undefined) {
          throw new CapabilityError(
            "invalid-theme-scope",
            `局部主题含未实现字段 ${unknown}`,
          );
        }
        const colors = nextScope.colors;
        if (colors !== undefined && !record(colors)) {
          throw new CapabilityError(
            "invalid-theme-scope",
            "theme.colors 必须是对象",
          );
        }
        for (const name of Object.keys(colors ?? {})) {
          if (nativeSlot(name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `局部主题不能重定义原生颜色槽 ${name}`,
            );
          }
          if (!Object.hasOwn(theme.colors, name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `主题中没有颜色令牌 ${name}`,
            );
          }
        }
        const fonts = nextScope.fonts;
        if (fonts !== undefined && !record(fonts)) {
          throw new CapabilityError(
            "invalid-theme-scope",
            "theme.fonts 必须是对象",
          );
        }
        for (const [name, font] of Object.entries(fonts ?? {})) {
          if (!Object.hasOwn(theme.fonts, name)) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `主题中没有字体令牌 ${name}`,
            );
          }
          if (
            !isFont(font) &&
            (typeof font !== "string" || font.trim() === "")
          ) {
            throw new CapabilityError(
              "invalid-theme-scope",
              `字体令牌 ${name} 必须替换为字体值或非空字族名`,
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
          throw new CapabilityError("invalid-style", "style 必须是对象");
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
        raw = value as string | ThemeShade;
      } else if (record(value) && "color" in value) {
        raw = value["color"] as string | ThemeShade;
        alpha =
          value["alpha"] === undefined
            ? undefined
            : checkedFraction(value["alpha"], "颜色透明度");
      } else {
        throw new CapabilityError("invalid-color", "颜色值格式无效");
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
        throw new CapabilityError("invalid-font-family", "fontFamily 不能为空");
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

/** 创建一次构建独有的主题链，局部作用域不会写回共享主题对象。 */
export function createThemeContext(
  input?: ThemeInput | Theme,
  style?: TextStyle,
): ThemeContext {
  if (style !== undefined) {
    const unknown = Object.keys(style).find((key) => !textStyleKeys.has(key));
    if (unknown !== undefined) {
      throw new CapabilityError(
        "invalid-style",
        `Deck.style 含未实现文字属性 ${unknown}`,
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

/** 将原生主题颜色写成 schemeClr，字面颜色写成 srgbClr。 */
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
