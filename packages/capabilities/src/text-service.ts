import { createHash } from "node:crypto";

import {
  CapabilityError,
  pxToEmu,
  resolvedColorXml,
  xmlEscape,
  type ColorValue,
  type Font,
  type FontFaceStyle,
  type FontReference,
  type FontSize,
  type ResolvedColor,
  type Shadow,
  type TextAlignment,
  type TextStyle,
  type ThemeContext,
} from "@react-pptx/core";
import * as fontkit from "fontkit";
// @ts-expect-error -- linebreak 未提供类型，下面只声明所用的断点接口。
import LineBreakerImplementation from "linebreak";

import type { FontEmbeddingUse } from "./font-embedding.ts";
import { textShadowXml } from "./visual.ts";
export type {
  FontReference,
  FontSize,
  TextAlignment,
  TextStyle,
} from "@react-pptx/core";

export type TextList =
  | {
      readonly type: "bullet";
      readonly character?: string;
      readonly level?: number;
    }
  | {
      readonly type: "number";
      readonly style?: "arabicPeriod";
      readonly startAt?: number;
      readonly level?: number;
    };

export type ParagraphStyle = Omit<TextStyle, "verticalAlign"> & {
  readonly align?: TextAlignment;
  readonly spacingAfter?: number;
  readonly list?: TextList;
};

export type TextResourceClaim = {
  readonly kind: "font-face";
  readonly key: string;
  readonly identity: string;
  /** 实际输出给该字体样式的文本字符，供后续字体处理继续复用。 */
  readonly characters: string;
};

export type TextServiceResult = {
  readonly xml: string;
  readonly text: string;
  readonly paragraphCount: number;
  readonly resourceClaims: readonly TextResourceClaim[];
  readonly fontUses: readonly FontEmbeddingUse[];
  readonly measurement?: { readonly width: number; readonly height: number };
  readonly warnings: readonly {
    readonly code: string;
    readonly message: string;
  }[];
};

export type TextLayoutMeasurement = {
  readonly width: number;
  readonly height: number;
  readonly minContentWidth: number;
  readonly minContentHeight: number;
  readonly baseline?: number;
};

type HostText = { readonly type: "#text"; readonly text: string };
type HostElement = {
  readonly type: `rp:${string}`;
  readonly props: Readonly<Record<string, unknown>>;
  readonly children: readonly unknown[];
};
type TextChild = HostText | HostElement;
type RunStyle = {
  readonly fontFamily?: FontReference | undefined;
  readonly fontSize?: FontSize | undefined;
  readonly fontWeight?: TextStyle["fontWeight"] | undefined;
  readonly fontStyle?: FontFaceStyle | undefined;
  readonly color?: ColorValue | ResolvedColor | undefined;
  readonly themeFont?: "major" | "minor" | undefined;
  readonly letterSpacing?: number | undefined;
  readonly textTransform?: TextStyle["textTransform"] | undefined;
  readonly textDecoration?: TextStyle["textDecoration"] | undefined;
  readonly textShadow?: Shadow | "none" | undefined;
};
type Run = { readonly text: string; readonly style: RunStyle };
type Paragraph = {
  readonly runs: readonly Run[];
  readonly style: ParagraphStyle;
  readonly endStyle: RunStyle;
};
type NormalizedFont = {
  readonly family: string;
  readonly weight: number;
  readonly style: FontFaceStyle;
  readonly substitute: boolean;
  readonly embeddable: boolean;
  readonly hasResource: boolean;
  readonly identity: string;
  readonly data?: Uint8Array;
};

type TextLayoutRequest = {
  readonly width: number | "max-content";
  readonly height?: number;
  readonly contentSized?: boolean;
  readonly insets?: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
};

type TextBodyDestination = "text-box" | "table-cell";

type MeasureFace = {
  readonly font: fontkit.Font;
  readonly identity: string;
};

type MeasurePiece = {
  readonly text: string;
  readonly style: RunStyle;
  readonly start: number;
  readonly end: number;
  readonly face: MeasureFace;
};

type MeasuredParagraph = {
  readonly lineCount: number;
  readonly width: number;
  readonly lineHeight: number;
  readonly spacingAfter: number;
};

type LineBreakerConstructor = new (text: string) => {
  nextBreak(): { readonly position: number; readonly required: boolean } | null;
};

const LineBreaker = LineBreakerImplementation as LineBreakerConstructor;
const graphemes = new Intl.Segmenter(undefined, { granularity: "grapheme" });
const measureFaces = new Map<string, fontkit.Font>();

const defaultTextStyle: TextStyle = Object.freeze({
  fontFamily: "minor",
  fontSize: 24,
  fontWeight: "normal",
  fontStyle: "normal",
  color: "dk1",
});

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function escapedXml(value: string): string {
  try {
    return xmlEscape(value);
  } catch {
    throw new CapabilityError(
      "invalid-xml-character",
      "文字内容或字体名称含 XML 1.0 不允许的字符",
    );
  }
}

function hostChild(value: unknown): value is TextChild {
  if (!record(value) || typeof value["type"] !== "string") return false;
  if (value["type"] === "#text") return typeof value["text"] === "string";
  return (
    value["type"].startsWith("rp:") &&
    record(value["props"]) &&
    Array.isArray(value["children"])
  );
}

function styleRecord(value: unknown, label: string): Record<string, unknown> {
  if (value === undefined) return {};
  if (!record(value)) {
    throw new CapabilityError("invalid-style", `${label} 必须是对象`);
  }
  return value;
}

function onlyKnownKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
  label: string,
): void {
  const unknown = Object.keys(value).find((key) => !allowed.includes(key));
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-style-property",
      `${label} 含未实现样式 ${unknown}`,
    );
  }
}

function positiveLength(value: unknown, label: string): number {
  let pixels: number;
  if (typeof value === "number") pixels = value;
  else if (typeof value === "string") {
    const match = /^(\d+(?:\.\d+)?)(px|pt)$/u.exec(value);
    if (match === null) {
      throw new CapabilityError(
        "invalid-text-length",
        `${label} 必须是正像素数值、px 或 pt 长度`,
      );
    }
    const amount = Number(match[1]);
    pixels = match[2] === "pt" ? (amount * 4) / 3 : amount;
  } else {
    throw new CapabilityError(
      "invalid-text-length",
      `${label} 必须是正像素数值、px 或 pt 长度`,
    );
  }
  if (!Number.isFinite(pixels) || pixels <= 0) {
    throw new CapabilityError("invalid-text-length", `${label} 必须大于 0`);
  }
  return pixels;
}

function centipoints(value: number): number {
  return Math.round(value * 100);
}

function spacingCentipoints(
  pixels: number,
  label: "lineHeight" | "spacingAfter",
): number {
  const value = Math.round(pixels * 75);
  if (value < 0 || value > 158_400) {
    throw new CapabilityError(
      label === "lineHeight"
        ? "invalid-text-length"
        : "invalid-paragraph-spacing",
      `${label} 换算后必须在 0 到 1584pt 的 DrawingML 范围内`,
    );
  }
  return value;
}

function letterSpacingCentipoints(value: number): number {
  const points = Math.round(value * 75);
  if (!Number.isFinite(value) || points < -400_000 || points > 400_000) {
    throw new CapabilityError(
      "invalid-letter-spacing",
      "letterSpacing 换算后必须在 -4000 到 4000pt 的 DrawingML 范围内",
    );
  }
  return points;
}

function fontSizePoints(value: FontSize | undefined): number {
  let points: number;
  if (value === undefined) {
    points = positiveLength(defaultTextStyle.fontSize, "字号") * 0.75;
  } else if (typeof value === "string" && value.endsWith("pt")) {
    points = Number(value.slice(0, -2));
    if (!Number.isFinite(points) || points <= 0) {
      throw new CapabilityError("invalid-font-size", "fontSize 必须大于 0");
    }
  } else {
    points = positiveLength(value, "fontSize") * 0.75;
  }
  const serialized = centipoints(points);
  if (serialized < 100 || serialized > 400_000) {
    throw new CapabilityError(
      "invalid-font-size",
      "fontSize 换算后必须在 1 到 4000pt 的 DrawingML 范围内",
    );
  }
  return serialized / 100;
}

function weightValue(value: TextStyle["fontWeight"]): number {
  if (value === undefined || value === "normal") return 400;
  if (value === "bold") return 700;
  if (
    typeof value !== "number" ||
    !Number.isInteger(value) ||
    value < 100 ||
    value > 900 ||
    value % 100 !== 0
  ) {
    throw new CapabilityError(
      "invalid-font-weight",
      "fontWeight 必须是 normal、bold 或 100 到 900 间的百位整数",
    );
  }
  return value;
}

function fontFamilyValue(value: FontReference | undefined): Font {
  if (value === undefined) return { family: "Aptos" };
  if (typeof value === "string") {
    if (value.trim() === "") {
      throw new CapabilityError("invalid-font-family", "fontFamily 不能为空");
    }
    return { family: value };
  }
  if (!record(value) || typeof value["family"] !== "string") {
    throw new CapabilityError("invalid-font-family", "fontFamily 必须是字体值");
  }
  if (value["family"].trim() === "") {
    throw new CapabilityError("invalid-font-family", "Font.family 不能为空");
  }
  onlyKnownKeys(value, ["family", "faces", "substitute", "embeddable"], "Font");
  if (value["faces"] !== undefined && !Array.isArray(value["faces"])) {
    throw new CapabilityError("invalid-font-family", "Font.faces 必须是数组");
  }
  if (
    value["substitute"] !== undefined &&
    typeof value["substitute"] !== "boolean"
  ) {
    throw new CapabilityError(
      "invalid-font-family",
      "Font.substitute 必须是布尔值",
    );
  }
  if (
    value["embeddable"] !== undefined &&
    typeof value["embeddable"] !== "boolean"
  ) {
    throw new CapabilityError(
      "invalid-font-family",
      "Font.embeddable 必须是布尔值",
    );
  }
  return value as Font;
}

function normalizedFonts(style: RunStyle): NormalizedFont[] {
  const family = fontFamilyValue(style.fontFamily);
  const weight = weightValue(style.fontWeight);
  const faceStyle = style.fontStyle ?? "normal";
  if (faceStyle !== "normal" && faceStyle !== "italic") {
    throw new CapabilityError(
      "invalid-font-style",
      "fontStyle 只能是 normal 或 italic",
    );
  }
  const faces = family.faces ?? [];
  for (const face of faces) {
    onlyKnownKeys(face, ["weight", "style", "data"], "FontFace");
    const faceWeight = face.weight ?? 400;
    const styleValue = face.style ?? "normal";
    if (
      !Number.isInteger(faceWeight) ||
      faceWeight < 100 ||
      faceWeight > 900 ||
      faceWeight % 100 !== 0 ||
      (styleValue !== "normal" && styleValue !== "italic") ||
      (face.data !== undefined &&
        (!(face.data instanceof Uint8Array) || face.data.byteLength === 0))
    ) {
      throw new CapabilityError(
        "invalid-font-face",
        "Font.faces 含无效字重、样式或资源",
      );
    }
  }
  const matchingFaces = faces.filter(
    (candidate) =>
      (candidate.weight ?? 400) === weight &&
      (candidate.style ?? "normal") === faceStyle,
  );
  const providedFaces = matchingFaces.filter(
    (candidate) => candidate.data !== undefined,
  );
  if (providedFaces.length === 0) {
    return [
      {
        family: family.family,
        weight,
        style: faceStyle,
        substitute: family.substitute ?? false,
        embeddable: family.embeddable ?? true,
        hasResource: false,
        identity: "missing",
      },
    ];
  }
  return providedFaces.map((face) => ({
    family: family.family,
    weight,
    style: faceStyle,
    substitute: family.substitute ?? false,
    embeddable: family.embeddable ?? true,
    hasResource: true,
    identity: createHash("sha256").update(face.data!).digest("hex"),
    data: face.data!,
  }));
}

function validateRunStyle(style: RunStyle): RunStyle {
  fontSizePoints(style.fontSize);
  weightValue(style.fontWeight);
  normalizedFonts(style);
  if (style.letterSpacing !== undefined)
    letterSpacingCentipoints(style.letterSpacing);
  if (
    style.textTransform !== undefined &&
    !["none", "uppercase"].includes(style.textTransform)
  ) {
    throw new CapabilityError(
      "invalid-text-transform",
      "textTransform 只能是 none 或 uppercase",
    );
  }
  if (
    style.textDecoration !== undefined &&
    !["none", "underline"].includes(style.textDecoration)
  ) {
    throw new CapabilityError(
      "invalid-text-decoration",
      "textDecoration 只能是 none 或 underline",
    );
  }
  return style;
}

function readRunStyle(value: unknown): RunStyle {
  const props = styleRecord(value, "Span.style");
  onlyKnownKeys(
    props,
    [
      "fontFamily",
      "fontSize",
      "fontWeight",
      "fontStyle",
      "color",
      "letterSpacing",
      "textTransform",
      "textDecoration",
      "textShadow",
    ],
    "Span.style",
  );
  return validateRunStyle(props as RunStyle);
}

function mergeRunStyle(
  parent: RunStyle,
  child: RunStyle,
  theme: ThemeContext,
): RunStyle {
  const declared: Record<string, unknown> = Object.fromEntries(
    Object.entries(child).filter(([, value]) => value !== undefined),
  );
  if (child.fontFamily !== undefined) {
    const font = theme.resolveFont(child.fontFamily);
    declared["fontFamily"] = font.font ?? font.family;
    if (child.themeFont === undefined && font.themeFont === undefined) {
      delete declared["themeFont"];
    } else {
      declared["themeFont"] = child.themeFont ?? font.themeFont;
    }
  }
  if (child.color !== undefined && !isResolvedColor(child.color)) {
    declared["color"] = theme.resolveColor(child.color as ColorValue);
  }
  // P 和 Span 未声明的字段会是 undefined，不能覆盖继承的父级样式。
  const merged = { ...parent, ...declared } as RunStyle;
  if (child.textShadow === "none")
    delete (merged as { textShadow?: Shadow | "none" }).textShadow;
  if (child.fontFamily !== undefined && declared["themeFont"] === undefined) {
    const { themeFont: _themeFont, ...withoutThemeFont } = merged;
    return validateRunStyle(withoutThemeFont);
  }
  return validateRunStyle(merged);
}

function isResolvedColor(value: unknown): value is ResolvedColor {
  return record(value) && typeof value["hex"] === "string";
}

function outputBold(font: NormalizedFont): boolean {
  const fallback = font.weight >= 600;
  if (!font.hasResource || font.substitute) return fallback;
  try {
    // 字重数值不能替代字体文件声明的同族粗体样式链接。
    return measureFace(font).font["OS/2"].fsSelection.bold;
  } catch {
    return fallback;
  }
}

type ResolvedTextStyle = TextStyle & {
  readonly themeFont?: "major" | "minor";
};

function validateParagraphDefaults(
  style: Pick<ParagraphStyle, "align" | "spacingAfter">,
): void {
  if (
    style.align !== undefined &&
    !["left", "center", "right", "justify"].includes(style.align)
  ) {
    throw new CapabilityError(
      "invalid-text-alignment",
      "align 只能是 left、center、right 或 justify",
    );
  }
  if (
    style.spacingAfter !== undefined &&
    (!Number.isFinite(style.spacingAfter) || style.spacingAfter < 0)
  ) {
    throw new CapabilityError(
      "invalid-paragraph-spacing",
      "spacingAfter 必须是有限非负像素数值",
    );
  }
  if (style.spacingAfter !== undefined) {
    spacingCentipoints(style.spacingAfter, "spacingAfter");
  }
}

function readTextStyle(
  value: unknown,
  label: string,
  theme: ThemeContext,
): ResolvedTextStyle {
  const props = styleRecord(value, label);
  onlyKnownKeys(
    props,
    [
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
    ],
    label,
  );
  const style = {
    ...defaultTextStyle,
    ...theme.textStyle,
    ...props,
  } as TextStyle;
  if (style.textShadow === "none")
    delete (style as { textShadow?: Shadow | "none" }).textShadow;
  if (
    style.verticalAlign !== undefined &&
    !["top", "middle", "bottom"].includes(style.verticalAlign)
  ) {
    throw new CapabilityError(
      "invalid-vertical-align",
      "verticalAlign 只能是 top、middle 或 bottom",
    );
  }
  validateParagraphDefaults(style);
  const runStyle: RunStyle = {
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    color: style.color,
    themeFont: (style as ResolvedTextStyle).themeFont,
    letterSpacing: style.letterSpacing,
    textTransform: style.textTransform,
    textDecoration: style.textDecoration,
    textShadow: style.textShadow,
  };
  const font = theme.resolveFont(
    runStyle.fontFamily ?? defaultTextStyle.fontFamily!,
  );
  if (runStyle.color !== undefined && !isResolvedColor(runStyle.color)) {
    theme.resolveColor(runStyle.color as ColorValue);
  }
  validateRunStyle({
    ...runStyle,
    fontFamily: font.font ?? font.family,
  });
  if (style.lineHeight !== undefined)
    spacingCentipoints(
      positiveLength(style.lineHeight, "lineHeight"),
      "lineHeight",
    );
  return {
    ...style,
    fontFamily: font.font ?? font.family,
    ...(font.themeFont === undefined ? {} : { themeFont: font.themeFont }),
  };
}

function paragraphStyle(value: unknown): ParagraphStyle {
  const props = styleRecord(value, "P.style");
  onlyKnownKeys(
    props,
    [
      "fontFamily",
      "fontSize",
      "fontWeight",
      "fontStyle",
      "color",
      "lineHeight",
      "align",
      "spacingAfter",
      "letterSpacing",
      "textTransform",
      "textDecoration",
      "textShadow",
      "list",
    ],
    "P.style",
  );
  const style = props as ParagraphStyle;
  validateParagraphDefaults(style);
  if (style.lineHeight !== undefined)
    spacingCentipoints(
      positiveLength(style.lineHeight, "lineHeight"),
      "lineHeight",
    );
  if (style.list !== undefined) validateList(style.list);
  return style;
}

function validateList(value: TextList): void {
  if (
    !record(value) ||
    (value["type"] !== "bullet" && value["type"] !== "number")
  ) {
    throw new CapabilityError(
      "invalid-list",
      "list.type 必须是 bullet 或 number",
    );
  }
  onlyKnownKeys(
    value,
    value["type"] === "bullet"
      ? ["type", "character", "level"]
      : ["type", "style", "startAt", "level"],
    "list",
  );
  const level = value["level"] ?? 0;
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new CapabilityError(
      "invalid-list",
      "list.level 必须是 0 到 8 的整数",
    );
  }
  if (value["type"] === "bullet") {
    const character = value["character"] ?? "•";
    const codePoint =
      typeof character === "string" ? character.codePointAt(0) : undefined;
    if (
      typeof character !== "string" ||
      codePoint === undefined ||
      String.fromCodePoint(codePoint) !== character
    ) {
      throw new CapabilityError(
        "invalid-list",
        "项目符号必须是单个 Unicode 字符",
      );
    }
    try {
      xmlEscape(character);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "项目符号含 XML 1.0 不允许的字符",
      );
    }
  } else {
    if (value["style"] !== undefined && value["style"] !== "arabicPeriod") {
      throw new CapabilityError(
        "invalid-list",
        "编号样式目前只支持 arabicPeriod",
      );
    }
    const startAt = value["startAt"] ?? 1;
    if (!Number.isInteger(startAt) || startAt < 1 || startAt > 32767) {
      throw new CapabilityError(
        "invalid-list",
        "编号起始值必须是 1 到 32767 的整数",
      );
    }
  }
}

function collectRuns(
  children: readonly unknown[],
  style: RunStyle,
  runs: Run[],
  theme: ThemeContext,
): void {
  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError("invalid-text-child", "文字内容树包含无效节点");
    }
    if (child.type === "#text") {
      if (child.text !== "") runs.push({ text: child.text, style });
      continue;
    }
    if (child.type !== "rp:span") {
      throw new CapabilityError(
        "unsupported-text-child",
        `${child.type} 不能作为文字段落或 run 使用`,
      );
    }
    const unknown = Object.keys(child.props).find((key) => key !== "style");
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Span 含未实现参数 ${unknown}`,
      );
    }
    const spanStyle = readRunStyle(child.props["style"]);
    const nextStyle = mergeRunStyle(style, spanStyle, theme);
    collectRuns(child.children, nextStyle, runs, theme);
  }
}

function runPropertiesXml(
  style: RunStyle,
  tag: "a:rPr" | "a:defRPr",
  theme: ThemeContext,
): string {
  const font = normalizedFonts(style)[0]!;
  const points = fontSizePoints(style.fontSize);
  const bold = outputBold(font);
  const color = isResolvedColor(style.color)
    ? style.color
    : theme.resolveColor(style.color ?? "dk1");
  const typeface =
    style.themeFont === "major"
      ? "+mj-lt"
      : style.themeFont === "minor"
        ? "+mn-lt"
        : font.family;
  const attributes = [
    `lang="en-US"`,
    `sz="${centipoints(points)}"`,
    `b="${bold ? 1 : 0}"`,
    `i="${font.style === "italic" ? 1 : 0}"`,
    ...(style.textDecoration === "underline" ? ['u="sng"'] : []),
    ...(style.textTransform === "uppercase" ? ['cap="all"'] : []),
    ...(style.letterSpacing === undefined || style.letterSpacing === 0
      ? []
      : [`spc="${letterSpacingCentipoints(style.letterSpacing)}"`]),
  ].join(" ");
  return `<${tag} ${attributes}><a:solidFill>${resolvedColorXml(color)}</a:solidFill>${textShadowXml(theme, style.textShadow)}<a:latin typeface="${escapedXml(typeface)}"/></${tag}>`;
}

function displayedText(run: Run): string {
  return run.style.textTransform === "uppercase"
    ? run.text.toUpperCase()
    : run.text;
}

function runXml(run: Run, theme: ThemeContext): string {
  const text = escapedXml(run.text);
  return `<a:r>${runPropertiesXml(run.style, "a:rPr", theme)}<a:t xml:space="preserve">${text}</a:t></a:r>`;
}

function textResourceClaims(runs: readonly Run[]): {
  readonly resourceClaims: TextResourceClaim[];
  readonly fontUses: FontEmbeddingUse[];
  readonly warnings: { readonly code: string; readonly message: string }[];
} {
  const used = new Map<
    string,
    {
      readonly fonts: NormalizedFont[];
      readonly fontPolicies: Map<string, NormalizedFont>;
      readonly text: string[];
    }
  >();
  for (const run of runs) {
    for (const font of normalizedFonts(run.style)) {
      const key = `${font.family.toLocaleLowerCase("en-US")}\u0000${font.weight}\u0000${font.style}`;
      const entry = used.get(key) ?? {
        fonts: [],
        fontPolicies: new Map<string, NormalizedFont>(),
        text: [],
      };
      if (!entry.fonts.some(({ identity }) => identity === font.identity)) {
        entry.fonts.push(font);
      }
      // 字节资源可以合并，但每种实际使用的嵌入策略都须独立校验。
      const policyKey = `${font.identity}\u0000${font.substitute ? 1 : 0}\u0000${font.embeddable ? 1 : 0}`;
      entry.fontPolicies.set(policyKey, font);
      entry.text.push(displayedText(run));
      used.set(key, entry);
    }
  }
  const resourceClaims: TextResourceClaim[] = [];
  const fontUses: FontEmbeddingUse[] = [];
  const warnings: { code: string; message: string }[] = [];
  for (const [key, entry] of used) {
    for (const font of entry.fontPolicies.values()) {
      fontUses.push({
        family: font.family,
        weight: font.weight,
        style: font.style,
        identity: font.identity,
        ...(font.data === undefined ? {} : { bytes: font.data }),
        substitute: font.substitute,
        embeddable: font.embeddable,
        bold: outputBold(font),
        italic: font.style === "italic",
      });
    }
    for (const font of entry.fonts) {
      if (!font.hasResource) {
        warnings.push({
          code: "font-measurement-unavailable",
          message: `未提供 ${font.family} ${font.weight} ${font.style} 的显式字体资源；固定区域文字不会核对实际换行或装入情况`,
        });
      } else {
        resourceClaims.push({
          kind: "font-face",
          key,
          identity: font.identity,
          characters: entry.text.join(""),
        });
      }
    }
  }
  return { resourceClaims, fontUses, warnings };
}

function alignmentValue(value: TextAlignment | undefined): string | undefined {
  if (value === undefined) return undefined;
  return { left: "l", center: "ctr", right: "r", justify: "just" }[value];
}

function paragraphXml(paragraph: Paragraph, theme: ThemeContext): string {
  const { style, runs } = paragraph;
  const attributes: string[] = [];
  const align = alignmentValue(style.align);
  if (align !== undefined) attributes.push(`algn="${align}"`);
  const level = style.list?.level ?? 0;
  if (style.list !== undefined) attributes.push(`lvl="${level}"`);
  const lineSpacing =
    style.lineHeight === undefined
      ? ""
      : `<a:lnSpc><a:spcPts val="${spacingCentipoints(positiveLength(style.lineHeight, "lineHeight"), "lineHeight")}"/></a:lnSpc>`;
  const spacingAfter =
    style.spacingAfter === undefined || style.spacingAfter === 0
      ? ""
      : `<a:spcAft><a:spcPts val="${spacingCentipoints(style.spacingAfter, "spacingAfter")}"/></a:spcAft>`;
  let list = "";
  if (style.list?.type === "bullet") {
    list = `<a:buChar char="${xmlEscape(style.list.character ?? "•")}"/>`;
  } else if (style.list?.type === "number") {
    list = `<a:buAutoNum type="${style.list.style ?? "arabicPeriod"}" startAt="${style.list.startAt ?? 1}"/>`;
  }
  const pPr =
    attributes.length +
      lineSpacing.length +
      spacingAfter.length +
      list.length ===
    0
      ? ""
      : `<a:pPr ${attributes.join(" ")}>${lineSpacing}${spacingAfter}${list}</a:pPr>`;
  const endProperties =
    runs.length === 0
      ? `<a:endParaRPr lang="en-US" sz="${centipoints(fontSizePoints(paragraph.endStyle.fontSize))}"/>`
      : "";
  return `<a:p>${pPr}${runs.map((run) => runXml(run, theme)).join("")}${endProperties}</a:p>`;
}

function normalizeParagraphs(
  children: readonly unknown[],
  textStyle: ResolvedTextStyle,
  theme: ThemeContext,
): Paragraph[] {
  const baseRunStyle: RunStyle = {
    fontFamily: textStyle.fontFamily,
    fontSize: textStyle.fontSize,
    fontWeight: textStyle.fontWeight,
    fontStyle: textStyle.fontStyle,
    color: textStyle.color,
    themeFont: textStyle.themeFont,
    letterSpacing: textStyle.letterSpacing,
    textTransform: textStyle.textTransform,
    textDecoration: textStyle.textDecoration,
    textShadow: textStyle.textShadow,
  };
  validateRunStyle(baseRunStyle);

  const paragraphs: Paragraph[] = [];
  let inline: Run[] = [];
  const baseParagraphStyle: ParagraphStyle = {
    ...(textStyle.lineHeight === undefined
      ? {}
      : { lineHeight: textStyle.lineHeight }),
    ...(textStyle.align === undefined ? {} : { align: textStyle.align }),
    ...(textStyle.spacingAfter === undefined
      ? {}
      : { spacingAfter: textStyle.spacingAfter }),
  };
  const flushInline = (): void => {
    if (inline.length > 0) {
      paragraphs.push({
        runs: inline,
        style: baseParagraphStyle,
        endStyle: baseRunStyle,
      });
      inline = [];
    }
  };

  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError("invalid-text-child", "文字内容树包含无效节点");
    }
    if (child.type === "#text" || child.type === "rp:span") {
      collectRuns([child], baseRunStyle, inline, theme);
      continue;
    }
    if (child.type !== "rp:paragraph") {
      throw new CapabilityError(
        "unsupported-text-child",
        `${child.type} 不能作为文本框的直接内容`,
      );
    }
    flushInline();
    const unknown = Object.keys(child.props).find((key) => key !== "style");
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `P 含未实现参数 ${unknown}`,
      );
    }
    const style = {
      ...baseParagraphStyle,
      ...paragraphStyle(child.props["style"]),
    };
    const paragraphRunStyle = mergeRunStyle(
      baseRunStyle,
      {
        fontFamily: style.fontFamily,
        fontSize: style.fontSize,
        fontWeight: style.fontWeight,
        fontStyle: style.fontStyle,
        color: style.color,
        letterSpacing: style.letterSpacing,
        textTransform: style.textTransform,
        textDecoration: style.textDecoration,
        textShadow: style.textShadow,
      },
      theme,
    );
    const runs: Run[] = [];
    collectRuns(child.children, paragraphRunStyle, runs, theme);
    paragraphs.push({ runs, style, endStyle: paragraphRunStyle });
  }
  flushInline();
  if (paragraphs.length === 0) {
    paragraphs.push({
      runs: [],
      style: baseParagraphStyle,
      endStyle: baseRunStyle,
    });
  }
  return paragraphs;
}

function fontFamilyKey(value: string): string {
  return value.trim().replace(/\s+/gu, " ").toLocaleLowerCase("en-US");
}

function measureFace(source: NormalizedFont): MeasureFace {
  if (!source.hasResource || source.data === undefined) {
    throw new CapabilityError(
      "font-measurement-unavailable",
      `内容定高文字缺少 ${source.family} ${source.weight} ${source.style} 的显式字体资源`,
    );
  }

  let parsed = measureFaces.get(source.identity);
  if (parsed === undefined) {
    try {
      const data = source.data;
      const bytes = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
      const opened = fontkit.create(bytes);
      if (!("unitsPerEm" in opened)) {
        throw new Error("字体集合不属于单字体文件");
      }
      parsed = opened;
      measureFaces.set(source.identity, parsed);
    } catch (error) {
      throw new CapabilityError(
        "font-format-unsupported",
        `无法解析 ${source.family} 的字体度量文件：${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }

  if (
    source.substitute !== true &&
    fontFamilyKey(parsed.familyName) !== fontFamilyKey(source.family)
  ) {
    throw new CapabilityError(
      "font-substitution-required",
      `度量文件声明为 ${parsed.familyName}，要保留输出字体名 ${source.family} 必须显式设置 Font.substitute`,
    );
  }
  if (parsed["OS/2"].usWeightClass !== source.weight) {
    throw new CapabilityError(
      "font-face-mismatch",
      `${source.family} 的 ${source.weight} ${source.style} 字体值与文件实际字重 ${parsed["OS/2"].usWeightClass} 不符`,
    );
  }
  const fileItalic =
    parsed["OS/2"].fsSelection.italic ||
    parsed["OS/2"].fsSelection.oblique ||
    parsed.italicAngle !== 0;
  if (fileItalic !== (source.style === "italic")) {
    throw new CapabilityError(
      "font-face-mismatch",
      `${source.family} 的 ${source.weight} ${source.style} 字体值与文件实际字形样式不符`,
    );
  }
  if (
    !Number.isFinite(parsed.unitsPerEm) ||
    parsed.unitsPerEm <= 0 ||
    !Number.isFinite(parsed.ascent) ||
    !Number.isFinite(parsed.descent) ||
    !Number.isFinite(parsed.lineGap)
  ) {
    throw new CapabilityError(
      "font-metrics-invalid",
      `${source.family} 的字体行距指标无效`,
    );
  }
  return { font: parsed, identity: source.identity };
}

function measureRunFace(style: RunStyle): MeasureFace {
  const candidates = normalizedFonts(style).filter(
    (candidate) => candidate.hasResource,
  );
  const unique = candidates.filter(
    (candidate, index) =>
      candidates.findIndex((other) => other.identity === candidate.identity) ===
      index,
  );
  if (unique.length !== 1) {
    throw new CapabilityError(
      unique.length === 0
        ? "font-measurement-unavailable"
        : "conflicting-resource",
      unique.length === 0
        ? "内容定高文字必须为每个实际使用的字重和样式提供字体文件"
        : "同一字体字重和样式对应多个度量文件，无法确定测量依据",
    );
  }
  return measureFace(unique[0]!);
}

function missingCharacters(face: MeasureFace, text: string): string[] {
  const missing = new Set<string>();
  for (const character of text) {
    if (
      /^[\p{Cc}\p{Cf}\u{FE00}-\u{FE0F}\u{E0100}-\u{E01EF}]$/u.test(character)
    ) {
      continue;
    }
    const codePoint = character.codePointAt(0);
    if (codePoint !== undefined && !face.font.hasGlyphForCodePoint(codePoint)) {
      missing.add(character);
    }
  }
  return [...missing];
}

function rangeWidth(
  pieces: readonly MeasurePiece[],
  text: string,
  start: number,
  end: number,
): number {
  let width = 0;
  for (const piece of pieces) {
    const from = Math.max(start, piece.start);
    const to = Math.min(end, piece.end);
    if (from >= to) continue;
    const source = text.slice(from, to);
    const transformed =
      piece.style.textTransform === "uppercase" ? source.toUpperCase() : source;
    let visible = "";
    for (const character of transformed) {
      if (
        !/^[\p{Cc}\p{Cf}\u{FE00}-\u{FE0F}\u{E0100}-\u{E01EF}]$/u.test(character)
      ) {
        visible += character;
      }
    }
    if (visible === "") continue;
    const advance = piece.face.font.layout(visible).advanceWidth;
    const pixels = fontSizePoints(piece.style.fontSize) * (4 / 3);
    const characterCount = Array.from(graphemes.segment(transformed)).length;
    width +=
      (advance / piece.face.font.unitsPerEm) * pixels +
      (piece.style.letterSpacing ?? 0) * characterCount;
  }
  return width;
}

function trimLineEnd(text: string, start: number, end: number): number {
  let trimmed = end;
  while (trimmed > start && /\s/u.test(text[trimmed - 1] ?? "")) {
    trimmed -= 1;
  }
  return trimmed;
}

function graphemeEnds(text: string, start: number, end: number): number[] {
  const result: number[] = [];
  for (const segment of graphemes.segment(text)) {
    const boundary = segment.index + segment.segment.length;
    if (boundary > start && boundary <= end) result.push(boundary);
  }
  return result;
}

function trimMandatoryBreakEnd(
  text: string,
  start: number,
  end: number,
): number {
  let trimmed = end;
  while (
    trimmed > start &&
    /[\n\r\v\f\u0085\u2028\u2029]/u.test(text[trimmed - 1] ?? "")
  ) {
    trimmed -= 1;
  }
  return trimmed;
}

function wrapParagraph(
  text: string,
  pieces: readonly MeasurePiece[],
  width: number,
): { readonly lines: number; readonly width: number } {
  if (text === "") return { lines: 1, width: 0 };
  const breaks: { readonly position: number; readonly required: boolean }[] =
    [];
  const breaker = new LineBreaker(text);
  let position = 0;
  for (
    let next = breaker.nextBreak();
    next !== null;
    next = breaker.nextBreak()
  ) {
    if (next.position > position) breaks.push(next);
    position = next.position;
  }
  if (position < text.length) {
    breaks.push({ position: text.length, required: false });
  }

  const lines: { readonly start: number; readonly end: number }[] = [];
  const epsilon = 0.01;
  let lineStart = 0;
  let lineEnd = 0;
  let unitStart = 0;
  let endedWithRequiredBreak = false;

  for (const lineBreak of breaks) {
    const unitEnd = lineBreak.position;
    const breakContentEnd = trimMandatoryBreakEnd(text, unitStart, unitEnd);
    // linebreak 把文本末尾的必需换行标为普通结束边界，控制符仍代表空行。
    const required = lineBreak.required || breakContentEnd < unitEnd;
    const candidateWidth = rangeWidth(
      pieces,
      text,
      lineStart,
      trimLineEnd(text, lineStart, unitEnd),
    );
    if (candidateWidth <= width + epsilon) {
      lineEnd = unitEnd;
      if (required) {
        lines.push({ start: lineStart, end: lineEnd });
        lineStart = unitEnd;
        lineEnd = unitEnd;
      }
      unitStart = unitEnd;
      endedWithRequiredBreak = required;
      continue;
    }

    if (lineEnd > lineStart) {
      lines.push({ start: lineStart, end: lineEnd });
      lineStart = unitStart;
      lineEnd = unitStart;
    }

    const unitWidth = rangeWidth(
      pieces,
      text,
      unitStart,
      trimLineEnd(text, unitStart, unitEnd),
    );
    if (unitWidth <= width + epsilon) {
      lineEnd = unitEnd;
    } else {
      let chunkStart = unitStart;
      let previous = chunkStart;
      const fallbackEnd = required ? breakContentEnd : unitEnd;
      for (const boundary of graphemeEnds(text, unitStart, fallbackEnd)) {
        const chunkWidth = rangeWidth(
          pieces,
          text,
          chunkStart,
          trimLineEnd(text, chunkStart, boundary),
        );
        if (chunkWidth > width + epsilon && previous > chunkStart) {
          lines.push({ start: chunkStart, end: previous });
          chunkStart = previous;
        }
        const singleWidth = rangeWidth(
          pieces,
          text,
          chunkStart,
          trimLineEnd(text, chunkStart, boundary),
        );
        if (singleWidth > width + epsilon && previous === chunkStart) {
          lines.push({ start: chunkStart, end: boundary });
          chunkStart = boundary;
        }
        previous = boundary;
      }
      lineStart = chunkStart;
      lineEnd = fallbackEnd;
    }

    if (required) {
      if (lineEnd > lineStart) {
        lines.push({ start: lineStart, end: lineEnd });
      }
      lineStart = unitEnd;
      lineEnd = unitEnd;
    }
    unitStart = unitEnd;
    endedWithRequiredBreak = required;
  }

  if (lineStart < text.length || lineEnd > lineStart || lines.length === 0) {
    lines.push({ start: lineStart, end: text.length });
  } else if (endedWithRequiredBreak) {
    lines.push({ start: text.length, end: text.length });
  }
  const lineWidths = lines.map(({ start, end }) =>
    rangeWidth(pieces, text, start, trimLineEnd(text, start, end)),
  );
  return {
    lines: lines.length,
    width: Math.max(0, ...lineWidths),
  };
}

function paragraphLineHeight(paragraph: Paragraph): number {
  if (paragraph.style.lineHeight !== undefined) {
    const value = positiveLength(paragraph.style.lineHeight, "lineHeight");
    return spacingCentipoints(value, "lineHeight") / 75;
  }
  const styles =
    paragraph.runs.length > 0
      ? paragraph.runs.map((run) => run.style)
      : [paragraph.endStyle];
  let maximum = 0;
  for (const style of styles) {
    const face = measureRunFace(style).font;
    // DrawingML 行距取字体 ascent - descent + lineGap，再按输出字号缩放。
    const units = face.ascent - face.descent + face.lineGap;
    const lineHeight =
      (units / face.unitsPerEm) * fontSizePoints(style.fontSize) * (4 / 3);
    if (!Number.isFinite(lineHeight) || lineHeight <= 0) {
      throw new CapabilityError(
        "font-metrics-invalid",
        "字体行距指标必须产生有限正值",
      );
    }
    maximum = Math.max(maximum, lineHeight);
  }
  if (maximum === 0) {
    throw new CapabilityError(
      "font-measurement-unavailable",
      "内容定高文字缺少可用于确定行距的字体",
    );
  }
  return spacingCentipoints(maximum, "lineHeight") / 75;
}

function paragraphBaseline(paragraph: Paragraph): number | undefined {
  // 混合 run 的首行基线需要逐段文字片段度量，当前结果不包含这类依据。
  if (paragraph.runs.length !== 1 || paragraph.runs[0]!.text === "") {
    return undefined;
  }
  const style = paragraph.runs[0]!.style;
  const face = measureRunFace(style).font;
  const fontSize = fontSizePoints(style.fontSize) * (4 / 3);
  const naturalLineHeight =
    ((face.ascent - face.descent + face.lineGap) / face.unitsPerEm) * fontSize;
  const lineHeight = paragraphLineHeight(paragraph);
  const leading = (lineHeight - naturalLineHeight) / 2;
  return (face.ascent / face.unitsPerEm) * fontSize + leading;
}

function paragraphPieces(paragraph: Paragraph): {
  readonly text: string;
  readonly pieces: readonly MeasurePiece[];
} {
  const sourceRuns =
    paragraph.runs.length > 0
      ? paragraph.runs
      : [{ text: "", style: paragraph.endStyle }];
  const pieces: MeasurePiece[] = [];
  let start = 0;
  for (const run of sourceRuns) {
    const face = measureRunFace(run.style);
    const shown =
      run.style.textTransform === "uppercase"
        ? run.text.toUpperCase()
        : run.text;
    const missing = missingCharacters(face, shown);
    if (missing.length > 0) {
      const shown = missing.slice(0, 10).join(" ");
      const extra = missing.length - 10;
      const family = fontFamilyValue(run.style.fontFamily).family;
      throw new CapabilityError(
        "font-missing-glyph",
        `${family} 缺少字符 ${shown}${extra > 0 ? `，另有 ${extra} 个字符` : ""} 的字形`,
      );
    }
    pieces.push({
      text: run.text,
      style: run.style,
      start,
      end: start + run.text.length,
      face,
    });
    start += run.text.length;
  }
  const text = sourceRuns.map((run) => run.text).join("");
  return { text, pieces };
}

function paragraphMinContentWidth(paragraph: Paragraph): number {
  const { text, pieces } = paragraphPieces(paragraph);
  if (text === "") return 0;
  const breaker = new LineBreaker(text);
  let start = 0;
  let widest = 0;
  for (
    let next = breaker.nextBreak();
    next !== null;
    next = breaker.nextBreak()
  ) {
    const contentEnd = trimMandatoryBreakEnd(text, start, next.position);
    widest = Math.max(
      widest,
      rangeWidth(pieces, text, start, trimLineEnd(text, start, contentEnd)),
    );
    start = next.position;
  }
  return widest;
}

function paragraphMeasurement(
  paragraph: Paragraph,
  width: number,
): MeasuredParagraph {
  const { text, pieces } = paragraphPieces(paragraph);
  const wrapped = wrapParagraph(text, pieces, width);
  const spacingAfter =
    paragraph.style.spacingAfter === undefined
      ? 0
      : spacingCentipoints(paragraph.style.spacingAfter, "spacingAfter") / 75;
  return {
    lineCount: wrapped.lines,
    width: wrapped.width,
    lineHeight: paragraphLineHeight(paragraph),
    spacingAfter,
  };
}

function measureParagraphs(
  paragraphs: readonly Paragraph[],
  width: number,
): {
  readonly width: number;
  readonly height: number;
  readonly paragraphs: readonly MeasuredParagraph[];
} {
  const measured = paragraphs.map((paragraph) =>
    paragraphMeasurement(paragraph, width),
  );
  let height = 0;
  measured.forEach((paragraph, index) => {
    height += paragraph.lineCount * paragraph.lineHeight;
    if (index < measured.length - 1) height += paragraph.spacingAfter;
  });
  return {
    width: Math.max(0, ...measured.map((paragraph) => paragraph.width)),
    height,
    paragraphs: measured,
  };
}

/** 以共享文字服务的实际字体与断行规则返回布局尺寸。 */
export function measureTextLayout(
  children: readonly unknown[],
  styleValue: unknown,
  theme: ThemeContext,
  availableWidth?: number,
): TextLayoutMeasurement {
  const style = readTextStyle(styleValue, "Text.style", theme);
  const paragraphs = normalizeParagraphs(children, style, theme);
  const measured = measureParagraphs(
    paragraphs,
    availableWidth ?? Number.POSITIVE_INFINITY,
  );
  const minContentWidth = Math.max(
    0,
    ...paragraphs.map(paragraphMinContentWidth),
  );
  const baseline =
    style.verticalAlign === undefined || style.verticalAlign === "top"
      ? paragraphs[0] === undefined
        ? undefined
        : paragraphBaseline(paragraphs[0])
      : undefined;
  return {
    width: measured.width,
    height: measured.height,
    minContentWidth,
    minContentHeight: measured.height,
    ...(baseline === undefined ? {} : { baseline }),
  };
}

function bodyAnchor(verticalAlign: TextStyle["verticalAlign"]): string {
  return { top: "t", middle: "ctr", bottom: "b" }[verticalAlign ?? "top"];
}

export function textInsets(value: unknown): {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
} {
  const style = record(value) ? value : {};
  const raw = style["padding"];
  let edges = { top: 0, right: 0, bottom: 0, left: 0 };
  if (typeof raw === "number") {
    edges = { top: raw, right: raw, bottom: raw, left: raw };
  } else if (Array.isArray(raw) && (raw.length === 2 || raw.length === 4)) {
    const values = raw;
    edges =
      values.length === 2
        ? {
            top: values[0] as number,
            right: values[1] as number,
            bottom: values[0] as number,
            left: values[1] as number,
          }
        : {
            top: values[0] as number,
            right: values[1] as number,
            bottom: values[2] as number,
            left: values[3] as number,
          };
  } else if (raw !== undefined) {
    throw new CapabilityError(
      "invalid-text-insets",
      "padding 必须是数字、二元组或四元组",
    );
  }
  const result: { top: number; right: number; bottom: number; left: number } = {
    top: (style["paddingTop"] ?? edges.top) as number,
    right: (style["paddingRight"] ?? edges.right) as number,
    bottom: (style["paddingBottom"] ?? edges.bottom) as number,
    left: (style["paddingLeft"] ?? edges.left) as number,
  };
  for (const [name, amount] of Object.entries(result)) {
    if (typeof amount !== "number" || !Number.isFinite(amount) || amount < 0) {
      throw new CapabilityError(
        "invalid-text-insets",
        `${name} inset 必须是有限非负像素数值`,
      );
    }
    pxToEmu(amount);
  }
  return result;
}

/** 为文本框与形状文字共用段落、run、列表和字体资源处理。 */
export function writeTextBody(
  children: readonly unknown[],
  styleValue: unknown,
  theme: ThemeContext,
  verticalAlign: TextStyle["verticalAlign"] = "top",
  layout?: TextLayoutRequest,
  destination: TextBodyDestination = "text-box",
): TextServiceResult {
  const style = readTextStyle(styleValue, "Text.style", theme);
  const paragraphs = normalizeParagraphs(children, style, theme);
  const runs = paragraphs.flatMap((paragraph) =>
    paragraph.runs.length > 0
      ? paragraph.runs
      : [{ text: "", style: paragraph.endStyle }],
  );
  const { resourceClaims, fontUses, warnings } = textResourceClaims(runs);
  let measured: ReturnType<typeof measureParagraphs> | undefined;
  if (layout !== undefined) {
    try {
      measured = measureParagraphs(
        paragraphs,
        layout.width === "max-content" ? Number.MAX_VALUE : layout.width,
      );
    } catch (error) {
      if (layout.contentSized === true) throw error;
      if (error instanceof CapabilityError) {
        if (!warnings.some((warning) => warning.code === error.code)) {
          warnings.push({
            code: error.code,
            message: `未能依据已提供的字体完成文字度量：${error.message}`,
          });
        }
      } else {
        warnings.push({
          code: "font-measurement-unavailable",
          message: `未能依据已提供的字体完成文字度量：${error instanceof Error ? error.message : String(error)}`,
        });
      }
    }
  }
  if (layout?.contentSized === true && measured === undefined) {
    throw new CapabilityError(
      "font-measurement-unavailable",
      "内容定高文字缺少完整的字体度量依据",
    );
  }
  if (
    measured !== undefined &&
    layout !== undefined &&
    ((typeof layout.width === "number" &&
      measured.width > layout.width + 0.01) ||
      (layout.height !== undefined && measured.height > layout.height + 0.01))
  ) {
    warnings.push({
      code: "overset",
      message:
        "按已提供的字体度量，文字超出文本框的宽或高；构建保留原始内容和尺寸",
    });
  }
  const outputParagraphs =
    layout?.contentSized === true && measured !== undefined
      ? paragraphs.map((paragraph, index) => ({
          ...paragraph,
          style: {
            ...paragraph.style,
            lineHeight: measured.paragraphs[index]!.lineHeight,
          },
        }))
      : paragraphs;
  const text = outputParagraphs
    .map(({ runs }) => runs.map(({ text: part }) => part).join(""))
    .join("\n");
  const xml =
    destination === "table-cell"
      ? `<a:txBody><a:bodyPr/><a:lstStyle/>${outputParagraphs.map((paragraph) => paragraphXml(paragraph, theme)).join("")}</a:txBody>`
      : `<p:txBody><a:bodyPr wrap="square" lIns="${pxToEmu(layout?.insets?.left ?? 0)}" tIns="${pxToEmu(layout?.insets?.top ?? 0)}" rIns="${pxToEmu(layout?.insets?.right ?? 0)}" bIns="${pxToEmu(layout?.insets?.bottom ?? 0)}" anchor="${bodyAnchor(style.verticalAlign ?? verticalAlign)}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${outputParagraphs.map((paragraph) => paragraphXml(paragraph, theme)).join("")}</p:txBody>`;
  return {
    xml,
    text,
    paragraphCount: outputParagraphs.length,
    resourceClaims,
    fontUses,
    ...(measured === undefined
      ? {}
      : { measurement: { width: measured.width, height: measured.height } }),
    warnings,
  };
}

/** 为原生图表文字生成共享字体样式，并按实际类别、系列和标签文本登记字体 claims。 */
export function writeChartTextProperties(
  styleValue: unknown,
  texts: readonly string[],
  label: string,
  theme: ThemeContext,
): Pick<
  TextServiceResult,
  "text" | "resourceClaims" | "fontUses" | "warnings"
> & {
  readonly xml: string;
} {
  const style = readTextStyle(styleValue, label, theme);
  const runStyle: RunStyle = {
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    color: style.color,
    themeFont: style.themeFont,
  };
  validateRunStyle(runStyle);
  const runs = texts.map((text) => ({ text, style: runStyle }));
  for (const text of texts) escapedXml(text);
  const { resourceClaims, fontUses, warnings } = textResourceClaims(runs);
  const defaultProperties = runPropertiesXml(runStyle, "a:defRPr", theme);
  const outputText = texts.join("\n");
  const lineSpacing =
    style.lineHeight === undefined
      ? ""
      : `<a:lnSpc><a:spcPts val="${spacingCentipoints(positiveLength(style.lineHeight, "lineHeight"), "lineHeight")}"/></a:lnSpc>`;
  const alignment = alignmentValue(style.align);
  const paragraphAlignment =
    alignment === undefined ? "" : ` algn="${alignment}"`;
  const spacingAfter =
    style.spacingAfter === undefined || style.spacingAfter === 0
      ? ""
      : `<a:spcAft><a:spcPts val="${spacingCentipoints(style.spacingAfter, "spacingAfter")}"/></a:spcAft>`;
  return {
    xml: `<c:txPr><a:bodyPr anchor="${bodyAnchor(style.verticalAlign)}"/><a:lstStyle/><a:p><a:pPr${paragraphAlignment}>${lineSpacing}${spacingAfter}${defaultProperties}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`,
    text: outputText,
    resourceClaims,
    fontUses,
    warnings,
  };
}

/** 把固定像素长度转换成 DrawingML 可表达的 EMU。 */
export function textExtent(value: number, label: string): number {
  if (!Number.isFinite(value) || value <= 0) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} 必须是有限正像素数值`,
    );
  }
  const extent = pxToEmu(value);
  if (extent < 1 || extent > 2_147_483_647) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} 换算后超出形状尺寸范围`,
    );
  }
  return extent;
}

/** 读取固定区域文本的原始纯文本，供形状和只读报告共用。 */
export function plainText(children: readonly unknown[]): string {
  const result: string[] = [];
  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError("invalid-text-child", "文字内容树包含无效节点");
    }
    if (child.type === "#text") result.push(child.text);
    else if (child.type === "rp:span" || child.type === "rp:paragraph") {
      result.push(plainText(child.children));
    } else {
      throw new CapabilityError(
        "unsupported-text-child",
        `${child.type} 不是文本内容`,
      );
    }
  }
  return result.join("");
}
