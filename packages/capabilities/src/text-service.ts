import { createHash } from "node:crypto";

import {
  assertValidLang,
  CapabilityError,
  pxToEmu,
  resolvedColorXml,
  textStyleKeys,
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
} from "@react-pptx/core/capability";
import * as fontkit from "fontkit";
// @ts-expect-error -- linebreak does not provide types; declare the breakpoint interface used below.
import LineBreakerImplementation from "linebreak";

import type { FontEmbeddingMode, FontEmbeddingUse } from "./font-embedding.ts";
import { matchesFontFamily } from "./font-family.ts";
import { textShadowXml } from "./visual.ts";
export type {
  FontReference,
  FontSize,
  TextAlignment,
  TextStyle,
} from "@react-pptx/core/capability";

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

export type ParagraphStyle = Omit<TextStyle, "verticalAlign" | "textWrap"> & {
  readonly align?: TextAlignment;
  readonly spacingAfter?: number;
  readonly list?: TextList;
};

export type TextResourceClaim = {
  readonly kind: "font-face";
  readonly key: string;
  readonly identity: string;
  /** Text characters actually emitted for this font style, reused for subsequent font processing. */
  readonly characters: string;
};

/** Retains control characters and default ignorable code points in text stream without requiring standalone cmap glyphs. */
export function requiresFontGlyph(character: string): boolean {
  return !/^[\p{Cc}\p{Default_Ignorable_Code_Point}]$/u.test(character);
}

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
  readonly baseline?: number | "unavailable";
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
  readonly lang?: string | undefined;
};
export type Run = {
  readonly text: string;
  readonly style: RunStyle;
  readonly field?: { readonly id: string; readonly type: "slidenum" };
};
export type TextFieldContext = {
  readonly slideNumber: number;
  readonly createId?: () => string;
};

/**
 * Author metadata accepted on P, Span, and SlideNumber declarations.
 *
 * Scoped id validation belongs to the formal build around these nodes; the shared text parser must
 * skip the field while resolving styles instead of reporting it as an unimplemented parameter.
 */
const textAuthorMetadataKeys: ReadonlySet<string> = new Set(["id"]);

export type Paragraph = {
  readonly runs: readonly Run[];
  readonly style: ParagraphStyle;
  readonly endStyle: RunStyle;
  readonly wrap: "normal" | "nowrap";
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
      "Text content or font name contains characters not allowed in XML 1.0",
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
    throw new CapabilityError("invalid-style", `${label} must be an object`);
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
      `${label} contains unimplemented style ${unknown}`,
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
        `${label} must be a positive pixel value or px/pt length`,
      );
    }
    const amount = Number(match[1]);
    pixels = match[2] === "pt" ? (amount * 4) / 3 : amount;
  } else {
    throw new CapabilityError(
      "invalid-text-length",
      `${label} must be a positive pixel value or px/pt length`,
    );
  }
  if (!Number.isFinite(pixels) || pixels <= 0) {
    throw new CapabilityError(
      "invalid-text-length",
      `${label} must be greater than 0`,
    );
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
      `${label} converted value must be within DrawingML range of 0 to 1584pt`,
    );
  }
  return value;
}

function letterSpacingCentipoints(value: number): number {
  const points = Math.round(value * 75);
  if (!Number.isFinite(value) || points < -400_000 || points > 400_000) {
    throw new CapabilityError(
      "invalid-letter-spacing",
      "letterSpacing converted value must be within DrawingML range of -4000 to 4000pt",
    );
  }
  return points;
}

function fontSizePoints(value: FontSize | undefined): number {
  let points: number;
  if (value === undefined) {
    points = positiveLength(defaultTextStyle.fontSize, "fontSize") * 0.75;
  } else if (typeof value === "string" && value.endsWith("pt")) {
    points = Number(value.slice(0, -2));
    if (!Number.isFinite(points) || points <= 0) {
      throw new CapabilityError(
        "invalid-font-size",
        "fontSize must be greater than 0",
      );
    }
  } else {
    points = positiveLength(value, "fontSize") * 0.75;
  }
  const serialized = centipoints(points);
  if (serialized < 100 || serialized > 400_000) {
    throw new CapabilityError(
      "invalid-font-size",
      "fontSize converted value must be within DrawingML range of 1 to 4000pt",
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
      "fontWeight must be normal, bold, or a hundreds integer between 100 and 900",
    );
  }
  return value;
}

function fontFamilyValue(value: FontReference | undefined): Font {
  if (value === undefined) return { family: "Aptos" };
  if (typeof value === "string") {
    if (value.trim() === "") {
      throw new CapabilityError(
        "invalid-font-family",
        "fontFamily cannot be empty",
      );
    }
    return { family: value };
  }
  if (!record(value) || typeof value["family"] !== "string") {
    throw new CapabilityError(
      "invalid-font-family",
      "fontFamily must be a font value",
    );
  }
  if (value["family"].trim() === "") {
    throw new CapabilityError(
      "invalid-font-family",
      "Font.family cannot be empty",
    );
  }
  onlyKnownKeys(value, ["family", "faces", "substitute", "embeddable"], "Font");
  if (value["faces"] !== undefined && !Array.isArray(value["faces"])) {
    throw new CapabilityError(
      "invalid-font-family",
      "Font.faces must be an array",
    );
  }
  if (
    value["substitute"] !== undefined &&
    typeof value["substitute"] !== "boolean"
  ) {
    throw new CapabilityError(
      "invalid-font-family",
      "Font.substitute must be a boolean",
    );
  }
  if (
    value["embeddable"] !== undefined &&
    typeof value["embeddable"] !== "boolean"
  ) {
    throw new CapabilityError(
      "invalid-font-family",
      "Font.embeddable must be a boolean",
    );
  }
  return value;
}

function normalizedFonts(style: RunStyle): NormalizedFont[] {
  const family = fontFamilyValue(style.fontFamily);
  const weight = weightValue(style.fontWeight);
  const faceStyle = style.fontStyle ?? "normal";
  if (faceStyle !== "normal" && faceStyle !== "italic") {
    throw new CapabilityError(
      "invalid-font-style",
      "fontStyle must be normal or italic",
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
        "Font.faces contains invalid weight, style, or resource",
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
      "textTransform must be none or uppercase",
    );
  }
  if (
    style.textDecoration !== undefined &&
    !["none", "underline"].includes(style.textDecoration)
  ) {
    throw new CapabilityError(
      "invalid-text-decoration",
      "textDecoration must be none or underline",
    );
  }
  if (style.lang !== undefined) {
    assertValidLang(style.lang);
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
      "lang",
    ],
    "Span.style",
  );
  return validateRunStyle(props);
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
    declared["color"] = theme.resolveColor(child.color);
  }
  if (child.lang !== undefined) {
    declared["lang"] = child.lang;
  }
  // Undeclared fields in P and Span are undefined and cannot overwrite inherited parent styles.
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
    // Numerical font weight cannot substitute for family bold style link declared by the font file.
    return measureFace(font).font["OS/2"].fsSelection.bold;
  } catch {
    return fallback;
  }
}

export type ResolvedTextStyle = TextStyle & {
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
      "align must be left, center, right, or justify",
    );
  }
  if (
    style.spacingAfter !== undefined &&
    (!Number.isFinite(style.spacingAfter) || style.spacingAfter < 0)
  ) {
    throw new CapabilityError(
      "invalid-paragraph-spacing",
      "spacingAfter must be a finite non-negative pixel number",
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
  onlyKnownKeys(props, [...textStyleKeys], label);
  if (props["lang"] !== undefined) {
    assertValidLang(props["lang"]);
  }
  const style = {
    ...defaultTextStyle,
    ...theme.textStyle,
    ...props,
  } as TextStyle;
  if (style.textShadow === "none")
    delete (style as { textShadow?: Shadow | "none" }).textShadow;
  if (
    style.textWrap !== undefined &&
    style.textWrap !== "normal" &&
    style.textWrap !== "nowrap"
  ) {
    throw new CapabilityError(
      "invalid-text-wrap",
      "textWrap must be normal or nowrap",
    );
  }
  if (
    style.verticalAlign !== undefined &&
    !["top", "middle", "bottom"].includes(style.verticalAlign)
  ) {
    throw new CapabilityError(
      "invalid-vertical-align",
      "verticalAlign must be top, middle, or bottom",
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
    lang: style.lang,
  };
  const font = theme.resolveFont(
    runStyle.fontFamily ?? defaultTextStyle.fontFamily!,
  );
  if (runStyle.color !== undefined && !isResolvedColor(runStyle.color)) {
    theme.resolveColor(runStyle.color);
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
      "lang",
    ],
    "P.style",
  );
  if (props["lang"] !== undefined) {
    assertValidLang(props["lang"]);
  }
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
      "list.type must be bullet or number",
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
      "list.level must be an integer between 0 and 8",
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
        "Bullet must be a single Unicode character",
      );
    }
    try {
      xmlEscape(character);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "Bullet contains characters not allowed in XML 1.0",
      );
    }
  } else {
    if (value["style"] !== undefined && value["style"] !== "arabicPeriod") {
      throw new CapabilityError(
        "invalid-list",
        "Numbering style currently only supports arabicPeriod",
      );
    }
    const startAt = value["startAt"] ?? 1;
    if (!Number.isInteger(startAt) || startAt < 1 || startAt > 32767) {
      throw new CapabilityError(
        "invalid-list",
        "Numbering start value must be an integer between 1 and 32767",
      );
    }
  }
}

function collectRuns(
  children: readonly unknown[],
  style: RunStyle,
  runs: Run[],
  theme: ThemeContext,
  fields: TextFieldContext,
): void {
  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError(
        "invalid-text-child",
        "Text content tree contains invalid nodes",
      );
    }
    if (child.type === "#text") {
      if (child.text !== "") runs.push({ text: child.text, style });
      continue;
    }
    if (child.type !== "rp:span" && child.type !== "rp:slide-number") {
      throw new CapabilityError(
        "unsupported-text-child",
        `${child.type} cannot be used as text paragraph or run`,
      );
    }
    const unknown = Object.keys(child.props).find(
      (key) => key !== "style" && !textAuthorMetadataKeys.has(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Span contains unimplemented parameter ${unknown}`,
      );
    }
    const spanStyle = readRunStyle(child.props["style"]);
    const nextStyle = mergeRunStyle(style, spanStyle, theme);
    if (child.type === "rp:slide-number") {
      if (child.children.length > 0) {
        throw new CapabilityError(
          "invalid-slide-number",
          "SlideNumber cannot contain children",
        );
      }
      runs.push({
        text: String(fields.slideNumber),
        style: nextStyle,
        field: {
          id:
            fields.createId?.() ??
            // Table normalization used only for layout will not emit XML for this field.
            "00000000-0000-4000-8000-000000000001",
          type: "slidenum",
        },
      });
      continue;
    }
    collectRuns(child.children, nextStyle, runs, theme, fields);
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
  const latinTypeface =
    style.themeFont === "major"
      ? "+mj-lt"
      : style.themeFont === "minor"
        ? "+mn-lt"
        : font.family;
  const eaTypeface =
    style.themeFont === "major"
      ? "+mj-ea"
      : style.themeFont === "minor"
        ? "+mn-ea"
        : font.family;
  const csTypeface =
    style.themeFont === "major"
      ? "+mj-cs"
      : style.themeFont === "minor"
        ? "+mn-cs"
        : font.family;
  const fontSlots = `<a:latin typeface="${escapedXml(latinTypeface)}"/><a:ea typeface="${escapedXml(eaTypeface)}"/><a:cs typeface="${escapedXml(csTypeface)}"/>`;
  const attributes = [
    ...(style.lang !== undefined ? [`lang="${escapedXml(style.lang)}"`] : []),
    `sz="${centipoints(points)}"`,
    `b="${bold ? 1 : 0}"`,
    `i="${font.style === "italic" ? 1 : 0}"`,
    ...(style.textDecoration === "underline" ? ['u="sng"'] : []),
    ...(style.textTransform === "uppercase" ? ['cap="all"'] : []),
    ...(style.letterSpacing === undefined || style.letterSpacing === 0
      ? []
      : [`spc="${letterSpacingCentipoints(style.letterSpacing)}"`]),
  ].join(" ");
  return `<${tag} ${attributes}><a:solidFill>${resolvedColorXml(color)}</a:solidFill>${textShadowXml(theme, style.textShadow)}${fontSlots}</${tag}>`;
}

function displayedText(run: Run): string {
  return run.style.textTransform === "uppercase"
    ? run.text.toUpperCase()
    : run.text;
}

function runXml(run: Run, theme: ThemeContext): string {
  const text = escapedXml(run.text);
  if (run.field !== undefined) {
    return `<a:fld id="${escapedXml(run.field.id)}" type="${run.field.type}">${runPropertiesXml(run.style, "a:rPr", theme)}<a:t>${text}</a:t></a:fld>`;
  }
  return `<a:r>${runPropertiesXml(run.style, "a:rPr", theme)}<a:t xml:space="preserve">${text}</a:t></a:r>`;
}

function textResourceClaims(
  runs: readonly Run[],
  embedFonts?: FontEmbeddingMode,
): {
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
      // Byte resources can be merged, but each actually used embedding policy must be independently validated.
      const policyKey = `${font.identity}\u0000${font.substitute ? 1 : 0}\u0000${font.embeddable ? 1 : 0}`;
      entry.fontPolicies.set(policyKey, font);
      entry.text.push(
        displayedText(run),
        ...(run.field?.type === "slidenum" ? ["0123456789"] : []),
      );
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
        claimKey: key,
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
          message: `Explicit font resource is not provided for ${font.family} ${font.weight} ${font.style}; fixed-area text will not verify actual line wrapping or fitting`,
        });
      } else {
        resourceClaims.push({
          kind: "font-face",
          key,
          identity: font.identity,
          characters: entry.text.join(""),
        });
        const face = fontkitFace(font);
        if (face !== undefined) {
          const missing = missingCharacters(face, entry.text.join(""));
          if (missing.length > 0) {
            const shown = missing.slice(0, 10).join(" ");
            const extra = missing.length - 10;
            const message = `${font.family} is missing glyphs for characters ${shown}${extra > 0 ? `, and ${extra} more characters` : ""}`;
            if (embedFonts === "full" || embedFonts === "subset") {
              throw new CapabilityError("font-missing-glyph", message);
            } else {
              warnings.push({
                code: "font-missing-glyph",
                message,
              });
            }
          }
        }
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
      ? `<a:endParaRPr${paragraph.endStyle.lang !== undefined ? ` lang="${escapedXml(paragraph.endStyle.lang)}"` : ""} sz="${centipoints(fontSizePoints(paragraph.endStyle.fontSize))}"/>`
      : "";
  return `<a:p>${pPr}${runs.map((run) => runXml(run, theme)).join("")}${endProperties}</a:p>`;
}

function normalizeParagraphs(
  children: readonly unknown[],
  textStyle: ResolvedTextStyle,
  theme: ThemeContext,
  fields: TextFieldContext,
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
    lang: textStyle.lang,
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
    ...(textStyle.lang === undefined ? {} : { lang: textStyle.lang }),
  };
  const flushInline = (): void => {
    if (inline.length > 0) {
      paragraphs.push({
        runs: inline,
        style: baseParagraphStyle,
        endStyle: baseRunStyle,
        wrap: textStyle.textWrap ?? "normal",
      });
      inline = [];
    }
  };

  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError(
        "invalid-text-child",
        "Text content tree contains invalid nodes",
      );
    }
    if (
      child.type === "#text" ||
      child.type === "rp:span" ||
      child.type === "rp:slide-number"
    ) {
      collectRuns([child], baseRunStyle, inline, theme, fields);
      continue;
    }
    if (child.type !== "rp:paragraph") {
      throw new CapabilityError(
        "unsupported-text-child",
        `${child.type} cannot be used directly in text box`,
      );
    }
    flushInline();
    const unknown = Object.keys(child.props).find(
      (key) => key !== "style" && !textAuthorMetadataKeys.has(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `P contains unimplemented parameter ${unknown}`,
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
        lang: style.lang,
      },
      theme,
    );
    const runs: Run[] = [];
    collectRuns(child.children, paragraphRunStyle, runs, theme, fields);
    paragraphs.push({
      runs,
      style,
      endStyle: paragraphRunStyle,
      wrap: textStyle.textWrap ?? "normal",
    });
  }
  flushInline();
  if (paragraphs.length === 0) {
    paragraphs.push({
      runs: [],
      style: baseParagraphStyle,
      endStyle: baseRunStyle,
      wrap: textStyle.textWrap ?? "normal",
    });
  }
  return paragraphs;
}

function measureFace(source: NormalizedFont): MeasureFace {
  if (!source.hasResource || source.data === undefined) {
    throw new CapabilityError(
      "font-measurement-unavailable",
      `Content-sized text is missing explicit font resource for ${source.family} ${source.weight} ${source.style}`,
    );
  }

  let parsed = measureFaces.get(source.identity);
  if (parsed === undefined) {
    try {
      const data = source.data;
      const bytes = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
      const opened = fontkit.create(bytes);
      if (!("unitsPerEm" in opened)) {
        throw new Error("Font collection is not a single font file");
      }
      parsed = opened;
      measureFaces.set(source.identity, parsed);
    } catch (error) {
      throw new CapabilityError(
        "font-format-unsupported",
        `Cannot parse font metrics file for ${source.family}: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }

  if (!source.substitute && !matchesFontFamily(parsed, source.family)) {
    throw new CapabilityError(
      "font-substitution-required",
      `Metrics file declares ${parsed.familyName}; retaining output font name ${source.family} requires setting Font.substitute explicitly`,
    );
  }
  if (parsed["OS/2"].usWeightClass !== source.weight) {
    throw new CapabilityError(
      "font-face-mismatch",
      `Font value for ${source.family} ${source.weight} ${source.style} does not match file actual weight ${parsed["OS/2"].usWeightClass}`,
    );
  }
  const fileItalic =
    parsed["OS/2"].fsSelection.italic ||
    parsed["OS/2"].fsSelection.oblique ||
    parsed.italicAngle !== 0;
  if (fileItalic !== (source.style === "italic")) {
    throw new CapabilityError(
      "font-face-mismatch",
      `Font value for ${source.family} ${source.weight} ${source.style} does not match file actual glyph style`,
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
      `Font line metrics for ${source.family} are invalid`,
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
        ? "Content-sized text must provide a font file for each actually used weight and style"
        : "Multiple metrics files found for the same font weight and style; cannot determine measurement basis",
    );
  }
  return measureFace(unique[0]!);
}

function fontkitFace(source: NormalizedFont): fontkit.Font | undefined {
  if (!source.hasResource || source.data === undefined) return undefined;
  let parsed = measureFaces.get(source.identity);
  if (parsed === undefined) {
    try {
      const data = source.data;
      const bytes = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
      const opened = fontkit.create(bytes);
      if (
        !("unitsPerEm" in opened) ||
        typeof opened.hasGlyphForCodePoint !== "function"
      ) {
        return undefined;
      }
      parsed = opened;
      measureFaces.set(source.identity, parsed);
    } catch {
      return undefined;
    }
  }
  return parsed;
}

function missingCharacters(
  face: MeasureFace | fontkit.Font,
  text: string,
): string[] {
  const font = "font" in face ? face.font : face;
  const missing = new Set<string>();
  for (const character of text) {
    if (!requiresFontGlyph(character)) continue;
    const codePoint = character.codePointAt(0);
    if (codePoint !== undefined && !font.hasGlyphForCodePoint(codePoint)) {
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
  // Native text layout quantizes each shaped glyph advance. PowerPoint rounds to 1/8 pt;
  // WPS rounds the serialized centipoint size up to the next 0.1 pt, then rounds shaped
  // advances to 1/20 pt (one twip). Keep the widest whole-range candidate so intrinsic
  // wrapping stays conservative without changing the author's size. The evidence covers
  // the current desktop sample, not every engine or font.
  const nativeEighthPointGrid = 8;
  const twipsPerPoint = 20;
  let natural = 0;
  let eighthPoint = 0;
  let twips = 0;
  let spacing = 0;
  for (const piece of pieces) {
    const from = Math.max(start, piece.start);
    const to = Math.min(end, piece.end);
    if (from >= to) continue;
    const source = text.slice(from, to);
    const transformed =
      piece.style.textTransform === "uppercase" ? source.toUpperCase() : source;
    let visible = "";
    for (const character of transformed) {
      if (requiresFontGlyph(character)) visible += character;
    }
    if (visible === "") continue;
    const face = piece.face.font;
    const points = fontSizePoints(piece.style.fontSize);
    const tenths = Math.ceil(centipoints(points) / 10) / 10;
    for (const position of face.layout(visible).positions) {
      const advance = (position.xAdvance / face.unitsPerEm) * points;
      natural += advance;
      eighthPoint +=
        Math.round(advance * nativeEighthPointGrid) / nativeEighthPointGrid;
      const tenthAdvance = (position.xAdvance / face.unitsPerEm) * tenths;
      twips += Math.round(tenthAdvance * twipsPerPoint) / twipsPerPoint;
    }
    const characterCount = Array.from(graphemes.segment(transformed)).length;
    spacing += (piece.style.letterSpacing ?? 0) * characterCount;
  }
  return Math.max(natural, eighthPoint, twips) * (4 / 3) + spacing;
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
    // linebreak marks mandatory line breaks at the end of text as normal break boundaries; control chars still represent empty lines.
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
    // DrawingML line spacing takes font ascent - descent + lineGap, then scales by output font size.
    const units = face.ascent - face.descent + face.lineGap;
    const lineHeight =
      (units / face.unitsPerEm) * fontSizePoints(style.fontSize) * (4 / 3);
    if (!Number.isFinite(lineHeight) || lineHeight <= 0) {
      throw new CapabilityError(
        "font-metrics-invalid",
        "Font line metrics must produce a finite positive value",
      );
    }
    maximum = Math.max(maximum, lineHeight);
  }
  if (maximum === 0) {
    throw new CapabilityError(
      "font-measurement-unavailable",
      "Content-sized text is missing font to determine line spacing",
    );
  }
  return spacingCentipoints(maximum, "lineHeight") / 75;
}

function paragraphBaseline(paragraph: Paragraph): number | undefined {
  // First line baseline of mixed runs requires measuring individual text fragments; current results do not include such basis.
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
    const shown = displayedText(run);
    const letterSpacing =
      run.style.letterSpacing === undefined
        ? undefined
        : letterSpacingCentipoints(run.style.letterSpacing) / 75;
    const measureStyle =
      letterSpacing === undefined ? run.style : { ...run.style, letterSpacing };
    pieces.push({
      text: shown,
      style: measureStyle,
      start,
      end: start + shown.length,
      face,
    });
    start += shown.length;
  }
  const text = pieces.map((piece) => piece.text).join("");
  return { text, pieces };
}

function paragraphMinContentWidth(paragraph: Paragraph): number {
  const { text, pieces } = paragraphPieces(paragraph);
  if (text === "") return 0;
  if (paragraph.wrap === "nowrap") {
    return wrapParagraph(text, pieces, Number.MAX_VALUE).width;
  }
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
  const wrapped = wrapParagraph(
    text,
    pieces,
    paragraph.wrap === "nowrap" ? Number.MAX_VALUE : width,
  );
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

/** Returns layout dimensions based on actual fonts and line breaking rules from normalized paragraphs. */
export function measureTextLayoutFromParagraphs(
  paragraphs: readonly Paragraph[],
  style: ResolvedTextStyle | TextStyle,
  availableWidth?: number,
): TextLayoutMeasurement {
  const measured = measureParagraphs(
    paragraphs,
    availableWidth ?? Number.POSITIVE_INFINITY,
  );
  const minContentWidth = Math.max(
    0,
    ...paragraphs.map(paragraphMinContentWidth),
  );
  let baseline: number | "unavailable" | undefined;
  if (style.verticalAlign === undefined || style.verticalAlign === "top") {
    if (paragraphs[0] === undefined) {
      baseline = undefined;
    } else {
      try {
        const val = paragraphBaseline(paragraphs[0]);
        baseline = val ?? "unavailable";
      } catch {
        baseline = "unavailable";
      }
    }
  }
  return {
    width: measured.width,
    height: measured.height,
    minContentWidth,
    minContentHeight: measured.height,
    ...(baseline === undefined ? {} : { baseline }),
  };
}

/** Returns layout dimensions based on actual fonts and line breaking rules of shared text service. */
export function measureTextLayout(
  children: readonly unknown[],
  styleValue: unknown,
  theme: ThemeContext,
  availableWidth?: number,
  slideNumber = 1,
): TextLayoutMeasurement {
  const style = readTextStyle(styleValue, "Text.style", theme);
  const paragraphs = normalizeParagraphs(children, style, theme, {
    slideNumber,
  });
  return measureTextLayoutFromParagraphs(paragraphs, style, availableWidth);
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
      "padding must be a number, a 2-tuple, or a 4-tuple",
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
        `${name} inset must be a finite non-negative pixel number`,
      );
    }
    pxToEmu(amount);
  }
  return result;
}

/** Parses text children, validates element structure, and resolves paragraphs and text styles once. */
export function parseTextContent(
  children: readonly unknown[],
  styleValue: unknown,
  theme: ThemeContext,
  fields: TextFieldContext = { slideNumber: 1 },
  label = "Text.style",
): {
  readonly paragraphs: readonly Paragraph[];
  readonly style: ResolvedTextStyle;
} {
  const style = readTextStyle(styleValue, label, theme);
  const paragraphs = normalizeParagraphs(children, style, theme, fields);
  return { paragraphs, style };
}

export const parseTextParagraphs = parseTextContent;

/** Generates PresentationML txBody XML and resource claims directly from pre-parsed paragraphs and styles. */
export function writeTextBodyFromParagraphs(
  paragraphs: readonly Paragraph[],
  style: ResolvedTextStyle | TextStyle,
  theme: ThemeContext,
  verticalAlign: TextStyle["verticalAlign"] = "top",
  layout?: TextLayoutRequest,
  destination: TextBodyDestination = "text-box",
  embedFonts?: FontEmbeddingMode,
  createFieldId?: () => string,
): TextServiceResult {
  const resolvedParagraphs =
    createFieldId !== undefined
      ? paragraphs.map((paragraph) => ({
          ...paragraph,
          runs: paragraph.runs.map((run) =>
            run.field !== undefined
              ? {
                  ...run,
                  field: {
                    ...run.field,
                    id: createFieldId(),
                  },
                }
              : run,
          ),
        }))
      : paragraphs;
  const runs = resolvedParagraphs.flatMap((paragraph) =>
    paragraph.runs.length > 0
      ? paragraph.runs
      : [{ text: "", style: paragraph.endStyle }],
  );
  const listRuns = resolvedParagraphs.flatMap((paragraph) => {
    const list = paragraph.style.list;
    if (list === undefined) return [];
    const characters =
      list.type === "bullet" ? (list.character ?? "•") : "0123456789.";
    return [
      {
        text: characters,
        style: paragraph.runs[0]?.style ?? paragraph.endStyle,
      },
    ];
  });
  const { resourceClaims, fontUses, warnings } = textResourceClaims(
    [...runs, ...listRuns],
    embedFonts,
  );
  let measured: ReturnType<typeof measureParagraphs> | undefined;
  if (layout !== undefined) {
    try {
      measured = measureParagraphs(
        resolvedParagraphs,
        layout.width === "max-content" ? Number.MAX_VALUE : layout.width,
      );
    } catch (error) {
      if (layout.contentSized === true) throw error;
      if (error instanceof CapabilityError) {
        if (!warnings.some((warning) => warning.code === error.code)) {
          warnings.push({
            code: error.code,
            message: `Failed to complete text measurement using provided fonts: ${error.message}`,
          });
        }
      } else {
        warnings.push({
          code: "font-measurement-unavailable",
          message: `Failed to complete text measurement using provided fonts: ${error instanceof Error ? error.message : String(error)}`,
        });
      }
    }
  }
  if (layout?.contentSized === true && measured === undefined) {
    throw new CapabilityError(
      "font-measurement-unavailable",
      "Content-sized text is missing complete font measurement basis",
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
        "Based on provided font metrics, text overflows text box width or height; build preserves original content and dimensions",
    });
  }
  const outputParagraphs =
    layout?.contentSized === true && measured !== undefined
      ? resolvedParagraphs.map((paragraph, index) => ({
          ...paragraph,
          style: {
            ...paragraph.style,
            lineHeight: measured.paragraphs[index]!.lineHeight,
          },
        }))
      : resolvedParagraphs;
  const text = outputParagraphs
    .map(({ runs }) => runs.map(({ text: part }) => part).join(""))
    .join("\n");
  const xml =
    destination === "table-cell"
      ? `<a:txBody><a:bodyPr/><a:lstStyle/>${outputParagraphs.map((paragraph) => paragraphXml(paragraph, theme)).join("")}</a:txBody>`
      : `<p:txBody><a:bodyPr wrap="${style.textWrap === "nowrap" ? "none" : "square"}" lIns="${pxToEmu(layout?.insets?.left ?? 0)}" tIns="${pxToEmu(layout?.insets?.top ?? 0)}" rIns="${pxToEmu(layout?.insets?.right ?? 0)}" bIns="${pxToEmu(layout?.insets?.bottom ?? 0)}" anchor="${bodyAnchor(style.verticalAlign ?? verticalAlign)}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${outputParagraphs.map((paragraph) => paragraphXml(paragraph, theme)).join("")}</p:txBody>`;
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

/** Shares paragraph, run, list, and font resource processing for text boxes and shape text. */
export function writeTextBody(
  children: readonly unknown[],
  styleValue: unknown,
  theme: ThemeContext,
  verticalAlign: TextStyle["verticalAlign"] = "top",
  layout?: TextLayoutRequest,
  destination: TextBodyDestination = "text-box",
  fields: TextFieldContext = { slideNumber: 1 },
  embedFonts?: FontEmbeddingMode,
): TextServiceResult {
  const style = readTextStyle(
    destination === "table-cell"
      ? { ...styleRecord(styleValue, "Text.style"), textWrap: "normal" }
      : styleValue,
    "Text.style",
    theme,
  );
  const paragraphs = normalizeParagraphs(children, style, theme, fields);
  return writeTextBodyFromParagraphs(
    paragraphs,
    style,
    theme,
    verticalAlign,
    layout,
    destination,
    embedFonts,
  );
}

/** Generates shared font styles for native chart text and registers font claims based on actual categories, series, and label text. */
export function writeChartTextProperties(
  styleValue: unknown,
  texts: readonly string[],
  label: string,
  theme: ThemeContext,
  embedFonts?: FontEmbeddingMode,
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
    letterSpacing: style.letterSpacing,
    textTransform: style.textTransform,
    textDecoration: style.textDecoration,
    textShadow: style.textShadow,
    lang: style.lang,
  };
  validateRunStyle(runStyle);
  const runs = texts.map((text) => ({ text, style: runStyle }));
  for (const text of texts) escapedXml(text);
  const { resourceClaims, fontUses, warnings } = textResourceClaims(
    runs,
    embedFonts,
  );
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
  const endParaLang =
    runStyle.lang !== undefined ? ` lang="${escapedXml(runStyle.lang)}"` : "";
  return {
    xml: `<c:txPr><a:bodyPr anchor="${bodyAnchor(style.verticalAlign)}"/><a:lstStyle/><a:p><a:pPr${paragraphAlignment}>${lineSpacing}${spacingAfter}${defaultProperties}</a:pPr><a:endParaRPr${endParaLang}/></a:p></c:txPr>`,
    text: outputText,
    resourceClaims,
    fontUses,
    warnings,
  };
}

/** Emits a native chart or axis title with actual run properties, not canvas text. */
export function writeChartTitle(
  text: string,
  styleValue: unknown,
  label: string,
  theme: ThemeContext,
  embedFonts?: FontEmbeddingMode,
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
    letterSpacing: style.letterSpacing,
    textTransform: style.textTransform,
    textDecoration: style.textDecoration,
    textShadow: style.textShadow,
    lang: style.lang,
  };
  validateRunStyle(runStyle);
  const escapedText = escapedXml(text);
  const { resourceClaims, fontUses, warnings } = textResourceClaims(
    [{ text, style: runStyle }],
    embedFonts,
  );
  const runProperties = runPropertiesXml(runStyle, "a:rPr", theme);
  const defaultProperties = runPropertiesXml(runStyle, "a:defRPr", theme);
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
  const endParaLang =
    runStyle.lang !== undefined ? ` lang="${escapedXml(runStyle.lang)}"` : "";
  const bodyProperties = `<a:bodyPr anchor="${bodyAnchor(style.verticalAlign)}"/>`;
  return {
    xml: `<c:title><c:tx><c:rich>${bodyProperties}<a:lstStyle/><a:p><a:pPr${paragraphAlignment}>${lineSpacing}${spacingAfter}${defaultProperties}</a:pPr><a:r>${runProperties}<a:t>${escapedText}</a:t></a:r><a:endParaRPr${endParaLang}/></a:p></c:rich></c:tx><c:overlay val="0"/><c:txPr>${bodyProperties}<a:lstStyle/><a:p><a:pPr${paragraphAlignment}>${lineSpacing}${spacingAfter}${defaultProperties}</a:pPr><a:endParaRPr${endParaLang}/></a:p></c:txPr></c:title>`,
    text,
    resourceClaims,
    fontUses,
    warnings,
  };
}

/** Converts fixed pixel length into DrawingML EMU. */
export function textExtent(value: number, label: string): number {
  if (!Number.isFinite(value) || value <= 0) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} must be a finite positive pixel value`,
    );
  }
  const extent = pxToEmu(value);
  if (extent < 1 || extent > 2_147_483_647) {
    throw new CapabilityError(
      "invalid-text-box",
      `${label} converted value exceeds shape dimension range`,
    );
  }
  return extent;
}

/** Reads raw plain text of fixed-area text, shared between shape and read-only report. */
export function plainText(children: readonly unknown[]): string {
  const result: string[] = [];
  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError(
        "invalid-text-child",
        "Text content tree contains invalid nodes",
      );
    }
    if (child.type === "#text") result.push(child.text);
    else if (child.type === "rp:span" || child.type === "rp:paragraph") {
      result.push(plainText(child.children));
    } else {
      throw new CapabilityError(
        "unsupported-text-child",
        `${child.type} is not text content`,
      );
    }
  }
  return result.join("");
}
