import { createHash } from "node:crypto";

import {
  CapabilityError,
  paddingStyleKeys,
  textStyleKeys,
  validateTextStyle,
  type CapabilityFontUsageClaim,
  type CapabilityPresentationContribution,
  type ColorValue,
  type Font,
  type FontFaceStyle,
  type FontReference,
  type FontSize,
  type GeometryBox,
  type PlaceholderResolutionContext,
  type ResolvedColor,
  type Shadow,
  type TextAlignment,
  type TextMetrics,
  type TextStyle,
  type ThemeContext,
  type UnmeasuredFontIssue,
} from "@pptxdk/core/capability";
import { pxToEmu, resolvedColorXml, xmlEscape } from "@pptxdk/emit";
import * as fontkit from "fontkit";
import * as hb from "harfbuzzjs";
// @ts-expect-error -- linebreak does not provide types; declare the breakpoint interface used below.
import LineBreakerImplementation from "linebreak";

import {
  fontPresentationContributions,
  isFontEmbeddable,
  requiresFontGlyph,
  requiresVisibleFontGlyph,
  type FontEmbeddingMode,
  type FontEmbeddingUse,
} from "../font/font-embedding.ts";
import { matchesFontFamily } from "../font/font-family.ts";
import type {} from "./emscripten.d.ts";
import {
  hasLocalDeclaration,
  localDeclaredKeySet,
  resolvePlaceholderBaseTextStyle,
  resolvePlaceholderLevelTextStyle,
} from "./style-inheritance.ts";
import {
  paragraphStyleSchema,
  parseParagraphStyle,
  parseRunStyle,
  parseTextExtent,
  parseTextInsets,
  runStyleSchema,
  textListSchema,
  type ParagraphStyle,
  type RunTextStyle,
  type TextList,
} from "./text-schemas.ts";
import { resolveShadow, textShadowXml, type ResolvedShadow } from "./visual.ts";
export { requiresFontGlyph } from "../font/font-embedding.ts";
export { paragraphStyleSchema, runStyleSchema, textListSchema };
export type {
  FontReference,
  FontSize,
  TextAlignment,
  TextStyle,
} from "@pptxdk/core/capability";
export type { ParagraphStyle, RunTextStyle, TextList };

export type TextResourceClaim = {
  readonly kind: "font-face";
  readonly key: string;
  readonly identity: string;
  /** Text characters actually emitted for this font style, reused for subsequent font processing. */
  readonly characters: string;
};

/** Resolve-stage text facts cached inside capability closures for independent emit and paint. */
export type ResolvedTextBody = {
  readonly paragraphs: readonly Paragraph[];
  readonly style: ResolvedTextStyle | TextStyle;
  readonly verticalAlign: TextStyle["verticalAlign"];
  readonly layout?: TextLayoutRequest;
  readonly destination: TextBodyDestination;
  readonly isPlaceholder?: boolean;
  readonly localBodyDeclared?: ReadonlySet<string>;
  readonly text: string;
  readonly paragraphCount: number;
  readonly resourceClaims: readonly TextResourceClaim[];
  readonly fontUses: readonly FontEmbeddingUse[];
  readonly measurement?: { readonly width: number; readonly height: number };
  readonly measuredLayout?: MeasuredTextLayout;
  readonly warnings: readonly {
    readonly code: string;
    readonly message: string;
  }[];
  readonly missingGlyphErrors: readonly string[];
  readonly textMetrics?: TextMetrics;
  readonly fontUsages: readonly CapabilityFontUsageClaim[];
  readonly unmeasuredIssues: {
    readonly missingResourceFonts: readonly UnmeasuredFontIssue[];
    readonly missingGlyphFonts: readonly UnmeasuredFontIssue[];
    readonly referencedFonts: readonly UnmeasuredFontIssue[];
  };
};

export type EmittedTextBody = {
  readonly xml: string;
  readonly presentationContributions: readonly CapabilityPresentationContribution[];
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
  readonly type: string;
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
  readonly localDeclared?: RunStyle;
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
  readonly localStyle?: ParagraphStyle;
  readonly endStyle: RunStyle;
  readonly localEndStyle?: RunStyle;
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

/** UTF-16 ranges retain the formal wrapping decision for derived static drawing. */
export type MeasuredLine = {
  readonly start: number;
  readonly end: number;
  readonly width: number;
};

export type MeasuredParagraph = {
  readonly lineCount: number;
  readonly lines: readonly MeasuredLine[];
  readonly text: string;
  readonly pieces: readonly MeasurePiece[];
  readonly width: number;
  readonly lineHeight: number;
  readonly spacingAfter: number;
};

export type MeasuredTextLayout = {
  readonly width: number;
  readonly height: number;
  readonly paragraphs: readonly MeasuredParagraph[];
};

export type ResolvedTextDecoration = "none" | "underline" | "line-through";
export type ResolvedTextTransform = "none" | "uppercase";
export type ResolvedTextAlign = "left" | "center" | "right" | "justify";
export type ResolvedVerticalAlign = "top" | "middle" | "bottom";
export type ResolvedTextWrap = "normal" | "nowrap";
export type ResolvedTextList =
  | {
      readonly type: "bullet";
      readonly character: string;
      readonly level: number;
    }
  | {
      readonly type: "number";
      readonly style: "arabicPeriod";
      readonly startAt: number;
      readonly level: number;
    };
export type ResolvedTextShadow =
  | { readonly kind: "none" }
  | { readonly kind: "shadow"; readonly shadow: ResolvedShadow };

export type PreviewDecoration = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly color: RunStyle["color"];
  readonly textShadow?: Shadow | "none" | undefined;
  readonly glyphStart?: number;
  readonly glyphEnd?: number;
};

export type PreviewGlyph = {
  readonly path: string;
  readonly x: number;
  readonly y: number;
  readonly scale: number;
  readonly color: RunStyle["color"];
  readonly textShadow?: Shadow | "none" | undefined;
  readonly isSpace?: boolean;
  readonly clusterKey?: string;
};

export type PreviewLineOutlines = {
  readonly start: number;
  readonly end: number;
  readonly width: number;
  readonly glyphs: readonly PreviewGlyph[];
  readonly decorations?: readonly PreviewDecoration[];
  readonly ascent: number;
  readonly descent: number;
  readonly advance: number;
  readonly canJustify?: boolean;
};

export type PreviewParagraphOutlines = {
  readonly lines: readonly PreviewLineOutlines[];
  readonly lineHeight: number;
  readonly spacingAfter: number;
  readonly baseDir: "ltr" | "rtl";
};

export type PreviewTextOutlinesLine = {
  readonly start: number;
  readonly end: number;
  readonly baseline: number;
  readonly glyphs: readonly PreviewGlyph[];
  readonly decorations?: readonly PreviewDecoration[];
};

type LineBreakerConstructor = new (text: string) => {
  nextBreak(): { readonly position: number; readonly required: boolean } | null;
};

const LineBreaker = LineBreakerImplementation as LineBreakerConstructor;
const graphemes = new Intl.Segmenter(undefined, { granularity: "grapheme" });
const measureFaces = new Map<string, fontkit.Font>();
type CachedHbFace = {
  readonly blob: hb.Blob;
  readonly face: hb.Face;
  readonly font: hb.Font;
};
const hbFaces = new Map<string, CachedHbFace>();
let sharedHbBuffer: hb.Buffer | undefined;
let sharedHbBufferBusy = false;

function withHbBuffer<T>(fn: (buffer: hb.Buffer) => T): T {
  if (!sharedHbBufferBusy) {
    sharedHbBufferBusy = true;
    sharedHbBuffer ??= new hb.Buffer();
    try {
      return fn(sharedHbBuffer);
    } finally {
      sharedHbBuffer.reset();
      sharedHbBufferBusy = false;
    }
  }
  const buffer = new hb.Buffer();
  try {
    return fn(buffer);
  } finally {
    buffer.reset();
  }
}

const defaultTextStyle: TextStyle = Object.freeze({
  fontFamily: "minor",
  fontSize: 24,
  fontWeight: "normal",
  fontStyle: "normal",
  color: "tx1",
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

function isHostText(child: TextChild): child is HostText {
  return child.type === "#text" && "text" in child;
}

function hostChild(value: unknown): value is TextChild {
  if (!record(value) || typeof value["type"] !== "string") return false;
  if (value["type"] === "#text") return typeof value["text"] === "string";
  return (
    value["type"].trim() !== "" &&
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
  return value;
}

function fontFamilyValue(value: FontReference | undefined): Font {
  if (value === undefined) return { family: "Aptos" };
  return typeof value === "string" ? { family: value } : value;
}

function normalizedFonts(style: RunStyle): NormalizedFont[] {
  const family = fontFamilyValue(style.fontFamily);
  const weight = weightValue(style.fontWeight);
  const faceStyle = style.fontStyle ?? "normal";
  const faces = family.faces ?? [];
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
  const { themeFont: _themeFont, color, ...rest } = style;
  validateTextStyle({
    ...rest,
    ...(color !== undefined && !isResolvedColor(color) ? { color } : {}),
  });
  return style;
}

function readRunStyle(value: unknown): RunStyle {
  return parseRunStyle(value, "Span.style");
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

function mergeLocalRunDeclared(
  parent: RunStyle,
  child: Readonly<Record<string, unknown>>,
  theme: ThemeContext,
): RunStyle {
  const declared: Record<string, unknown> = { ...parent };
  for (const [key, value] of Object.entries(child)) {
    if (!Object.hasOwn(child, key) || value === undefined) continue;
    declared[key] = value;
  }
  if (hasLocalDeclaration(child, "fontFamily")) {
    const resolved = theme.resolveFont(child["fontFamily"] as FontReference);
    declared["fontFamily"] = resolved.font ?? resolved.family;
    if (resolved.themeFont === undefined) delete declared["themeFont"];
    else declared["themeFont"] = resolved.themeFont;
  }
  if (hasLocalDeclaration(child, "color") && !isResolvedColor(child["color"])) {
    declared["color"] = theme.resolveColor(child["color"] as ColorValue);
  }
  return declared;
}

function resolveRunStyleFromRecord(
  style: Readonly<Record<string, unknown>>,
  theme: ThemeContext,
): RunStyle {
  const font = theme.resolveFont(
    (style["fontFamily"] as FontReference | undefined) ??
      defaultTextStyle.fontFamily!,
  );
  const rawColor = style["color"] as ColorValue | ResolvedColor | undefined;
  const color =
    rawColor === undefined
      ? theme.resolveColor(defaultTextStyle.color!)
      : isResolvedColor(rawColor)
        ? rawColor
        : theme.resolveColor(rawColor);
  const resolved: RunStyle = {
    fontFamily: font.font ?? font.family,
    fontSize:
      (style["fontSize"] as FontSize | undefined) ?? defaultTextStyle.fontSize,
    fontWeight:
      (style["fontWeight"] as TextStyle["fontWeight"] | undefined) ??
      defaultTextStyle.fontWeight,
    fontStyle:
      (style["fontStyle"] as FontFaceStyle | undefined) ??
      defaultTextStyle.fontStyle,
    color,
    ...(font.themeFont === undefined ? {} : { themeFont: font.themeFont }),
    ...(style["letterSpacing"] === undefined
      ? {}
      : { letterSpacing: style["letterSpacing"] as number }),
    ...(style["textTransform"] === undefined
      ? {}
      : {
          textTransform: style["textTransform"] as TextStyle["textTransform"],
        }),
    ...(style["textDecoration"] === undefined
      ? {}
      : {
          textDecoration: style[
            "textDecoration"
          ] as TextStyle["textDecoration"],
        }),
    ...(style["textShadow"] === undefined || style["textShadow"] === "none"
      ? {}
      : { textShadow: style["textShadow"] as Shadow }),
    ...(style["lang"] === undefined ? {} : { lang: style["lang"] as string }),
  };
  return validateRunStyle(resolved);
}

function readPlaceholderTextStyle(
  value: unknown,
  label: string,
  placeholder: PlaceholderResolutionContext,
): {
  readonly style: ResolvedTextStyle;
  readonly localDeclaredStyle: Record<string, unknown>;
} {
  const props = styleRecord(value, label);
  onlyKnownKeys(props, [...textStyleKeys], label);
  const localDeclaredStyle: Record<string, unknown> = {};
  for (const key of textStyleKeys) {
    if (hasLocalDeclaration(props, key)) {
      localDeclaredStyle[key] = props[key];
    }
  }
  const effectiveBase = resolvePlaceholderBaseTextStyle(
    placeholder,
    localDeclaredStyle,
  );
  const masterTheme = placeholder.masterTheme;
  const merged = {
    ...defaultTextStyle,
    ...effectiveBase,
  } as TextStyle;
  const font = masterTheme.resolveFont(
    merged.fontFamily ?? defaultTextStyle.fontFamily!,
  );
  if (merged.color !== undefined && !isResolvedColor(merged.color)) {
    masterTheme.resolveColor(merged.color);
  }
  validateTextStyle({
    ...merged,
    fontFamily: font.font ?? font.family,
    ...(isResolvedColor(merged.color) ? { color: undefined } : {}),
  });
  const result = {
    ...merged,
    fontFamily: font.font ?? font.family,
    ...(font.themeFont === undefined ? {} : { themeFont: font.themeFont }),
  };
  if (result.textShadow === "none") {
    delete (result as { textShadow?: Shadow | "none" }).textShadow;
  }
  return { style: result, localDeclaredStyle };
}

function readTextStyle(
  value: unknown,
  label: string,
  theme: ThemeContext,
): ResolvedTextStyle {
  const props = styleRecord(value, label);
  onlyKnownKeys(props, [...textStyleKeys], label);
  const style = {
    ...defaultTextStyle,
    ...theme.textStyle,
    ...props,
  } as TextStyle;
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
  validateTextStyle({
    ...style,
    fontFamily: font.font ?? font.family,
    ...(isResolvedColor(style.color) ? { color: undefined } : {}),
  });
  const result = {
    ...style,
    fontFamily: font.font ?? font.family,
    ...(font.themeFont === undefined ? {} : { themeFont: font.themeFont }),
  };
  if (result.textShadow === "none") {
    delete (result as { textShadow?: Shadow | "none" }).textShadow;
  }
  return result;
}

function paragraphStyle(value: unknown): ParagraphStyle {
  return parseParagraphStyle(value, "P.style");
}

function collectRuns(
  children: readonly unknown[],
  style: RunStyle,
  runs: Run[],
  theme: ThemeContext,
  fields: TextFieldContext,
  localDeclared?: RunStyle,
): void {
  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError(
        "invalid-text-child",
        "Text content tree contains invalid nodes",
      );
    }
    if (isHostText(child)) {
      if (child.text !== "") {
        runs.push({
          text: child.text,
          style,
          ...(localDeclared === undefined ? {} : { localDeclared }),
        });
      }
      continue;
    }
    if (child.type !== "span" && child.type !== "slide-number") {
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
    const nextLocalDeclared =
      localDeclared === undefined
        ? undefined
        : mergeLocalRunDeclared(localDeclared, spanStyle, theme);
    if (child.type === "slide-number") {
      if (child.children.length > 0) {
        throw new CapabilityError(
          "invalid-slide-number",
          "SlideNumber cannot contain children",
        );
      }
      runs.push({
        text: String(fields.slideNumber),
        style: nextStyle,
        ...(nextLocalDeclared === undefined
          ? {}
          : { localDeclared: nextLocalDeclared }),
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
    collectRuns(
      child.children,
      nextStyle,
      runs,
      theme,
      fields,
      nextLocalDeclared,
    );
  }
}

function resolveDecoration(
  value: TextStyle["textDecoration"] | undefined,
): ResolvedTextDecoration {
  return value ?? "none";
}

function resolveTransform(
  value: TextStyle["textTransform"] | undefined,
): ResolvedTextTransform {
  return value ?? "none";
}

function resolveTextShadowModel(
  theme: ThemeContext,
  value: Shadow | "none" | undefined,
): ResolvedTextShadow {
  if (value === undefined || value === "none") {
    return { kind: "none" };
  }
  return {
    kind: "shadow",
    shadow: resolveShadow(theme, value, "textShadow"),
  };
}

function resolveAlign(
  value: TextAlignment | undefined,
  fallback: ResolvedTextAlign = "left",
): ResolvedTextAlign {
  return value ?? fallback;
}

function resolveVerticalAlign(
  value: TextStyle["verticalAlign"] | undefined,
): ResolvedVerticalAlign {
  return value ?? "top";
}

function resolveListModel(
  value: TextList | undefined,
): ResolvedTextList | undefined {
  if (value === undefined) return undefined;
  switch (value.type) {
    case "bullet":
      return {
        type: "bullet",
        character: value.character ?? "•",
        level: value.level ?? 0,
      };
    case "number":
      return {
        type: "number",
        style: value.style ?? "arabicPeriod",
        startAt: value.startAt ?? 1,
        level: value.level ?? 0,
      };
    default: {
      const exhaustive: never = value;
      return exhaustive;
    }
  }
}

function decorationAttribute(decoration: ResolvedTextDecoration): string[] {
  switch (decoration) {
    case "none":
      return [];
    case "underline":
      return ['u="sng"'];
    case "line-through":
      return ['strike="sngStrike"'];
    default: {
      const exhaustive: never = decoration;
      return exhaustive;
    }
  }
}

function transformAttribute(transform: ResolvedTextTransform): string[] {
  switch (transform) {
    case "none":
      return [];
    case "uppercase":
      return ['cap="all"'];
    default: {
      const exhaustive: never = transform;
      return exhaustive;
    }
  }
}

function runFontSlotsXml(style: RunStyle): string {
  const font = normalizedFonts(style)[0]!;
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
  return `<a:latin typeface="${escapedXml(latinTypeface)}"/><a:ea typeface="${escapedXml(eaTypeface)}"/><a:cs typeface="${escapedXml(csTypeface)}"/>`;
}

function runPropertiesXml(
  style: RunStyle,
  tag: "a:rPr" | "a:defRPr",
  theme: ThemeContext,
  differential = false,
): string {
  if (differential) {
    const attributes: string[] = [];
    if (style.lang !== undefined) {
      attributes.push(`lang="${escapedXml(style.lang)}"`);
    }
    if (style.fontSize !== undefined) {
      attributes.push(`sz="${centipoints(fontSizePoints(style.fontSize))}"`);
    }
    if (style.fontWeight !== undefined) {
      const font = normalizedFonts(style)[0]!;
      attributes.push(`b="${outputBold(font) ? 1 : 0}"`);
    }
    if (style.fontStyle !== undefined) {
      attributes.push(`i="${style.fontStyle === "italic" ? 1 : 0}"`);
    }
    if (style.textDecoration !== undefined) {
      attributes.push(
        ...decorationAttribute(resolveDecoration(style.textDecoration)),
      );
    }
    if (style.textTransform !== undefined) {
      attributes.push(
        ...transformAttribute(resolveTransform(style.textTransform)),
      );
    }
    if (style.letterSpacing !== undefined) {
      attributes.push(`spc="${letterSpacingCentipoints(style.letterSpacing)}"`);
    }
    const fillXml =
      style.color === undefined
        ? ""
        : `<a:solidFill>${resolvedColorXml(
            isResolvedColor(style.color)
              ? style.color
              : theme.resolveColor(style.color),
          )}</a:solidFill>`;
    const shadowXml =
      style.textShadow === undefined
        ? ""
        : textShadowXml(theme, style.textShadow);
    const fontXml =
      style.fontFamily === undefined && style.themeFont === undefined
        ? ""
        : runFontSlotsXml(style);
    const attrText = attributes.length === 0 ? "" : ` ${attributes.join(" ")}`;
    const childrenXml = `${fillXml}${shadowXml}${fontXml}`;
    return childrenXml === ""
      ? `<${tag}${attrText}/>`
      : `<${tag}${attrText}>${childrenXml}</${tag}>`;
  }
  const font = normalizedFonts(style)[0]!;
  const points = fontSizePoints(style.fontSize);
  const bold = outputBold(font);
  const color = isResolvedColor(style.color)
    ? style.color
    : theme.resolveColor(style.color ?? "dk1");
  const fontSlots = runFontSlotsXml(style);
  const decoration = resolveDecoration(style.textDecoration);
  const transform = resolveTransform(style.textTransform);
  const attributes = [
    ...(style.lang !== undefined ? [`lang="${escapedXml(style.lang)}"`] : []),
    `sz="${centipoints(points)}"`,
    `b="${bold ? 1 : 0}"`,
    `i="${font.style === "italic" ? 1 : 0}"`,
    ...decorationAttribute(decoration),
    ...transformAttribute(transform),
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

/** Emits a native `<a:fld type="slidenum">` run using a fresh or preallocated field GUID. */
export function emitSlideNumberRunXml(
  run: Run,
  theme: ThemeContext,
  createFieldId?: () => string,
  differential = false,
): string {
  const fieldId =
    createFieldId !== undefined
      ? createFieldId()
      : (run.field?.id ?? "00000000-0000-4000-8000-000000000001");
  const fieldType = run.field?.type ?? "slidenum";
  const rPr = differential
    ? runPropertiesXml(run.localDeclared ?? {}, "a:rPr", theme, true)
    : runPropertiesXml(run.style, "a:rPr", theme, false);
  return `<a:fld id="${escapedXml(fieldId)}" type="${fieldType}">${rPr}<a:t>${escapedXml(run.text)}</a:t></a:fld>`;
}

function runXml(run: Run, theme: ThemeContext, differential = false): string {
  if (run.field !== undefined) {
    return emitSlideNumberRunXml(run, theme, undefined, differential);
  }
  const text = escapedXml(run.text);
  const rPr = differential
    ? runPropertiesXml(run.localDeclared ?? {}, "a:rPr", theme, true)
    : runPropertiesXml(run.style, "a:rPr", theme, false);
  return `<a:r>${rPr}<a:t xml:space="preserve">${text}</a:t></a:r>`;
}

function textResourceClaims(runs: readonly Run[]): {
  readonly resourceClaims: TextResourceClaim[];
  readonly fontUses: FontEmbeddingUse[];
  readonly fontUsages: CapabilityFontUsageClaim[];
  readonly warnings: { readonly code: string; readonly message: string }[];
  readonly missingGlyphErrors: string[];
  readonly unmeasuredIssues: {
    readonly missingResourceFonts: readonly UnmeasuredFontIssue[];
    readonly missingGlyphFonts: readonly UnmeasuredFontIssue[];
    readonly referencedFonts: readonly UnmeasuredFontIssue[];
  };
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
  const fontUsages: CapabilityFontUsageClaim[] = [];
  const warnings: { code: string; message: string }[] = [];
  const missingGlyphErrors: string[] = [];
  const missingResourceMap = new Map<string, UnmeasuredFontIssue>();
  const missingGlyphSets = new Map<string, Set<string>>();
  const missingGlyphFonts = new Map<
    string,
    {
      readonly family: string;
      readonly weight: number;
      readonly style: "normal" | "italic";
    }
  >();
  const referencedFontMap = new Map<string, UnmeasuredFontIssue>();
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
      if (!referencedFontMap.has(key)) {
        referencedFontMap.set(key, {
          family: font.family,
          weight: font.weight,
          style: font.style,
        });
      }
      if (!font.hasResource) {
        fontUsages.push({
          family: font.family,
          weight: font.weight,
          style: font.style,
          hasResource: false,
          ...(font.substitute ? { substitute: true } : {}),
        });
        if (!missingResourceMap.has(key)) {
          missingResourceMap.set(key, {
            family: font.family,
            weight: font.weight,
            style: font.style,
          });
        }
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
        const missing =
          face !== undefined
            ? missingCharacters(face, entry.text.join(""))
            : [];
        const embeddable = isFontEmbeddable(
          font.data!,
          font.substitute,
          font.embeddable,
        );
        fontUsages.push({
          family: font.family,
          weight: font.weight,
          style: font.style,
          hasResource: true,
          ...(font.substitute ? { substitute: true } : {}),
          embeddable,
          ...(missing.length > 0 ? { missingGlyphs: missing } : {}),
        });
        if (missing.length > 0) {
          const set = missingGlyphSets.get(key) ?? new Set<string>();
          for (const ch of missing) set.add(ch);
          missingGlyphSets.set(key, set);
          if (!missingGlyphFonts.has(key)) {
            missingGlyphFonts.set(key, {
              family: font.family,
              weight: font.weight,
              style: font.style,
            });
          }
          const shown = missing.slice(0, 10).join(" ");
          const extra = missing.length - 10;
          const message = `${font.family} is missing glyphs for characters ${shown}${extra > 0 ? `, and ${extra} more characters` : ""}`;
          warnings.push({
            code: "font-missing-glyph",
            message,
          });
          missingGlyphErrors.push(message);
        }
      }
    }
  }
  const unmeasuredIssues = {
    missingResourceFonts: [...missingResourceMap.values()],
    missingGlyphFonts: [...missingGlyphFonts.entries()].map(
      ([key, descriptor]) => ({
        ...descriptor,
        missingGlyphs: [...(missingGlyphSets.get(key) ?? [])],
      }),
    ),
    referencedFonts: [...referencedFontMap.values()],
  };
  return {
    resourceClaims,
    fontUses,
    fontUsages,
    warnings,
    missingGlyphErrors,
    unmeasuredIssues,
  };
}

function alignmentValue(value: TextAlignment | undefined): string | undefined {
  if (value === undefined) return undefined;
  switch (value) {
    case "left":
      return "l";
    case "center":
      return "ctr";
    case "right":
      return "r";
    case "justify":
      return "just";
    default: {
      const exhaustive: never = value;
      return exhaustive;
    }
  }
}

function paragraphXml(
  paragraph: Paragraph,
  theme: ThemeContext,
  differential = false,
): string {
  const style = differential ? (paragraph.localStyle ?? {}) : paragraph.style;
  const { runs } = paragraph;
  const attributes: string[] = [];
  const align = alignmentValue(style.align);
  if (align !== undefined) attributes.push(`algn="${align}"`);
  const paragraphText = runs.map((run) => displayedText(run)).join("");
  if (paragraphBaseDirection(paragraphText) === "rtl") {
    attributes.push('rtl="1"');
  }
  const resolvedList = resolveListModel(style.list);
  if (resolvedList !== undefined)
    attributes.push(`lvl="${resolvedList.level}"`);
  const lineSpacing =
    style.lineHeight === undefined
      ? ""
      : `<a:lnSpc><a:spcPts val="${spacingCentipoints(positiveLength(style.lineHeight, "lineHeight"), "lineHeight")}"/></a:lnSpc>`;
  const spacingAfter =
    style.spacingAfter === undefined || style.spacingAfter === 0
      ? ""
      : `<a:spcAft><a:spcPts val="${spacingCentipoints(style.spacingAfter, "spacingAfter")}"/></a:spcAft>`;
  let list = "";
  if (resolvedList !== undefined) {
    switch (resolvedList.type) {
      case "bullet":
        list = `<a:buChar char="${xmlEscape(resolvedList.character)}"/>`;
        break;
      case "number":
        list = `<a:buAutoNum type="${resolvedList.style}" startAt="${resolvedList.startAt}"/>`;
        break;
      default: {
        const exhaustive: never = resolvedList;
        return exhaustive;
      }
    }
  }
  const pPr =
    attributes.length +
      lineSpacing.length +
      spacingAfter.length +
      list.length ===
    0
      ? ""
      : `<a:pPr${attributes.length === 0 ? "" : ` ${attributes.join(" ")}`}${lineSpacing.length + spacingAfter.length + list.length === 0 ? "/>" : `>${lineSpacing}${spacingAfter}${list}</a:pPr>`}`;
  const endStyle = differential
    ? (paragraph.localEndStyle ?? {})
    : paragraph.endStyle;
  const endProperties =
    runs.length === 0
      ? differential
        ? runPropertiesXml(endStyle, "a:rPr" as const, theme, true).replace(
            /<(\/?)a:rPr\b/gu,
            "<$1a:endParaRPr",
          )
        : `<a:endParaRPr${endStyle.lang !== undefined ? ` lang="${escapedXml(endStyle.lang)}"` : ""} sz="${centipoints(fontSizePoints(endStyle.fontSize))}"/>`
      : "";
  return `<a:p>${pPr}${runs.map((run) => runXml(run, theme, differential)).join("")}${endProperties}</a:p>`;
}

function normalizeParagraphs(
  children: readonly unknown[],
  textStyle: ResolvedTextStyle,
  theme: ThemeContext,
  fields: TextFieldContext,
  placeholder?: PlaceholderResolutionContext,
  localSlideDeclaredStyle?: Readonly<Record<string, unknown>>,
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

  const localSlideRunDeclared: RunStyle | undefined =
    placeholder === undefined
      ? undefined
      : mergeLocalRunDeclared(
          {},
          localSlideDeclaredStyle ?? {},
          placeholder.masterTheme,
        );
  const localSlidePStyle: ParagraphStyle | undefined =
    placeholder === undefined
      ? undefined
      : {
          ...(hasLocalDeclaration(localSlideDeclaredStyle, "lineHeight")
            ? {
                lineHeight: localSlideDeclaredStyle![
                  "lineHeight"
                ] as ParagraphStyle["lineHeight"],
              }
            : {}),
          ...(hasLocalDeclaration(localSlideDeclaredStyle, "align")
            ? {
                align: localSlideDeclaredStyle![
                  "align"
                ] as ParagraphStyle["align"],
              }
            : {}),
          ...(hasLocalDeclaration(localSlideDeclaredStyle, "spacingAfter")
            ? {
                spacingAfter: localSlideDeclaredStyle![
                  "spacingAfter"
                ] as ParagraphStyle["spacingAfter"],
              }
            : {}),
        };

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
        ...(localSlidePStyle === undefined
          ? {}
          : { localStyle: localSlidePStyle }),
        endStyle: baseRunStyle,
        ...(localSlideRunDeclared === undefined
          ? {}
          : { localEndStyle: localSlideRunDeclared }),
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
      isHostText(child) ||
      child.type === "span" ||
      child.type === "slide-number"
    ) {
      collectRuns(
        [child],
        baseRunStyle,
        inline,
        theme,
        fields,
        localSlideRunDeclared,
      );
      continue;
    }
    if (child.type !== "paragraph") {
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
    const localPStyle = paragraphStyle(child.props["style"]);
    const rawLevel = localPStyle.list?.level ?? 0;
    const levelBase =
      placeholder === undefined
        ? undefined
        : resolvePlaceholderLevelTextStyle(
            placeholder,
            rawLevel,
            localSlideDeclaredStyle,
          );
    const effectiveTheme = placeholder?.masterTheme ?? theme;
    const effectiveBaseP: ParagraphStyle =
      levelBase === undefined
        ? baseParagraphStyle
        : {
            ...(levelBase.lineHeight === undefined
              ? {}
              : { lineHeight: levelBase.lineHeight }),
            ...(levelBase.align === undefined
              ? {}
              : { align: levelBase.align }),
            ...(levelBase.spacingAfter === undefined
              ? {}
              : { spacingAfter: levelBase.spacingAfter }),
            ...(levelBase.lang === undefined ? {} : { lang: levelBase.lang }),
            ...(levelBase.list === undefined ? {} : { list: levelBase.list }),
          };
    const levelRunBase: RunStyle =
      levelBase === undefined
        ? baseRunStyle
        : resolveRunStyleFromRecord(levelBase, effectiveTheme);
    const style = {
      ...effectiveBaseP,
      ...localPStyle,
    };
    const paragraphRunStyle = mergeRunStyle(
      levelRunBase,
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
      effectiveTheme,
    );
    const mergedLocalPStyle: ParagraphStyle | undefined =
      placeholder === undefined
        ? undefined
        : { ...localSlidePStyle, ...localPStyle };
    const localPRunDeclared: RunStyle | undefined =
      placeholder === undefined
        ? undefined
        : mergeLocalRunDeclared(
            localSlideRunDeclared ?? {},
            localPStyle,
            effectiveTheme,
          );
    const runs: Run[] = [];
    collectRuns(
      child.children,
      paragraphRunStyle,
      runs,
      effectiveTheme,
      fields,
      localPRunDeclared,
    );
    paragraphs.push({
      runs,
      style,
      ...(mergedLocalPStyle === undefined
        ? {}
        : { localStyle: mergedLocalPStyle }),
      endStyle: paragraphRunStyle,
      ...(localPRunDeclared === undefined
        ? {}
        : { localEndStyle: localPRunDeclared }),
      wrap: textStyle.textWrap ?? "normal",
    });
  }
  flushInline();
  if (paragraphs.length === 0) {
    paragraphs.push({
      runs: [],
      style: baseParagraphStyle,
      ...(localSlidePStyle === undefined
        ? {}
        : { localStyle: localSlidePStyle }),
      endStyle: baseRunStyle,
      ...(localSlideRunDeclared === undefined
        ? {}
        : { localEndStyle: localSlideRunDeclared }),
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
  if (parsed.type === "TTF" && !hbFaces.has(source.identity)) {
    const blob = new hb.Blob(source.data);
    const hbFace = new hb.Face(blob, 0);
    const font = new hb.Font(hbFace);
    hbFaces.set(source.identity, { blob, face: hbFace, font });
  }
  return {
    font: parsed,
    identity: source.identity,
  };
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
    if (!requiresVisibleFontGlyph(character)) continue;
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      (!font.hasGlyphForCodePoint(codePoint) ||
        font.glyphForCodePoint(codePoint).id === 0)
    ) {
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
  let hasVisibleGrapheme = false;
  let pendingTracking = 0;
  for (const piece of pieces) {
    const from = Math.max(start, piece.start);
    const to = Math.min(end, piece.end);
    if (from >= to) continue;
    const source = text.slice(from, to);
    const transformed =
      piece.style.textTransform === "uppercase" ? source.toUpperCase() : source;
    // Drawing and measurement must shape the same visible stream, including control filtering.
    const visible = Array.from(transformed).filter(requiresFontGlyph).join("");
    if (visible === "") continue;
    const shaped = piece.face.font.layout(visible);
    const face = piece.face.font;
    const points = fontSizePoints(piece.style.fontSize);
    const tenths = Math.ceil(centipoints(points) / 10) / 10;
    for (const position of shaped.positions) {
      const advance = (position.xAdvance / face.unitsPerEm) * points;
      natural += advance;
      eighthPoint +=
        Math.round(advance * nativeEighthPointGrid) / nativeEighthPointGrid;
      const tenthAdvance = (position.xAdvance / face.unitsPerEm) * tenths;
      twips += Math.round(tenthAdvance * twipsPerPoint) / twipsPerPoint;
    }
    const pieceTracking = piece.style.letterSpacing ?? 0;
    const visibleCount = Array.from(graphemes.segment(visible)).length;
    if (visibleCount > 0) {
      if (hasVisibleGrapheme) {
        spacing += pendingTracking;
      }
      spacing += pieceTracking * (visibleCount - 1);
      hasVisibleGrapheme = true;
      pendingTracking = pieceTracking;
    }
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
): { readonly lines: readonly MeasuredLine[]; readonly width: number } {
  if (text === "") {
    return { lines: [{ start: 0, end: 0, width: 0 }], width: 0 };
  }
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
    lines: lines.map((line, index) => ({ ...line, width: lineWidths[index]! })),
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

function listIndentMetrics(paragraph: Paragraph): {
  readonly levelIndent: number;
  readonly hangingIndent: number;
  readonly contentIndent: number;
  readonly baseFontPx: number;
} {
  const resolvedList = resolveListModel(paragraph.style.list);
  if (resolvedList === undefined) {
    return {
      levelIndent: 0,
      hangingIndent: 0,
      contentIndent: 0,
      baseFontPx: 0,
    };
  }
  const firstStyle = paragraph.runs[0]?.style ?? paragraph.endStyle;
  const baseFontPx = (fontSizePoints(firstStyle.fontSize) * 4) / 3;
  const hangingIndent = Math.max(baseFontPx * 1.25, 20);
  const levelIndent = resolvedList.level * (baseFontPx * 1.5);
  return {
    levelIndent,
    hangingIndent,
    contentIndent: levelIndent + hangingIndent,
    baseFontPx,
  };
}

function paragraphMinContentWidth(paragraph: Paragraph): number {
  const { text, pieces } = paragraphPieces(paragraph);
  const { contentIndent } = listIndentMetrics(paragraph);
  if (text === "") return contentIndent;
  if (paragraph.wrap === "nowrap") {
    return wrapParagraph(text, pieces, Number.MAX_VALUE).width + contentIndent;
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
  return widest + contentIndent;
}

function paragraphMeasurement(
  paragraph: Paragraph,
  width: number,
): MeasuredParagraph {
  const { text, pieces } = paragraphPieces(paragraph);
  const { contentIndent } = listIndentMetrics(paragraph);
  const wrapWidth =
    paragraph.wrap === "nowrap"
      ? Number.MAX_VALUE
      : Math.max(1, width - contentIndent);
  const wrapped = wrapParagraph(text, pieces, wrapWidth);
  const spacingAfter =
    paragraph.style.spacingAfter === undefined
      ? 0
      : spacingCentipoints(paragraph.style.spacingAfter, "spacingAfter") / 75;
  return {
    lineCount: wrapped.lines.length,
    lines: wrapped.lines,
    text,
    pieces,
    width: wrapped.width + contentIndent,
    lineHeight: paragraphLineHeight(paragraph),
    spacingAfter,
  };
}

function assertParagraphsHaveVisibleGlyphs(
  paragraphs: readonly Paragraph[],
): void {
  for (const paragraph of paragraphs) {
    for (const run of paragraph.runs) {
      const textToCheck =
        displayedText(run) +
        (run.field?.type === "slidenum" ? "0123456789" : "");
      if (!Array.from(textToCheck).some(requiresVisibleFontGlyph)) continue;
      const face = measureRunFace(run.style);
      const missing = missingCharacters(face, textToCheck);
      if (missing.length > 0) {
        const family = normalizedFonts(run.style)[0]?.family ?? "Font";
        const shown = missing.slice(0, 10).join(" ");
        const extra = missing.length - 10;
        throw new CapabilityError(
          "font-missing-glyph",
          `${family} is missing glyphs for characters ${shown}${extra > 0 ? `, and ${extra} more characters` : ""}`,
        );
      }
    }
    const list = paragraph.style.list;
    if (list !== undefined) {
      const markerText =
        list.type === "bullet" ? (list.character ?? "•") : "0123456789.";
      const markerStyle = paragraph.runs[0]?.style ?? paragraph.endStyle;
      const face = measureRunFace(markerStyle);
      const missing = missingCharacters(face, markerText);
      if (missing.length > 0) {
        const family = normalizedFonts(markerStyle)[0]?.family ?? "Font";
        const shown = missing.slice(0, 10).join(" ");
        const extra = missing.length - 10;
        throw new CapabilityError(
          "font-missing-glyph",
          `${family} is missing glyphs for characters ${shown}${extra > 0 ? `, and ${extra} more characters` : ""}`,
        );
      }
    }
  }
}

function measureParagraphs(
  paragraphs: readonly Paragraph[],
  width: number,
): MeasuredTextLayout {
  assertParagraphsHaveVisibleGlyphs(paragraphs);
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

const unsupportedColorFontTables = [
  "COLR",
  "CPAL",
  "sbix",
  "CBDT",
  "CBLC",
  "SVG ",
] as const;

function hasUnsupportedColorFontTable(font: fontkit.Font): boolean {
  const directory = (
    font as {
      readonly directory?: { readonly tables?: Record<string, unknown> };
    }
  ).directory;
  const tables = directory?.tables;
  return unsupportedColorFontTables.some(
    (table) => table in font || (tables !== undefined && table in tables),
  );
}

function roundSvgNumber(value: number): string {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function decorationForSegment(
  decoration: ResolvedTextDecoration,
  face: fontkit.Font,
  fontPx: number,
  startX: number,
  endX: number,
  color: RunStyle["color"],
  textShadow: Shadow | "none" | undefined,
): PreviewDecoration | undefined {
  if (endX <= startX) return undefined;
  const unitsPerEm = face.unitsPerEm;
  switch (decoration) {
    case "none":
      return undefined;
    case "underline": {
      const thickness = Math.max(
        1,
        face.underlineThickness
          ? (face.underlineThickness / unitsPerEm) * fontPx
          : fontPx * 0.07,
      );
      const yCenter = face.underlinePosition
        ? (-face.underlinePosition / unitsPerEm) * fontPx
        : fontPx * 0.12;
      return {
        x: startX,
        y: yCenter - thickness / 2,
        width: endX - startX,
        height: thickness,
        color,
        ...(textShadow === undefined ? {} : { textShadow }),
      };
    }
    case "line-through": {
      const os2 = face["OS/2"];
      const thickness = Math.max(
        1,
        os2?.yStrikeoutSize
          ? (os2.yStrikeoutSize / unitsPerEm) * fontPx
          : fontPx * 0.07,
      );
      const yCenter = -(os2?.yStrikeoutPosition
        ? (os2.yStrikeoutPosition / unitsPerEm) * fontPx
        : fontPx * 0.3);
      return {
        x: startX,
        y: yCenter - thickness / 2,
        width: endX - startX,
        height: thickness,
        color,
        ...(textShadow === undefined ? {} : { textShadow }),
      };
    }
    default: {
      const exhaustive: never = decoration;
      return exhaustive;
    }
  }
}

type BidiType =
  | "L"
  | "R"
  | "AL"
  | "EN"
  | "AN"
  | "NSM"
  | "ES"
  | "CS"
  | "ET"
  | "ON";

function bidiCharType(ch: string): BidiType {
  if (ch === "\u200F") return "R";
  if (ch === "\u200E") return "L";
  if (/^\p{Mark}$/u.test(ch)) return "NSM";
  if (/^[\u0660-\u0669\u06F0-\u06F9]$/u.test(ch)) return "AN";
  if (/^\p{Nd}$/u.test(ch)) return "EN";
  if (
    /^[\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\uFB50-\uFDFF\uFE70-\uFEFF]$/u.test(
      ch,
    ) &&
    /^\p{Letter}$/u.test(ch)
  ) {
    return "AL";
  }
  if (
    /^[\p{Script=Hebrew}\uFB1D-\uFB4F]$/u.test(ch) &&
    /^\p{Letter}$/u.test(ch)
  ) {
    return "R";
  }
  if (/^\p{Letter}$/u.test(ch)) return "L";
  if (ch === "+" || ch === "-" || ch === "\u2212") return "ES";
  if (ch === "," || ch === "." || ch === ":" || ch === "/") return "CS";
  if (ch === "%" || /^\p{Sc}$/u.test(ch)) return "ET";
  return "ON";
}

function paragraphBaseDirection(text: string): "ltr" | "rtl" {
  for (const ch of text) {
    const type = bidiCharType(ch);
    if (type === "R" || type === "AL") return "rtl";
    if (type === "L") return "ltr";
  }
  return "ltr";
}

type CodePointSpan = {
  readonly start: number;
  readonly end: number;
  readonly level: number;
};

function computeLineBidiLevels(
  lineText: string,
  baseOffset: number,
  baseDir: "ltr" | "rtl",
): CodePointSpan[] {
  const cps: {
    readonly start: number;
    readonly end: number;
    readonly ch: string;
  }[] = [];
  let offset = baseOffset;
  for (const ch of lineText) {
    cps.push({ start: offset, end: offset + ch.length, ch });
    offset += ch.length;
  }
  const n = cps.length;
  if (n === 0) return [];

  const embeddingLevel = baseDir === "rtl" ? 1 : 0;
  const dirStrong: "L" | "R" = embeddingLevel === 0 ? "L" : "R";
  const types: BidiType[] = cps.map((cp) => bidiCharType(cp.ch));

  // W1: NSM takes the type of the preceding character (or sos).
  let prevW1: BidiType = dirStrong;
  for (let i = 0; i < n; i += 1) {
    if (types[i] === "NSM") {
      types[i] = prevW1;
    } else {
      prevW1 = types[i]!;
    }
  }

  // W2: EN -> AN when preceded by AL as the last strong type.
  let lastStrongW2: "L" | "R" | "AL" = dirStrong;
  for (let i = 0; i < n; i += 1) {
    const t = types[i]!;
    if (t === "L" || t === "R" || t === "AL") {
      lastStrongW2 = t;
    } else if (t === "EN" && lastStrongW2 === "AL") {
      types[i] = "AN";
    }
  }

  // W3: AL -> R.
  for (let i = 0; i < n; i += 1) {
    if (types[i] === "AL") types[i] = "R";
  }

  // W4: Single ES between EN/EN, single CS between EN/EN or AN/AN.
  for (let i = 1; i < n - 1; i += 1) {
    const t = types[i]!;
    const prev = types[i - 1]!;
    const next = types[i + 1]!;
    if (t === "ES" && prev === "EN" && next === "EN") {
      types[i] = "EN";
    } else if (
      t === "CS" &&
      ((prev === "EN" && next === "EN") || (prev === "AN" && next === "AN"))
    ) {
      types[i] = prev;
    }
  }

  // W5: ET sequences adjacent to EN become EN.
  for (let i = 0; i < n; i += 1) {
    if (types[i] !== "ET") continue;
    let end = i;
    while (end < n && types[end] === "ET") end += 1;
    const prevEn = i > 0 && types[i - 1] === "EN";
    const nextEn = end < n && types[end] === "EN";
    if (prevEn || nextEn) {
      for (let k = i; k < end; k += 1) types[k] = "EN";
    }
    i = end - 1;
  }

  // W6: Remaining ES, ET, CS -> ON.
  for (let i = 0; i < n; i += 1) {
    const t = types[i]!;
    if (t === "ES" || t === "ET" || t === "CS") {
      types[i] = "ON";
    }
  }

  // W7: EN -> L when last strong type was L.
  let lastStrongW7: "L" | "R" = dirStrong;
  for (let i = 0; i < n; i += 1) {
    const t = types[i]!;
    if (t === "L" || t === "R") {
      lastStrongW7 = t;
    } else if (t === "EN" && lastStrongW7 === "L") {
      types[i] = "L";
    }
  }

  // N1 & N2: Neutral runs.
  const neutralEffective = (t: BidiType): "L" | "R" | undefined => {
    if (t === "L") return "L";
    if (t === "R" || t === "EN" || t === "AN") return "R";
    return undefined;
  };
  for (let i = 0; i < n; i += 1) {
    if (types[i] !== "ON") continue;
    let end = i;
    while (end < n && types[end] === "ON") end += 1;
    const leftSide =
      i === 0 ? dirStrong : (neutralEffective(types[i - 1]!) ?? dirStrong);
    const rightSide =
      end === n ? dirStrong : (neutralEffective(types[end]!) ?? dirStrong);
    const resolved: "L" | "R" = leftSide === rightSide ? leftSide : dirStrong;
    for (let k = i; k < end; k += 1) types[k] = resolved;
    i = end - 1;
  }

  // I1 & I2: Resolve final embedding levels.
  return cps.map((cp, index) => {
    const t = types[index]!;
    let level: number;
    if (embeddingLevel === 0) {
      level = t === "L" ? 0 : t === "R" ? 1 : 2;
    } else {
      level = t === "R" ? 1 : 2;
    }
    return { start: cp.start, end: cp.end, level };
  });
}

type VisualLineSegment = {
  readonly text: string;
  readonly piece: MeasurePiece;
  readonly level: number;
  readonly isRtl: boolean;
};

function buildVisualLineSegments(
  text: string,
  lineStart: number,
  trimmedEnd: number,
  pieces: readonly MeasurePiece[],
  baseDir: "ltr" | "rtl",
): VisualLineSegment[] {
  if (lineStart >= trimmedEnd) return [];
  const spans = computeLineBidiLevels(
    text.slice(lineStart, trimmedEnd),
    lineStart,
    baseDir,
  );
  const segments: VisualLineSegment[] = [];
  let spanIdx = 0;
  for (const piece of pieces) {
    const pStart = Math.max(lineStart, piece.start);
    const pEnd = Math.min(trimmedEnd, piece.end);
    if (pStart >= pEnd) continue;
    while (spanIdx < spans.length && spans[spanIdx]!.end <= pStart) {
      spanIdx += 1;
    }
    let segStart = pStart;
    let currentLevel: number | undefined;
    while (spanIdx < spans.length && spans[spanIdx]!.start < pEnd) {
      const sp = spans[spanIdx]!;
      if (currentLevel === undefined) {
        currentLevel = sp.level;
        segStart = Math.max(pStart, sp.start);
      } else if (sp.level !== currentLevel) {
        const segEnd = Math.min(pEnd, sp.start);
        if (segEnd > segStart) {
          segments.push({
            text: text.slice(segStart, segEnd),
            piece,
            level: currentLevel,
            isRtl: currentLevel % 2 === 1,
          });
        }
        currentLevel = sp.level;
        segStart = segEnd;
      }
      if (sp.end <= pEnd) {
        spanIdx += 1;
      } else {
        break;
      }
    }
    if (currentLevel !== undefined && pEnd > segStart) {
      segments.push({
        text: text.slice(segStart, pEnd),
        piece,
        level: currentLevel,
        isRtl: currentLevel % 2 === 1,
      });
    }
  }

  // L2: Reverse contiguous runs at or above each level from maxLevel down to 1.
  const maxLevel = Math.max(0, ...segments.map((seg) => seg.level));
  for (let lvl = maxLevel; lvl >= 1; lvl -= 1) {
    let i = 0;
    while (i < segments.length) {
      if (segments[i]!.level < lvl) {
        i += 1;
        continue;
      }
      let j = i;
      while (j < segments.length && segments[j]!.level >= lvl) {
        j += 1;
      }
      let left = i;
      let right = j - 1;
      while (left < right) {
        const tmp = segments[left]!;
        segments[left] = segments[right]!;
        segments[right] = tmp;
        left += 1;
        right -= 1;
      }
      i = j;
    }
  }
  return segments;
}

type ShapedOutlineGlyph = {
  readonly path: string;
  readonly xAdvance: number;
  readonly xOffset: number;
  readonly yOffset: number;
  readonly cluster: number;
  readonly isSpace: boolean;
};

function shapeSegmentOutlines(
  face: MeasureFace,
  text: string,
  isRtl: boolean,
  lang: string | undefined,
): ShapedOutlineGlyph[] {
  const visible = Array.from(text).filter(requiresFontGlyph).join("");
  if (visible === "") return [];
  const hbFont = hbFaces.get(face.identity)?.font;
  if (hbFont !== undefined) {
    return withHbBuffer((buffer) => {
      buffer.addText(visible);
      buffer.guessSegmentProperties();
      buffer.setDirection(isRtl ? hb.Direction.RTL : hb.Direction.LTR);
      if (lang !== undefined) {
        buffer.setLanguage(lang);
      }
      hb.shape(hbFont, buffer);
      return buffer.getGlyphInfosAndPositions().map((item) => ({
        path: hbFont.glyphToPath(item.codepoint),
        xAdvance: item.xAdvance ?? 0,
        xOffset: item.xOffset ?? 0,
        yOffset: item.yOffset ?? 0,
        cluster: item.cluster,
        isSpace: /\p{White_Space}/u.test(visible[item.cluster] ?? ""),
      }));
    });
  }
  const fkRun = face.font.layout(
    visible,
    undefined,
    undefined,
    undefined,
    isRtl ? "rtl" : "ltr",
  );
  return fkRun.glyphs.map((glyph, index) => {
    const pos = fkRun.positions[index]!;
    const isSpace = glyph.codePoints.some((cp) =>
      /\p{White_Space}/u.test(String.fromCodePoint(cp)),
    );
    return {
      path: glyph.path.toSVG(),
      xAdvance: pos.xAdvance,
      xOffset: pos.xOffset,
      yOffset: pos.yOffset,
      cluster: index,
      isSpace,
    };
  });
}

/**
 * Converts font-measurement/outline unavailability errors encountered during paint()
 * into `preview-unsupported-font` while preserving `resolve()` error semantics.
 */
export function withPreviewFontBoundary<T>(fn: () => T): T {
  try {
    return fn();
  } catch (error) {
    if (error instanceof CapabilityError) {
      if (
        error.code === "preview-unsupported-font" ||
        error.code === "preview-missing-glyph"
      ) {
        throw error;
      }
      if (error.code === "font-measurement-unavailable") {
        throw new CapabilityError(
          "preview-unsupported-font",
          "Preview requires an explicit monochrome outline font file for each used font weight and style",
        );
      }
      if (
        error.code === "font-format-unsupported" ||
        error.code === "font-substitution-required" ||
        error.code === "font-face-mismatch" ||
        error.code === "font-metrics-invalid" ||
        error.code === "conflicting-resource"
      ) {
        throw new CapabilityError("preview-unsupported-font", error.message);
      }
    }
    throw error;
  }
}

/**
 * Extracts vector glyph outlines for one paragraph using its formal measurement / line-break ranges.
 * Coordinates are relative to each line's baseline, in pixels.
 */
export function previewParagraphOutlines(
  paragraph: Paragraph,
  widthOrMeasured: number | MeasuredParagraph,
): PreviewParagraphOutlines {
  const { measured, text, pieces, baseDir } = withPreviewFontBoundary(() => {
    const resolvedMeasured =
      typeof widthOrMeasured === "number"
        ? paragraphMeasurement(paragraph, widthOrMeasured)
        : widthOrMeasured;
    for (const piece of resolvedMeasured.pieces) {
      if (hasUnsupportedColorFontTable(piece.face.font)) {
        throw new CapabilityError(
          "preview-unsupported-font",
          "Preview requires monochrome outline fonts",
        );
      }
    }
    return {
      measured: resolvedMeasured,
      text: resolvedMeasured.text,
      pieces: resolvedMeasured.pieces,
      baseDir: paragraphBaseDirection(resolvedMeasured.text),
    };
  });
  const lines = measured.lines.map((line, lineIndex) => {
    let cursor = 0;
    let pendingTracking = 0;
    const glyphs: PreviewGlyph[] = [];
    const decorations: PreviewDecoration[] = [];
    let ascent = 0;
    let descent = 0;
    const trimmedEnd = trimLineEnd(text, line.start, line.end);
    const visualSegments = buildVisualLineSegments(
      text,
      line.start,
      trimmedEnd,
      pieces,
      baseDir,
    );
    visualSegments.forEach((seg, segIndex) => {
      const { piece } = seg;
      const missing = missingCharacters(piece.face, seg.text);
      if (missing.length > 0) {
        throw new CapabilityError(
          "preview-missing-glyph",
          "Preview requires every visible glyph in the supplied font",
        );
      }
      const shaped = shapeSegmentOutlines(
        piece.face,
        seg.text,
        seg.isRtl,
        piece.style.lang,
      );
      if (shaped.length === 0) return;
      const face = piece.face.font;
      const fontPx = (fontSizePoints(piece.style.fontSize) * 4) / 3;
      const scale = fontPx / face.unitsPerEm;
      const trackingPx = piece.style.letterSpacing ?? 0;
      ascent = Math.max(ascent, face.ascent * scale);
      descent = Math.max(descent, -face.descent * scale);
      let segStartX: number | undefined;
      let segEndX = cursor;
      const glyphStart = glyphs.length;
      let lastCluster: number | undefined;
      let clusterHadAdvance = false;
      for (const item of shaped) {
        if (lastCluster !== undefined && item.cluster !== lastCluster) {
          if (clusterHadAdvance && pendingTracking !== 0) {
            cursor += pendingTracking;
            pendingTracking = 0;
          }
          clusterHadAdvance = false;
        } else if (lastCluster === undefined && pendingTracking !== 0) {
          cursor += pendingTracking;
          pendingTracking = 0;
        }
        lastCluster = item.cluster;
        if (segStartX === undefined) {
          segStartX = cursor;
        }
        glyphs.push({
          path: item.path,
          x: cursor + item.xOffset * scale,
          y: -item.yOffset * scale,
          scale,
          color: piece.style.color,
          ...(piece.style.textShadow === undefined
            ? {}
            : { textShadow: piece.style.textShadow }),
          isSpace: item.isSpace && item.xAdvance > 0,
          clusterKey: `${segIndex}:${item.cluster}`,
        });
        cursor += item.xAdvance * scale;
        segEndX = cursor;
        if (item.xAdvance > 0) {
          clusterHadAdvance = true;
          pendingTracking = trackingPx;
        }
      }
      const glyphEnd = glyphs.length;
      if (segStartX !== undefined && glyphEnd > glyphStart) {
        const decoration = decorationForSegment(
          resolveDecoration(piece.style.textDecoration),
          face,
          fontPx,
          segStartX,
          segEndX,
          piece.style.color,
          piece.style.textShadow,
        );
        if (decoration !== undefined) {
          decorations.push({
            ...decoration,
            glyphStart,
            glyphEnd,
          });
        }
      }
    });
    if (glyphs.length === 0 && pieces.length > 0) {
      const fallbackPiece = pieces[0]!;
      const fallbackScale =
        (fontSizePoints(fallbackPiece.style.fontSize) * 4) /
        (3 * fallbackPiece.face.font.unitsPerEm);
      ascent = fallbackPiece.face.font.ascent * fallbackScale;
      descent = -fallbackPiece.face.font.descent * fallbackScale;
    }
    const endsWithMandatoryBreak =
      trimMandatoryBreakEnd(text, line.start, line.end) < line.end ||
      (line.end < text.length &&
        /[\n\r\v\f\u0085\u2028\u2029]/u.test(text[line.end] ?? ""));
    const canJustify =
      lineIndex < measured.lines.length - 1 && !endsWithMandatoryBreak;
    return {
      ...line,
      glyphs,
      ...(decorations.length === 0 ? {} : { decorations }),
      ascent,
      descent,
      advance: cursor,
      canJustify,
    };
  });
  return {
    lines,
    lineHeight: measured.lineHeight,
    spacingAfter: measured.spacingAfter,
    baseDir,
  };
}

function applyJustifyLine(
  line: PreviewLineOutlines,
  availableWidth: number,
): {
  readonly glyphs: readonly PreviewGlyph[];
  readonly decorations: readonly PreviewDecoration[] | undefined;
  readonly advance: number;
} {
  const slack = availableWidth - line.advance;
  if (!line.canJustify || slack <= 0.01 || line.glyphs.length <= 1) {
    return {
      glyphs: line.glyphs,
      decorations: line.decorations,
      advance: line.advance,
    };
  }
  const count = line.glyphs.length;
  const deltaBefore = Array.from<number>({ length: count }).fill(0);
  const spaceCount = line.glyphs.filter((g) => g.isSpace === true).length;
  if (spaceCount > 0) {
    const extraPerSpace = slack / spaceCount;
    let cumulative = 0;
    for (let i = 0; i < count; i += 1) {
      deltaBefore[i] = cumulative;
      if (line.glyphs[i]!.isSpace === true) {
        cumulative += extraPerSpace;
      }
    }
  } else {
    let clusterTransitions = 0;
    for (let i = 1; i < count; i += 1) {
      if (line.glyphs[i]!.clusterKey !== line.glyphs[i - 1]!.clusterKey) {
        clusterTransitions += 1;
      }
    }
    if (clusterTransitions <= 0) {
      return {
        glyphs: line.glyphs,
        decorations: line.decorations,
        advance: line.advance,
      };
    }
    const extraPerCluster = slack / clusterTransitions;
    let cumulative = 0;
    for (let i = 1; i < count; i += 1) {
      if (line.glyphs[i]!.clusterKey !== line.glyphs[i - 1]!.clusterKey) {
        cumulative += extraPerCluster;
      }
      deltaBefore[i] = cumulative;
    }
  }
  const glyphs = line.glyphs.map((glyph, index) => ({
    ...glyph,
    x: glyph.x + deltaBefore[index]!,
  }));
  const decorations = line.decorations?.map((deco) => {
    const gStart = deco.glyphStart ?? 0;
    const gEnd = deco.glyphEnd ?? count;
    const startDelta = deltaBefore[gStart] ?? 0;
    const endDelta =
      gEnd < count
        ? (deltaBefore[gEnd] ?? 0)
        : (deltaBefore[Math.max(gStart, gEnd - 1)] ?? 0);
    return {
      ...deco,
      x: deco.x + startDelta,
      width: deco.width + (endDelta - startDelta),
    };
  });
  return {
    glyphs,
    decorations,
    advance: availableWidth,
  };
}

function shapeListPrefix(
  paragraph: Paragraph,
  prefixText: string,
): { readonly glyphs: readonly PreviewGlyph[]; readonly advance: number } {
  const prefixStyle = paragraph.runs[0]?.style ?? paragraph.endStyle;
  const prefixFace = withPreviewFontBoundary(() => {
    const face = measureRunFace(prefixStyle);
    if (hasUnsupportedColorFontTable(face.font)) {
      throw new CapabilityError(
        "preview-unsupported-font",
        "Preview requires monochrome outline fonts",
      );
    }
    return face;
  });
  if (missingCharacters(prefixFace, prefixText).length > 0) {
    throw new CapabilityError(
      "preview-missing-glyph",
      "Preview requires every visible glyph in the supplied font",
    );
  }
  const shaped = shapeSegmentOutlines(
    prefixFace,
    prefixText,
    false,
    prefixStyle.lang,
  );
  const scale =
    (fontSizePoints(prefixStyle.fontSize) * 4) /
    (3 * prefixFace.font.unitsPerEm);
  let cursor = 0;
  const glyphs: PreviewGlyph[] = shaped.map((item) => {
    const glyph: PreviewGlyph = {
      path: item.path,
      x: cursor + item.xOffset * scale,
      y: -item.yOffset * scale,
      scale,
      color: prefixStyle.color,
      ...(prefixStyle.textShadow === undefined
        ? {}
        : { textShadow: prefixStyle.textShadow }),
    };
    cursor += item.xAdvance * scale;
    return glyph;
  });
  return { glyphs, advance: cursor };
}

/**
 * Positions paragraph glyph outlines inside a compiled box with insets, horizontal alignment,
 * and vertical alignment. Reuses pre-computed formal build line measurements when provided.
 */
export function previewTextBoxOutlines(options: {
  readonly paragraphs: readonly Paragraph[];
  readonly box: GeometryBox;
  readonly insets?: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
  readonly verticalAlign?: TextStyle["verticalAlign"];
  readonly measuredLayout?: MeasuredTextLayout | undefined;
}): readonly PreviewTextOutlinesLine[] {
  const {
    paragraphs,
    box,
    insets = { top: 0, right: 0, bottom: 0, left: 0 },
    verticalAlign: rawVerticalAlign,
    measuredLayout,
  } = options;
  const verticalAlign = resolveVerticalAlign(rawVerticalAlign);
  const innerWidth = Math.max(0, box.width - insets.left - insets.right);
  const useBoundMeasurement =
    measuredLayout !== undefined &&
    measuredLayout.paragraphs.length === paragraphs.length;
  const outlinedParagraphs = paragraphs.map((paragraph, index) => ({
    paragraph,
    layout: previewParagraphOutlines(
      paragraph,
      useBoundMeasurement ? measuredLayout.paragraphs[index]! : innerWidth,
    ),
  }));
  const totalHeight = outlinedParagraphs.reduce(
    (total, { layout }, index) =>
      total +
      layout.lines.length * layout.lineHeight +
      (index < outlinedParagraphs.length - 1 ? layout.spacingAfter : 0),
    0,
  );
  const freeHeight = box.height - insets.top - insets.bottom - totalHeight;
  let top =
    box.y +
    insets.top +
    (verticalAlign === "middle"
      ? freeHeight / 2
      : verticalAlign === "bottom"
        ? freeHeight
        : 0);
  const activeNumberSeq = new Map<
    number,
    { readonly startAt: number; current: number }
  >();
  return outlinedParagraphs.flatMap(({ paragraph, layout }) => {
    const resolvedList = resolveListModel(paragraph.style.list);
    let prefixText: string | undefined;
    if (resolvedList === undefined) {
      activeNumberSeq.clear();
    } else {
      switch (resolvedList.type) {
        case "bullet": {
          for (const key of activeNumberSeq.keys()) {
            if (key >= resolvedList.level) activeNumberSeq.delete(key);
          }
          prefixText = resolvedList.character;
          break;
        }
        case "number": {
          for (const key of activeNumberSeq.keys()) {
            if (key > resolvedList.level) activeNumberSeq.delete(key);
          }
          const existing = activeNumberSeq.get(resolvedList.level);
          if (
            existing !== undefined &&
            existing.startAt === resolvedList.startAt
          ) {
            existing.current += 1;
          } else {
            activeNumberSeq.set(resolvedList.level, {
              startAt: resolvedList.startAt,
              current: resolvedList.startAt,
            });
          }
          prefixText = `${activeNumberSeq.get(resolvedList.level)!.current}.`;
          break;
        }
        default: {
          const exhaustive: never = resolvedList;
          return exhaustive;
        }
      }
    }
    const listPrefix =
      prefixText === undefined
        ? undefined
        : shapeListPrefix(paragraph, prefixText);
    const listIndent = listIndentMetrics(paragraph);
    const availableLineWidth = Math.max(
      1,
      innerWidth - listIndent.contentIndent,
    );
    const align = resolveAlign(
      paragraph.style.align,
      layout.baseDir === "rtl" ? "right" : "left",
    );
    const lines = layout.lines.map((line, lineIndex) => {
      const justified =
        align === "justify"
          ? applyJustifyLine(line, availableLineWidth)
          : {
              glyphs: line.glyphs,
              decorations: line.decorations,
              advance: line.advance,
            };
      let horizontalOffset: number;
      let prefixOffset = 0;
      if (layout.baseDir === "rtl" && resolvedList !== undefined) {
        switch (align) {
          case "center":
            horizontalOffset = (availableLineWidth - justified.advance) / 2;
            prefixOffset =
              horizontalOffset +
              justified.advance +
              0.4 * listIndent.baseFontPx;
            break;
          case "left":
          case "justify":
            horizontalOffset = 0;
            prefixOffset =
              innerWidth - listIndent.levelIndent - (listPrefix?.advance ?? 0);
            break;
          case "right":
            horizontalOffset = availableLineWidth - justified.advance;
            prefixOffset =
              innerWidth - listIndent.levelIndent - (listPrefix?.advance ?? 0);
            break;
          default: {
            const exhaustive: never = align;
            return exhaustive;
          }
        }
      } else {
        switch (align) {
          case "center":
            horizontalOffset =
              listIndent.contentIndent +
              (availableLineWidth - justified.advance) / 2;
            prefixOffset =
              horizontalOffset -
              (listPrefix?.advance ?? 0) -
              0.4 * listIndent.baseFontPx;
            break;
          case "right":
            horizontalOffset =
              listIndent.contentIndent +
              (availableLineWidth - justified.advance);
            prefixOffset =
              horizontalOffset -
              (listPrefix?.advance ?? 0) -
              0.4 * listIndent.baseFontPx;
            break;
          case "left":
          case "justify":
            horizontalOffset = listIndent.contentIndent;
            prefixOffset = listIndent.levelIndent;
            break;
          default: {
            const exhaustive: never = align;
            return exhaustive;
          }
        }
      }
      const left = box.x + insets.left + horizontalOffset;
      const baseline =
        top +
        line.ascent +
        (layout.lineHeight - line.ascent - line.descent) / 2;
      top += layout.lineHeight;
      const prefixGlyphs =
        lineIndex === 0 && listPrefix !== undefined
          ? listPrefix.glyphs.map((glyph) => ({
              ...glyph,
              x: box.x + insets.left + prefixOffset + glyph.x,
              y: baseline + glyph.y,
            }))
          : [];
      return {
        start: line.start,
        end: line.end,
        baseline,
        glyphs: [
          ...prefixGlyphs,
          ...justified.glyphs.map((glyph) => ({
            ...glyph,
            x: left + glyph.x,
            y: baseline + glyph.y,
          })),
        ],
        ...(justified.decorations === undefined
          ? {}
          : {
              decorations: justified.decorations.map((deco) => ({
                ...deco,
                x: left + deco.x,
                y: baseline + deco.y,
              })),
            }),
      };
    });
    top += layout.spacingAfter;
    return lines;
  });
}

/** Serializes a single positioned glyph outline to an SVG path element. */
export function renderGlyphPathSvg(
  glyph: PreviewGlyph,
  theme: ThemeContext,
): string {
  if (glyph.path === "") return "";
  const color = isResolvedColor(glyph.color)
    ? glyph.color
    : theme.resolveColor(glyph.color ?? "tx1");
  return `<path d="${xmlEscape(glyph.path)}" fill="#${color.hex}" fill-opacity="${color.alpha ?? 1}" transform="translate(${glyph.x} ${glyph.y}) scale(${glyph.scale} ${-glyph.scale})"/>`;
}

function renderDecorationRectSvg(
  deco: PreviewDecoration,
  theme: ThemeContext,
): string {
  const color = isResolvedColor(deco.color)
    ? deco.color
    : theme.resolveColor(deco.color ?? "tx1");
  return `<rect x="${roundSvgNumber(deco.x)}" y="${roundSvgNumber(deco.y)}" width="${roundSvgNumber(deco.width)}" height="${roundSvgNumber(deco.height)}" fill="#${color.hex}" fill-opacity="${color.alpha ?? 1}"/>`;
}

/** Renders box-level positioned glyph outlines into SVG path fragments without importing textCapability. */
export function renderTextOutlinesSvg(options: {
  readonly paragraphs: readonly Paragraph[];
  readonly box: GeometryBox;
  readonly theme: ThemeContext;
  readonly insets?: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
  readonly verticalAlign?: TextStyle["verticalAlign"];
  readonly measuredLayout?: MeasuredTextLayout | undefined;
  readonly idPrefix?: string;
  readonly rotate?: number;
}): string {
  const { theme, idPrefix = "txt", rotate = 0, ...layoutOptions } = options;
  const lines = previewTextBoxOutlines(layoutOptions);
  const defs: string[] = [];
  const shadowFilterIds = new Map<string, string>();
  let shadowIndex = 0;

  const filterIdForShadow = (
    rawShadow: Shadow | "none" | undefined,
  ): string | undefined => {
    const resolved = resolveTextShadowModel(theme, rawShadow);
    switch (resolved.kind) {
      case "none":
        return undefined;
      case "shadow": {
        const { shadow } = resolved;
        const rad = (shadow.angle * Math.PI) / 180;
        const worldDx = shadow.distance * Math.cos(rad);
        const worldDy = shadow.distance * Math.sin(rad);
        const rotRad = (-rotate * Math.PI) / 180;
        const localDx = worldDx * Math.cos(rotRad) - worldDy * Math.sin(rotRad);
        const localDy = worldDx * Math.sin(rotRad) + worldDy * Math.cos(rotRad);
        const sigma = shadow.blur / 2;
        const alpha = shadow.color.alpha ?? 1;
        const key = `${shadow.color.hex}:${roundSvgNumber(alpha)}:${roundSvgNumber(sigma)}:${roundSvgNumber(localDx)}:${roundSvgNumber(localDy)}`;
        const existing = shadowFilterIds.get(key);
        if (existing !== undefined) return existing;
        const id = `${idPrefix}-tshdw-${shadowIndex++}`;
        shadowFilterIds.set(key, id);
        const pad =
          shadow.blur * 3 + Math.abs(localDx) + Math.abs(localDy) + 64;
        const fx = roundSvgNumber(options.box.x - pad);
        const fy = roundSvgNumber(options.box.y - pad);
        const fw = roundSvgNumber(options.box.width + pad * 2);
        const fh = roundSvgNumber(options.box.height + pad * 2);
        defs.push(
          `<filter id="${id}" filterUnits="userSpaceOnUse" x="${fx}" y="${fy}" width="${fw}" height="${fh}"><feDropShadow dx="${roundSvgNumber(localDx)}" dy="${roundSvgNumber(localDy)}" stdDeviation="${roundSvgNumber(sigma)}" flood-color="#${shadow.color.hex}" flood-opacity="${roundSvgNumber(alpha)}"/></filter>`,
        );
        return id;
      }
      default: {
        const exhaustive: never = resolved;
        return exhaustive;
      }
    }
  };

  const parts: string[] = [];
  let activeFilter: string | undefined;
  let activeGroup: string[] = [];
  const flushGroup = (): void => {
    if (activeGroup.length === 0) return;
    const joined = activeGroup.join("");
    parts.push(
      activeFilter === undefined
        ? joined
        : `<g filter="url(#${activeFilter})">${joined}</g>`,
    );
    activeGroup = [];
  };
  const pushElement = (
    elementSvg: string,
    rawShadow: Shadow | "none" | undefined,
  ): void => {
    if (elementSvg === "") return;
    const filterId = filterIdForShadow(rawShadow);
    if (filterId !== activeFilter) {
      flushGroup();
      activeFilter = filterId;
    }
    activeGroup.push(elementSvg);
  };

  for (const line of lines) {
    for (const glyph of line.glyphs) {
      pushElement(renderGlyphPathSvg(glyph, theme), glyph.textShadow);
    }
    for (const deco of line.decorations ?? []) {
      pushElement(renderDecorationRectSvg(deco, theme), deco.textShadow);
    }
  }
  flushGroup();

  const defsXml = defs.length > 0 ? `<defs>${defs.join("")}</defs>` : "";
  return `${defsXml}${parts.join("")}`;
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

function bodyAnchor(verticalAlign: TextStyle["verticalAlign"]): string {
  return { top: "t", middle: "ctr", bottom: "b" }[verticalAlign ?? "top"];
}

export function textInsets(
  value: unknown,
  placeholder?: PlaceholderResolutionContext,
): {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
} {
  if (placeholder === undefined) return parseTextInsets(value);
  const effective: Record<string, unknown> = {};
  const layers = [
    placeholder.masterPlaceholder?.style,
    placeholder.layoutPlaceholder?.style,
    record(value) ? value : undefined,
  ];
  for (const layer of layers) {
    if (layer === undefined) continue;
    for (const key of paddingStyleKeys) {
      if (hasLocalDeclaration(layer, key)) {
        if (key === "padding") {
          delete effective["paddingTop"];
          delete effective["paddingRight"];
          delete effective["paddingBottom"];
          delete effective["paddingLeft"];
        }
        effective[key] = layer[key];
      }
    }
  }
  return parseTextInsets(effective);
}

/** Parses text children, validates element structure, and resolves paragraphs and text styles once. */
export function parseTextContent(
  children: readonly unknown[],
  styleValue: unknown,
  theme: ThemeContext,
  fields: TextFieldContext = { slideNumber: 1 },
  label = "Text.style",
  placeholder?: PlaceholderResolutionContext,
  fullSlideStyle?: unknown,
): {
  readonly paragraphs: readonly Paragraph[];
  readonly style: ResolvedTextStyle;
  readonly localBodyDeclared?: ReadonlySet<string>;
} {
  if (placeholder === undefined) {
    const style = readTextStyle(styleValue, label, theme);
    const paragraphs = normalizeParagraphs(children, style, theme, fields);
    return { paragraphs, style };
  }
  const { style, localDeclaredStyle } = readPlaceholderTextStyle(
    styleValue,
    label,
    placeholder,
  );
  const fullRecord = record(fullSlideStyle) ? fullSlideStyle : undefined;
  const localBodyDeclared = localDeclaredKeySet(fullRecord, [
    "textWrap",
    "verticalAlign",
    ...paddingStyleKeys,
  ]);
  const paragraphs = normalizeParagraphs(
    children,
    style,
    placeholder.masterTheme,
    fields,
    placeholder,
    localDeclaredStyle,
  );
  return { paragraphs, style, localBodyDeclared };
}

export const parseTextParagraphs = parseTextContent;

function validateEmittableParagraphs(
  paragraphs: readonly Paragraph[],
  theme: ThemeContext,
): void {
  for (const paragraph of paragraphs) {
    if (paragraph.style.list?.type === "bullet") {
      escapedXml(paragraph.style.list.character ?? "•");
    }
    if (paragraph.runs.length === 0) {
      if (paragraph.endStyle.lang !== undefined) {
        escapedXml(paragraph.endStyle.lang);
      }
      centipoints(fontSizePoints(paragraph.endStyle.fontSize));
      continue;
    }
    for (const run of paragraph.runs) {
      escapedXml(run.text);
      for (const font of normalizedFonts(run.style)) {
        escapedXml(font.family);
      }
      if (run.style.lang !== undefined) {
        escapedXml(run.style.lang);
      }
      if (
        run.style.textShadow !== undefined &&
        run.style.textShadow !== "none"
      ) {
        textShadowXml(theme, run.style.textShadow);
      }
    }
  }
}

const TEXT_FIT_EPSILON = 0.01;

/** Computes tri-state TextMetrics from a resolved text body and its finalized available content dimensions. */
export function computeTextMetrics(
  resolved: Pick<ResolvedTextBody, "measuredLayout" | "unmeasuredIssues">,
  availableWidth: number,
  availableHeight: number,
): TextMetrics {
  const { missingResourceFonts, missingGlyphFonts, referencedFonts } =
    resolved.unmeasuredIssues;
  if (missingResourceFonts.length > 0) {
    return {
      status: "unmeasured",
      reason: "missing-font-resource",
      fonts: missingResourceFonts,
    };
  }
  if (missingGlyphFonts.length > 0) {
    return {
      status: "unmeasured",
      reason: "missing-glyphs",
      fonts: missingGlyphFonts,
    };
  }
  if (resolved.measuredLayout === undefined) {
    return {
      status: "unmeasured",
      reason: "missing-font-resource",
      fonts: referencedFonts,
    };
  }
  const available = {
    width: Math.max(0, availableWidth),
    height: Math.max(0, availableHeight),
  };
  const content = {
    width: resolved.measuredLayout.width,
    height: resolved.measuredLayout.height,
    lineCount: resolved.measuredLayout.paragraphs.reduce(
      (sum, paragraph) => sum + paragraph.lineCount,
      0,
    ),
  };
  const fits =
    content.width <= available.width + TEXT_FIT_EPSILON &&
    content.height <= available.height + TEXT_FIT_EPSILON;
  return {
    status: "measured",
    available,
    content,
    fits,
  };
}

/**
 * Resolves text layout, wrapping metrics, overset diagnostics, plain-text observations,
 * and font resource claims from pre-parsed paragraphs without assigning emit-only field IDs.
 */
export function resolveTextBody(
  paragraphs: readonly Paragraph[],
  style: ResolvedTextStyle | TextStyle,
  theme: ThemeContext,
  verticalAlign: TextStyle["verticalAlign"] = "top",
  layout?: TextLayoutRequest,
  destination: TextBodyDestination = "text-box",
  placeholderOptions?: {
    readonly isPlaceholder?: boolean;
    readonly localBodyDeclared?: ReadonlySet<string>;
  },
): ResolvedTextBody {
  const runs = paragraphs.flatMap((paragraph) =>
    paragraph.runs.length > 0
      ? paragraph.runs
      : [{ text: "", style: paragraph.endStyle }],
  );
  const listRuns = paragraphs.flatMap((paragraph) => {
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
  const {
    resourceClaims,
    fontUses,
    fontUsages,
    warnings,
    missingGlyphErrors,
    unmeasuredIssues,
  } = textResourceClaims([...runs, ...listRuns]);
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
      ? paragraphs.map((paragraph, index) => ({
          ...paragraph,
          style: {
            ...paragraph.style,
            lineHeight: measured.paragraphs[index]!.lineHeight,
          },
        }))
      : paragraphs;
  validateEmittableParagraphs(outputParagraphs, theme);
  const text = outputParagraphs
    .map(({ runs: paragraphRuns }) =>
      paragraphRuns.map(({ text: part }) => part).join(""),
    )
    .join("\n");
  const textMetrics =
    layout === undefined
      ? undefined
      : computeTextMetrics(
          {
            ...(measured === undefined ? {} : { measuredLayout: measured }),
            unmeasuredIssues,
          },
          layout.width === "max-content"
            ? (measured?.width ?? 0)
            : layout.width,
          layout.height ?? measured?.height ?? 0,
        );
  return {
    paragraphs: outputParagraphs,
    style,
    verticalAlign,
    ...(layout === undefined ? {} : { layout }),
    destination,
    ...(placeholderOptions?.isPlaceholder === true
      ? { isPlaceholder: true }
      : {}),
    ...(placeholderOptions?.localBodyDeclared === undefined
      ? {}
      : { localBodyDeclared: placeholderOptions.localBodyDeclared }),
    text,
    paragraphCount: outputParagraphs.length,
    resourceClaims,
    fontUses,
    fontUsages,
    ...(measured === undefined
      ? {}
      : {
          measurement: { width: measured.width, height: measured.height },
          measuredLayout: measured,
        }),
    warnings,
    missingGlyphErrors,
    unmeasuredIssues,
    ...(textMetrics === undefined ? {} : { textMetrics }),
  };
}

/**
 * Emits PresentationML txBody XML and font embedding contributions from resolved text facts.
 * Preserves native paragraph/run structure without converting soft wraps into hard line breaks.
 */
export function emitTextBody(
  resolved: ResolvedTextBody,
  theme: ThemeContext,
  options: {
    readonly createFieldId?: () => string;
    readonly embedFonts?: FontEmbeddingMode;
  } = {},
): EmittedTextBody {
  const { createFieldId, embedFonts } = options;
  if (
    (embedFonts === "full" || embedFonts === "subset") &&
    resolved.missingGlyphErrors.length > 0
  ) {
    throw new CapabilityError(
      "font-missing-glyph",
      resolved.missingGlyphErrors[0]!,
    );
  }
  const emittedParagraphs =
    createFieldId !== undefined
      ? resolved.paragraphs.map((paragraph) => ({
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
      : resolved.paragraphs;
  const { style, verticalAlign, layout, destination, isPlaceholder } = resolved;
  let xml: string;
  if (destination === "table-cell") {
    xml = `<a:txBody><a:bodyPr/><a:lstStyle/>${emittedParagraphs.map((paragraph) => paragraphXml(paragraph, theme)).join("")}</a:txBody>`;
  } else if (isPlaceholder) {
    const localBody = resolved.localBodyDeclared ?? new Set<string>();
    const bodyAttrs: string[] = [];
    if (localBody.has("textWrap")) {
      bodyAttrs.push(
        `wrap="${style.textWrap === "nowrap" ? "none" : "square"}"`,
      );
    }
    if (
      paddingStyleKeys.some((key) => localBody.has(key)) &&
      layout?.insets !== undefined
    ) {
      bodyAttrs.push(
        `lIns="${pxToEmu(layout.insets.left)}"`,
        `tIns="${pxToEmu(layout.insets.top)}"`,
        `rIns="${pxToEmu(layout.insets.right)}"`,
        `bIns="${pxToEmu(layout.insets.bottom)}"`,
      );
    }
    if (localBody.has("verticalAlign")) {
      bodyAttrs.push(
        `anchor="${bodyAnchor(style.verticalAlign ?? verticalAlign)}"`,
      );
    }
    const bodyPr =
      bodyAttrs.length === 0
        ? "<a:bodyPr/>"
        : `<a:bodyPr ${bodyAttrs.join(" ")}/>`;
    xml = `<p:txBody>${bodyPr}<a:lstStyle/>${emittedParagraphs.map((paragraph) => paragraphXml(paragraph, theme, true)).join("")}</p:txBody>`;
  } else {
    xml = `<p:txBody><a:bodyPr wrap="${style.textWrap === "nowrap" ? "none" : "square"}" lIns="${pxToEmu(layout?.insets?.left ?? 0)}" tIns="${pxToEmu(layout?.insets?.top ?? 0)}" rIns="${pxToEmu(layout?.insets?.right ?? 0)}" bIns="${pxToEmu(layout?.insets?.bottom ?? 0)}" anchor="${bodyAnchor(style.verticalAlign ?? verticalAlign)}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${emittedParagraphs.map((paragraph) => paragraphXml(paragraph, theme)).join("")}</p:txBody>`;
  }
  return {
    xml,
    presentationContributions: fontPresentationContributions(
      resolved.fontUses,
      embedFonts,
    ),
  };
}

type ChartTextResult = {
  readonly xml: string;
  readonly text: string;
  readonly resourceClaims: readonly TextResourceClaim[];
  readonly fontUses: readonly FontEmbeddingUse[];
  readonly fontUsages: readonly CapabilityFontUsageClaim[];
  readonly warnings: readonly {
    readonly code: string;
    readonly message: string;
  }[];
  readonly missingGlyphErrors: readonly string[];
};

/** Generates shared font styles for native chart text and registers font claims based on actual categories, series, and label text. */
export function writeChartTextProperties(
  styleValue: unknown,
  texts: readonly string[],
  label: string,
  theme: ThemeContext,
): ChartTextResult {
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
  const { resourceClaims, fontUses, fontUsages, warnings, missingGlyphErrors } =
    textResourceClaims(runs);
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
    fontUsages,
    warnings,
    missingGlyphErrors,
  };
}

/** Emits a native chart or axis title with actual run properties, not canvas text. */
export function writeChartTitle(
  text: string,
  styleValue: unknown,
  label: string,
  theme: ThemeContext,
): ChartTextResult {
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
  const { resourceClaims, fontUses, fontUsages, warnings, missingGlyphErrors } =
    textResourceClaims([{ text, style: runStyle }]);
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
    fontUsages,
    warnings,
    missingGlyphErrors,
  };
}

/** Converts fixed pixel length into DrawingML EMU. */
export function textExtent(value: number, label: string): number {
  return pxToEmu(parseTextExtent(value, label));
}
