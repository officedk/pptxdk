import { createHash } from "node:crypto";

import { CapabilityError, pxToEmu, xmlEscape } from "@react-pptx/core";

import type { Font, FontFaceStyle } from "./font.ts";

export type FontReference = string | Font;
export type FontSize = number | `${number}px` | `${number}pt`;
export type TextAlignment = "left" | "center" | "right" | "justify";

export type TextStyle = {
  readonly fontFamily?: FontReference;
  readonly fontSize?: FontSize;
  readonly fontWeight?: number | "normal" | "bold";
  readonly fontStyle?: FontFaceStyle;
  readonly color?: string;
  readonly lineHeight?: FontSize;
  readonly verticalAlign?: "top" | "middle" | "bottom";
};

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
};

export type TextServiceResult = {
  readonly xml: string;
  readonly text: string;
  readonly resourceClaims: readonly TextResourceClaim[];
  readonly warnings: readonly {
    readonly code: string;
    readonly message: string;
  }[];
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
  readonly color?: string | undefined;
};
type Run = { readonly text: string; readonly style: RunStyle };
type Paragraph = {
  readonly runs: readonly Run[];
  readonly style: ParagraphStyle;
};
type NormalizedFont = {
  readonly family: string;
  readonly weight: number;
  readonly style: FontFaceStyle;
  readonly hasResource: boolean;
  readonly identity: string;
};

const defaultTextStyle: TextStyle = Object.freeze({
  fontFamily: "Aptos",
  fontSize: 24,
  fontWeight: "normal",
  fontStyle: "normal",
  color: "#000000",
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

function lengthCentipoints(pixels: number): number {
  return Math.round(pixels * 75);
}

function fontSizePoints(value: FontSize | undefined): number {
  if (value === undefined)
    return positiveLength(defaultTextStyle.fontSize, "字号") * 0.75;
  if (typeof value === "string" && value.endsWith("pt")) {
    const points = Number(value.slice(0, -2));
    if (!Number.isFinite(points) || points <= 0) {
      throw new CapabilityError("invalid-font-size", "fontSize 必须大于 0");
    }
    return points;
  }
  return positiveLength(value, "fontSize") * 0.75;
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
  onlyKnownKeys(value, ["family", "faces"], "Font");
  if (value["faces"] !== undefined && !Array.isArray(value["faces"])) {
    throw new CapabilityError("invalid-font-family", "Font.faces 必须是数组");
  }
  return value as Font;
}

function normalizedFont(style: RunStyle): NormalizedFont {
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
  const face = faces.find(
    (candidate) =>
      (candidate.weight ?? 400) === weight &&
      (candidate.style ?? "normal") === faceStyle,
  );
  const data = face?.data;
  const identity =
    data === undefined
      ? "missing"
      : createHash("sha256").update(data).digest("hex");
  return {
    family: family.family,
    weight,
    style: faceStyle,
    hasResource: data !== undefined,
    identity,
  };
}

function validateRunStyle(style: RunStyle): RunStyle {
  fontSizePoints(style.fontSize);
  weightValue(style.fontWeight);
  normalizedFont(style);
  if (style.color !== undefined && !/^#[0-9a-f]{6}$/iu.test(style.color)) {
    throw new CapabilityError(
      "invalid-text-color",
      "color 必须是 #RRGGBB 颜色",
    );
  }
  return style;
}

function readRunStyle(value: unknown): RunStyle {
  const props = styleRecord(value, "Span.style");
  onlyKnownKeys(
    props,
    ["fontFamily", "fontSize", "fontWeight", "fontStyle", "color"],
    "Span.style",
  );
  const style = props as RunStyle;
  validateRunStyle(style);
  return style;
}

function mergeRunStyle(parent: RunStyle, child: RunStyle): RunStyle {
  return validateRunStyle({ ...parent, ...child });
}

function readTextStyle(value: unknown, label: string): TextStyle {
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
    ],
    label,
  );
  const style = props as TextStyle;
  if (
    style.verticalAlign !== undefined &&
    !["top", "middle", "bottom"].includes(style.verticalAlign)
  ) {
    throw new CapabilityError(
      "invalid-vertical-align",
      "verticalAlign 只能是 top、middle 或 bottom",
    );
  }
  const runStyle: RunStyle = {
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    color: style.color,
  };
  validateRunStyle(runStyle);
  if (style.lineHeight !== undefined)
    positiveLength(style.lineHeight, "lineHeight");
  return style;
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
      "list",
    ],
    "P.style",
  );
  const style = props as ParagraphStyle;
  if (
    style.align !== undefined &&
    !["left", "center", "right", "justify"].includes(style.align)
  ) {
    throw new CapabilityError(
      "invalid-text-alignment",
      "align 只能是 left、center、right 或 justify",
    );
  }
  if (style.lineHeight !== undefined)
    positiveLength(style.lineHeight, "lineHeight");
  if (
    style.spacingAfter !== undefined &&
    (!Number.isFinite(style.spacingAfter) || style.spacingAfter < 0)
  ) {
    throw new CapabilityError(
      "invalid-paragraph-spacing",
      "spacingAfter 必须是有限非负像素数值",
    );
  }
  if (style.list !== undefined) validateList(style.list);
  const runStyle: RunStyle = {
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    color: style.color,
  };
  validateRunStyle(runStyle);
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
    const nextStyle = mergeRunStyle(style, spanStyle);
    collectRuns(child.children, nextStyle, runs);
  }
}

function runXml(run: Run): string {
  const font = normalizedFont(run.style);
  const points = fontSizePoints(run.style.fontSize);
  const weight = font.weight;
  const bold = weight >= 600;
  const color = run.style.color ?? "#000000";
  const text = escapedXml(run.text);
  return `<a:r><a:rPr lang="en-US" sz="${centipoints(points)}" b="${bold ? 1 : 0}" i="${font.style === "italic" ? 1 : 0}"><a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill><a:latin typeface="${escapedXml(font.family)}"/></a:rPr><a:t xml:space="preserve">${text}</a:t></a:r>`;
}

function alignmentValue(value: TextAlignment | undefined): string | undefined {
  if (value === undefined) return undefined;
  return { left: "l", center: "ctr", right: "r", justify: "just" }[value];
}

function paragraphXml(paragraph: Paragraph): string {
  const { style, runs } = paragraph;
  const attributes: string[] = [];
  const align = alignmentValue(style.align);
  if (align !== undefined) attributes.push(`algn="${align}"`);
  const level = style.list?.level ?? 0;
  if (style.list !== undefined) attributes.push(`lvl="${level}"`);
  const lineSpacing =
    style.lineHeight === undefined
      ? ""
      : `<a:lnSpc><a:spcPts val="${lengthCentipoints(positiveLength(style.lineHeight, "lineHeight"))}"/></a:lnSpc>`;
  const spacingAfter =
    style.spacingAfter === undefined || style.spacingAfter === 0
      ? ""
      : `<a:spcAft><a:spcPts val="${lengthCentipoints(style.spacingAfter)}"/></a:spcAft>`;
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
      ? `<a:endParaRPr lang="en-US" sz="${centipoints(fontSizePoints(style.fontSize))}"/>`
      : "";
  return `<a:p>${pPr}${runs.map(runXml).join("")}${endProperties}</a:p>`;
}

function normalizeParagraphs(
  children: readonly unknown[],
  textStyle: TextStyle,
): Paragraph[] {
  const baseRunStyle: RunStyle = {
    fontFamily: textStyle.fontFamily,
    fontSize: textStyle.fontSize,
    fontWeight: textStyle.fontWeight,
    fontStyle: textStyle.fontStyle,
    color: textStyle.color,
  };
  validateRunStyle(baseRunStyle);

  const paragraphs: Paragraph[] = [];
  let inline: Run[] = [];
  const baseParagraphStyle: ParagraphStyle =
    textStyle.lineHeight === undefined
      ? {}
      : { lineHeight: textStyle.lineHeight };
  const flushInline = (): void => {
    if (inline.length > 0) {
      paragraphs.push({ runs: inline, style: baseParagraphStyle });
      inline = [];
    }
  };

  for (const child of children) {
    if (!hostChild(child)) {
      throw new CapabilityError("invalid-text-child", "文字内容树包含无效节点");
    }
    if (child.type === "#text" || child.type === "rp:span") {
      collectRuns([child], baseRunStyle, inline);
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
    const paragraphRunStyle = mergeRunStyle(baseRunStyle, {
      fontFamily: style.fontFamily,
      fontSize: style.fontSize,
      fontWeight: style.fontWeight,
      fontStyle: style.fontStyle,
      color: style.color,
    });
    const runs: Run[] = [];
    collectRuns(child.children, paragraphRunStyle, runs);
    paragraphs.push({ runs, style });
  }
  flushInline();
  if (paragraphs.length === 0) paragraphs.push({ runs: [], style: {} });
  return paragraphs;
}

function bodyAnchor(verticalAlign: TextStyle["verticalAlign"]): string {
  return { top: "t", middle: "ctr", bottom: "b" }[verticalAlign ?? "top"];
}

/** 为文本框与形状文字共用段落、run、列表和字体资源处理。 */
export function writeTextBody(
  children: readonly unknown[],
  styleValue: unknown,
  verticalAlign: TextStyle["verticalAlign"] = "top",
): TextServiceResult {
  const style = readTextStyle(styleValue, "Text.style");
  const paragraphs = normalizeParagraphs(children, style);
  const resolvedRuns = paragraphs.flatMap(({ runs }) => runs);
  const used = new Map<string, NormalizedFont[]>();
  for (const run of resolvedRuns) {
    const font = normalizedFont(run.style);
    const key = `${font.family.toLocaleLowerCase("en-US")}\u0000${font.weight}\u0000${font.style}`;
    const resources = used.get(key) ?? [];
    if (!resources.some(({ identity }) => identity === font.identity)) {
      resources.push(font);
      used.set(key, resources);
    }
  }
  const resourceClaims: TextResourceClaim[] = [];
  const warnings: { code: string; message: string }[] = [];
  for (const [key, resources] of used) {
    for (const font of resources) {
      resourceClaims.push({
        kind: "font-face",
        key,
        identity: font.identity,
      });
      if (!font.hasResource) {
        warnings.push({
          code: "font-measurement-unavailable",
          message: `未提供 ${font.family} ${font.weight} ${font.style} 的显式字体资源；固定区域文字不会核对实际换行或装入情况`,
        });
      }
    }
  }
  const xml = `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${bodyAnchor(style.verticalAlign ?? verticalAlign)}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paragraphs.map(paragraphXml).join("")}</p:txBody>`;
  const text = paragraphs
    .map(({ runs }) => runs.map(({ text: part }) => part).join(""))
    .join("\n");
  return { xml, text, resourceClaims, warnings };
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
