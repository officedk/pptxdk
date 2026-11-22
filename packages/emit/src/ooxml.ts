import { createHash } from "node:crypto";

import type {
  HeaderFooter,
  LayoutPlaceholderSpec,
  MasterPlaceholderSpec,
  MasterTextStyles,
  NativePlaceholderRef,
  PlaceholderLevelStyle,
  PlaceholderListStyles,
} from "@pptxdk/core";
import {
  createThemeContext,
  defaultColorMap,
  defaultTheme,
  isFont,
  nativeColorSlots,
  schemeColorSlots,
  type ColorMap,
  type ColorMapOverride,
  type FontSize,
  type FormatColorTransform,
  type FormatEffectStyle,
  type FormatFillStyle,
  type FormatLineStyle,
  type FormatScheme,
  type NativeColorSlot,
  type ResolvedColor,
  type TextAlignment,
  type TextStyle,
  type Theme,
  type ThemeContext,
  type ThemeOverride,
} from "@pptxdk/core/theme";

import { shapeTreeXml } from "#/shape-tree.ts";

export function pxToEmu(value: number): number {
  return Math.round(value * 9525);
}

export function assertXml10Characters(value: string): void {
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      codePoint !== 0x9 &&
      codePoint !== 0xa &&
      codePoint !== 0xd &&
      (codePoint < 0x20 ||
        (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
        codePoint === 0xfffe ||
        codePoint === 0xffff)
    ) {
      throw new Error("String contains characters not allowed in XML 1.0");
    }
  }
}

export function xmlEscape(value: string): string {
  assertXml10Characters(value);
  // XML normalizes raw whitespace in attributes; character references preserve identity values.
  return value
    .replace(/&/gu, "&amp;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;")
    .replace(/"/gu, "&quot;")
    .replace(/'/gu, "&apos;")
    .replace(/\r/gu, "&#xD;")
    .replace(/\n/gu, "&#xA;")
    .replace(/\t/gu, "&#x9;");
}

export function xmlDocument(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${body}`;
}

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
  const schemeVal =
    value.scheme ??
    (value.slot === undefined ? undefined : nativeSchemeReference[value.slot]);
  const element = schemeVal === undefined ? "srgbClr" : "schemeClr";
  const color = schemeVal ?? value.hex;
  return children === ""
    ? `<a:${element} val="${color}"/>`
    : `<a:${element} val="${color}">${children}</a:${element}>`;
}

const namespaces = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  relationships: "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypes: "http://schemas.openxmlformats.org/package/2006/content-types",
} as const;

const minimumSlideLayoutId = 2_147_483_648;

function relationships(
  entries: readonly {
    id: string;
    type: string;
    target: string;
    targetMode?: "External";
  }[],
): string {
  return xmlDocument(
    `<Relationships xmlns="${namespaces.relationships}">${entries
      .map(
        ({ id, type, target, targetMode }) =>
          `<Relationship Id="${xmlEscape(id)}" Type="${xmlEscape(type)}" Target="${xmlEscape(target)}"${targetMode === undefined ? "" : ` TargetMode="${xmlEscape(targetMode)}"`}/>`,
      )
      .join("")}</Relationships>`,
  );
}

export type PptxPackageResource = {
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly extension: string;
  readonly contentType: string;
};

export type PptxRelationship = {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: "External";
};

export type PptxPackagePart = {
  readonly path: string;
  readonly content: string | Uint8Array;
  readonly contentType: string;
};

export type PptxPackageRelationshipSet = {
  readonly sourcePath: string;
  readonly relationships: readonly PptxRelationship[];
};

export type EmittedSlideLayoutSpec = {
  readonly key: string;
  readonly name?: string;
  readonly type?: string;
  readonly preserve?: boolean;
  readonly matchingName?: string;
  readonly userDrawn?: boolean;
  readonly showMasterShapes?: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly placeholders?: Readonly<Record<string, LayoutPlaceholderSpec>>;
  readonly colorMapOverride?: ColorMapOverride;
  readonly themeOverride?: ThemeOverride;
  readonly headerFooter?: HeaderFooter;
  readonly backgroundXml?: string;
  readonly shapeXml?: string;
  readonly extraRelationships?: readonly PptxRelationship[];
  readonly transitionXml?: string;
};

export type EmittedSlideMasterSpec = {
  readonly key: string;
  readonly name?: string;
  readonly preserve?: boolean;
  readonly theme?: Theme;
  readonly themeContext?: ThemeContext;
  readonly colorMap?: ColorMap;
  readonly headerFooter?: HeaderFooter;
  readonly backgroundXml?: string;
  readonly shapeXml?: string;
  readonly extraRelationships?: readonly PptxRelationship[];
  readonly placeholders?: Readonly<Record<string, MasterPlaceholderSpec>>;
  readonly textStyles?: MasterTextStyles;
  readonly transitionXml?: string;
  readonly layouts: readonly EmittedSlideLayoutSpec[];
};

const defaultEmittedMasters: readonly EmittedSlideMasterSpec[] = [
  {
    key: "master-1",
    name: "Office Theme",
    layouts: [
      {
        key: "layout-1",
        name: "Blank",
        type: "blank",
        preserve: true,
      },
    ],
  },
];

function contentTypes(
  masterCount: number,
  layoutCount: number,
  themeCount: number,
  themeOverrideCount: number,
  slideCount: number,
  resources: readonly PptxPackageResource[],
  packageParts: readonly PptxPackagePart[],
  hasCoreProperties: boolean,
): string {
  const overrides = [
    [
      "/ppt/presentation.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    ],
    ...Array.from({ length: masterCount }, (_, index) => [
      `/ppt/slideMasters/slideMaster${index + 1}.xml`,
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    ]),
    ...Array.from({ length: layoutCount }, (_, index) => [
      `/ppt/slideLayouts/slideLayout${index + 1}.xml`,
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    ]),
    ...Array.from({ length: themeCount }, (_, index) => [
      `/ppt/theme/theme${index + 1}.xml`,
      "application/vnd.openxmlformats-officedocument.theme+xml",
    ]),
    ...Array.from({ length: themeOverrideCount }, (_, index) => [
      `/ppt/theme/themeOverride${index + 1}.xml`,
      "application/vnd.openxmlformats-officedocument.themeOverride+xml",
    ]),
    ...Array.from({ length: slideCount }, (_, index) => [
      `/ppt/slides/slide${index + 1}.xml`,
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    ]),
    ...(hasCoreProperties
      ? [
          [
            "/docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
          ],
        ]
      : []),
    ...packageParts.map((part) => [`/${part.path}`, part.contentType]),
  ];
  const defaults = new Map<string, string>([
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"],
  ]);
  for (const resource of resources) {
    defaults.set(resource.extension, resource.contentType);
  }
  return xmlDocument(
    `<Types xmlns="${namespaces.contentTypes}">${[...defaults]
      .map(
        ([extension, contentType]) =>
          `<Default Extension="${xmlEscape(extension)}" ContentType="${xmlEscape(contentType)}"/>`,
      )
      .join("")}${overrides
      .map(
        ([partName, contentType]) =>
          `<Override PartName="${xmlEscape(partName!)}" ContentType="${xmlEscape(contentType!)}"/>`,
      )
      .join("")}</Types>`,
  );
}

function normalizedHex6(color: string, label: string): string {
  const trimmed = color.trim();
  const six = /^#[0-9a-f]{3}$/iu.test(trimmed)
    ? trimmed
        .slice(1)
        .split("")
        .map((c) => c + c)
        .join("")
    : /^#[0-9a-f]{6}$/iu.test(trimmed)
      ? trimmed.slice(1)
      : undefined;
  if (six === undefined) {
    throw new Error(`${label} must be a #RRGGBB color`);
  }
  return six.toUpperCase();
}

function colorSchemeXml(
  name: string,
  colors: Readonly<Record<NativeColorSlot, string>>,
): string {
  return `<a:clrScheme name="${xmlEscape(name)}">${nativeColorSlots
    .map((slot) => {
      const hex = normalizedHex6(colors[slot], `Theme.colors.${slot}`);
      return `<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`;
    })
    .join("")}</a:clrScheme>`;
}

function fontSchemeXml(
  name: string,
  majorValue: Theme["fonts"]["major"],
  minorValue: Theme["fonts"]["minor"],
): string {
  const major = isFont(majorValue) ? majorValue.family : majorValue;
  const minor = isFont(minorValue) ? minorValue.family : minorValue;
  const majorEscaped = xmlEscape(major);
  const minorEscaped = xmlEscape(minor);
  return `<a:fontScheme name="${xmlEscape(name)}"><a:majorFont><a:latin typeface="${majorEscaped}"/><a:ea typeface="${majorEscaped}"/><a:cs typeface="${majorEscaped}"/></a:majorFont><a:minorFont><a:latin typeface="${minorEscaped}"/><a:ea typeface="${minorEscaped}"/><a:cs typeface="${minorEscaped}"/></a:minorFont></a:fontScheme>`;
}

function formatColorTransformXml(transform: FormatColorTransform): string {
  const mods = [
    transform.tint === undefined
      ? ""
      : `<a:tint val="${Math.round(transform.tint * 100_000)}"/>`,
    transform.shade === undefined
      ? ""
      : `<a:shade val="${Math.round(transform.shade * 100_000)}"/>`,
    transform.satMod === undefined
      ? ""
      : `<a:satMod val="${Math.round(transform.satMod * 100_000)}"/>`,
    transform.lumMod === undefined
      ? ""
      : `<a:lumMod val="${Math.round(transform.lumMod * 100_000)}"/>`,
    transform.lumOff === undefined
      ? ""
      : `<a:lumOff val="${Math.round(transform.lumOff * 100_000)}"/>`,
    transform.alpha === undefined
      ? ""
      : `<a:alpha val="${Math.round(transform.alpha * 100_000)}"/>`,
  ].join("");
  if (transform.color.startsWith("#")) {
    const hex = normalizedHex6(transform.color, "FormatScheme color");
    return mods === ""
      ? `<a:srgbClr val="${hex}"/>`
      : `<a:srgbClr val="${hex}">${mods}</a:srgbClr>`;
  }
  return mods === ""
    ? `<a:schemeClr val="${xmlEscape(transform.color)}"/>`
    : `<a:schemeClr val="${xmlEscape(transform.color)}">${mods}</a:schemeClr>`;
}

function formatFillStyleXml(fill: FormatFillStyle): string {
  if (fill.kind === "none") return "<a:noFill/>";
  if (fill.kind === "solid") {
    return `<a:solidFill>${formatColorTransformXml(fill.color)}</a:solidFill>`;
  }
  const rotAttr =
    fill.rotWithShape === undefined
      ? ""
      : ` rotWithShape="${fill.rotWithShape ? "1" : "0"}"`;
  const stops = fill.stops
    .map(
      (stop) =>
        `<a:gs pos="${Math.round(stop.position * 100_000)}">${formatColorTransformXml(stop.color)}</a:gs>`,
    )
    .join("");
  const linXml =
    fill.linear === undefined
      ? ""
      : `<a:lin ang="${Math.round(fill.linear.angle * 60_000)}"${
          fill.linear.scaled === undefined
            ? ""
            : ` scaled="${fill.linear.scaled ? "1" : "0"}"`
        }/>`;
  const pathXml =
    fill.path === undefined
      ? ""
      : `<a:path path="${fill.path.shape}">${
          fill.path.fillToRect === undefined
            ? ""
            : `<a:fillToRect l="${Math.round(fill.path.fillToRect.l * 100_000)}" t="${Math.round(fill.path.fillToRect.t * 100_000)}" r="${Math.round(fill.path.fillToRect.r * 100_000)}" b="${Math.round(fill.path.fillToRect.b * 100_000)}"/>`
        }</a:path>`;
  return `<a:gradFill${rotAttr}><a:gsLst>${stops}</a:gsLst>${linXml}${pathXml}</a:gradFill>`;
}

function formatLineStyleXml(line: FormatLineStyle): string {
  const attrs = [
    line.widthEmu === undefined ? "" : ` w="${line.widthEmu}"`,
    line.cap === undefined ? "" : ` cap="${line.cap}"`,
    line.compound === undefined ? "" : ` cmpd="${line.compound}"`,
    line.align === undefined ? "" : ` algn="${line.align}"`,
  ].join("");
  const fill =
    line.fill === undefined
      ? `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`
      : formatFillStyleXml(line.fill);
  const dash =
    line.dash === undefined
      ? ""
      : `<a:prstDash val="${xmlEscape(line.dash)}"/>`;
  return `<a:ln${attrs}>${fill}${dash}</a:ln>`;
}

function formatEffectStyleXml(effect: FormatEffectStyle): string {
  if (effect.shadow === undefined) {
    return "<a:effectStyle><a:effectLst/></a:effectStyle>";
  }
  const blur = Math.round((effect.shadow.blur ?? 0) * 12_700);
  const dist = Math.round((effect.shadow.distance ?? 0) * 12_700);
  const dir = Math.round((effect.shadow.angle ?? 0) * 60_000);
  const color =
    typeof effect.shadow.color === "string" &&
    effect.shadow.color.startsWith("#")
      ? `<a:srgbClr val="${normalizedHex6(effect.shadow.color, "shadow.color")}"/>`
      : '<a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr>';
  return `<a:effectStyle><a:effectLst><a:outerShdw blurRad="${blur}" dist="${dist}" dir="${dir}" rotWithShape="0">${color}</a:outerShdw></a:effectLst></a:effectStyle>`;
}

function formatSchemeXml(scheme?: FormatScheme): string {
  const fill = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const line = `<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  const effect = `<a:effectStyle><a:effectLst/></a:effectStyle>`;
  if (scheme === undefined) {
    return `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst><a:lnStyleLst>${line}${line}${line}</a:lnStyleLst><a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst><a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>`;
  }
  const fillList =
    scheme.fillStyles !== undefined && scheme.fillStyles.length > 0
      ? scheme.fillStyles.map(formatFillStyleXml).join("")
      : `${fill}${fill}${fill}`;
  const lnList =
    scheme.lineStyles !== undefined && scheme.lineStyles.length > 0
      ? scheme.lineStyles.map(formatLineStyleXml).join("")
      : `${line}${line}${line}`;
  const effList =
    scheme.effectStyles !== undefined && scheme.effectStyles.length > 0
      ? scheme.effectStyles.map(formatEffectStyleXml).join("")
      : `${effect}${effect}${effect}`;
  const bgList =
    scheme.bgFillStyles !== undefined && scheme.bgFillStyles.length > 0
      ? scheme.bgFillStyles.map(formatFillStyleXml).join("")
      : `${fill}${fill}${fill}`;
  return `<a:fmtScheme name="${xmlEscape(scheme.name ?? "Office")}"><a:fillStyleLst>${fillList}</a:fillStyleLst><a:lnStyleLst>${lnList}</a:lnStyleLst><a:effectStyleLst>${effList}</a:effectStyleLst><a:bgFillStyleLst>${bgList}</a:bgFillStyleLst></a:fmtScheme>`;
}

/**
 * Serializes a resolved deck or master theme as a native theme part.
 */
export function themeXml(value: Theme): string {
  const name = xmlEscape(value.name);
  const colors = colorSchemeXml(value.name, value.colors);
  const fonts = fontSchemeXml(value.name, value.fonts.major, value.fonts.minor);
  const format = formatSchemeXml(value.formatScheme);
  return xmlDocument(
    `<a:theme xmlns:a="${namespaces.a}" name="${name}"><a:themeElements>${colors}${fonts}${format}</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`,
  );
}

function fontSizeCentipoints(value: FontSize): number {
  if (typeof value === "string" && value.endsWith("pt")) {
    return Math.round(Number(value.slice(0, -2)) * 100);
  }
  if (typeof value === "string" && value.endsWith("px")) {
    return Math.round(Number(value.slice(0, -2)) * 75);
  }
  return Math.round(Number(value) * 75);
}

function lengthCentipoints(value: FontSize): number {
  if (typeof value === "string" && value.endsWith("pt")) {
    return Math.round(Number(value.slice(0, -2)) * 100);
  }
  if (typeof value === "string" && value.endsWith("px")) {
    return Math.round(Number(value.slice(0, -2)) * 75);
  }
  return Math.round(Number(value) * 75);
}

function ooxmlAlign(value: TextAlignment): string {
  switch (value) {
    case "left":
      return "l";
    case "center":
      return "ctr";
    case "right":
      return "r";
    case "justify":
      return "just";
  }
}

function ooxmlAnchor(value: NonNullable<TextStyle["verticalAlign"]>): string {
  return value === "middle" ? "ctr" : value === "bottom" ? "b" : "t";
}

export function placeholderRefXml(ref: NativePlaceholderRef): string {
  const attrs: string[] = [];
  if (ref.explicitType !== false) attrs.push(`type="${xmlEscape(ref.type)}"`);
  if (ref.sz !== undefined) attrs.push(`sz="${xmlEscape(ref.sz)}"`);
  if (ref.orient !== undefined) attrs.push(`orient="${xmlEscape(ref.orient)}"`);
  if (ref.idx !== undefined) attrs.push(`idx="${ref.idx}"`);
  if (ref.hasCustomPrompt === true) attrs.push('hasCustomPrompt="1"');
  return `<p:ph${attrs.length === 0 ? "" : ` ${attrs.join(" ")}`}/>`;
}

function levelDefRPrXml(
  style: PlaceholderLevelStyle,
  themeContext: ThemeContext,
): string {
  const attrs: string[] = [];
  if (style.lang !== undefined) attrs.push(`lang="${xmlEscape(style.lang)}"`);
  if (style.fontSize !== undefined) {
    attrs.push(`sz="${fontSizeCentipoints(style.fontSize)}"`);
  }
  if (style.fontWeight !== undefined) {
    const bold =
      style.fontWeight === "bold" ||
      (typeof style.fontWeight === "number" && style.fontWeight >= 600);
    attrs.push(`b="${bold ? 1 : 0}"`);
  }
  if (style.fontStyle !== undefined) {
    attrs.push(`i="${style.fontStyle === "italic" ? 1 : 0}"`);
  }
  if (style.textDecoration === "underline") attrs.push('u="sng"');
  else if (style.textDecoration === "line-through")
    attrs.push('strike="sngStrike"');
  if (style.textTransform === "uppercase") attrs.push('cap="all"');
  if (style.letterSpacing !== undefined) {
    attrs.push(`spc="${Math.round(style.letterSpacing * 75)}"`);
  }

  let fillXml = "";
  if (style.color !== undefined) {
    const resolved = themeContext.resolveColor(style.color);
    fillXml = `<a:solidFill>${resolvedColorXml(resolved)}</a:solidFill>`;
  }
  let fontXml = "";
  if (style.fontFamily !== undefined) {
    const resolved = themeContext.resolveFont(style.fontFamily);
    const latin =
      resolved.themeFont === "major"
        ? "+mj-lt"
        : resolved.themeFont === "minor"
          ? "+mn-lt"
          : resolved.family;
    const ea =
      resolved.themeFont === "major"
        ? "+mj-ea"
        : resolved.themeFont === "minor"
          ? "+mn-ea"
          : resolved.family;
    const cs =
      resolved.themeFont === "major"
        ? "+mj-cs"
        : resolved.themeFont === "minor"
          ? "+mn-cs"
          : resolved.family;
    fontXml = `<a:latin typeface="${xmlEscape(latin)}"/><a:ea typeface="${xmlEscape(ea)}"/><a:cs typeface="${xmlEscape(cs)}"/>`;
  }
  if (attrs.length === 0 && fillXml === "" && fontXml === "") return "";
  const attrText = attrs.length === 0 ? "" : ` ${attrs.join(" ")}`;
  const inner = `${fillXml}${fontXml}`;
  return inner === ""
    ? `<a:defRPr${attrText}/>`
    : `<a:defRPr${attrText}>${inner}</a:defRPr>`;
}

function levelParagraphPropertiesXml(
  tagName: string,
  level: PlaceholderLevelStyle,
  themeContext: ThemeContext,
): string {
  const attrs: string[] = [];
  if (level.marginLeft !== undefined) {
    attrs.push(`marL="${pxToEmu(level.marginLeft)}"`);
  }
  if (level.indent !== undefined) {
    attrs.push(`indent="${pxToEmu(level.indent)}"`);
  }
  if (level.align !== undefined) {
    attrs.push(`algn="${ooxmlAlign(level.align)}"`);
  }
  const spcBef =
    level.spacingBefore === undefined
      ? ""
      : `<a:spcBef><a:spcPts val="${Math.round(level.spacingBefore * 75)}"/></a:spcBef>`;
  const lnSpc =
    level.lineHeight === undefined
      ? ""
      : `<a:lnSpc><a:spcPts val="${lengthCentipoints(level.lineHeight)}"/></a:lnSpc>`;
  const spcAft =
    level.spacingAfter === undefined
      ? ""
      : `<a:spcAft><a:spcPts val="${Math.round(level.spacingAfter * 75)}"/></a:spcAft>`;
  let listXml = "";
  if (level.list !== undefined) {
    if (level.list.type === "bullet") {
      listXml = `<a:buChar char="${xmlEscape(level.list.character ?? "•")}"/>`;
    } else {
      listXml = `<a:buAutoNum type="${level.list.style ?? "arabicPeriod"}"${level.list.startAt === undefined ? "" : ` startAt="${level.list.startAt}"`}/>`;
    }
  }
  const defRPr = levelDefRPrXml(level, themeContext);
  const children = `${lnSpc}${spcBef}${spcAft}${listXml}${defRPr}`;
  if (attrs.length === 0 && children === "") return "";
  const attrText = attrs.length === 0 ? "" : ` ${attrs.join(" ")}`;
  return children === ""
    ? `<${tagName}${attrText}/>`
    : `<${tagName}${attrText}>${children}</${tagName}>`;
}

function listStylesInnerXml(
  styles: PlaceholderListStyles | undefined,
  themeContext: ThemeContext,
): string {
  if (styles === undefined) return "";
  const parts: string[] = [];
  for (const levelNumber of [1, 2, 3, 4, 5, 6, 7, 8, 9] as const) {
    const entry = styles[levelNumber];
    if (entry === undefined) continue;
    const xml = levelParagraphPropertiesXml(
      `a:lvl${levelNumber}pPr`,
      entry,
      themeContext,
    );
    if (xml !== "") parts.push(xml);
  }
  return parts.join("");
}

function styleAsLevel1(
  style: LayoutPlaceholderSpec["style"],
): PlaceholderLevelStyle | undefined {
  if (style === undefined) return undefined;
  const lvl: Record<string, unknown> = {};
  for (const key of [
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
    "lineHeight",
    "align",
    "spacingAfter",
  ] as const) {
    if (style[key] !== undefined) lvl[key] = style[key];
  }
  return Object.keys(lvl).length === 0 ? undefined : lvl;
}

function placeholderBodyPrXml(style: LayoutPlaceholderSpec["style"]): string {
  if (style === undefined) return "<a:bodyPr/>";
  const attrs: string[] = [];
  if (style.textWrap !== undefined) {
    attrs.push(`wrap="${style.textWrap === "nowrap" ? "none" : "square"}"`);
  }
  const rawPad = style.padding;
  const padTop =
    style.paddingTop ??
    (typeof rawPad === "number"
      ? rawPad
      : Array.isArray(rawPad)
        ? rawPad[0]
        : undefined);
  const padRight =
    style.paddingRight ??
    (typeof rawPad === "number"
      ? rawPad
      : Array.isArray(rawPad)
        ? rawPad[1]
        : undefined);
  const padBottom =
    style.paddingBottom ??
    (typeof rawPad === "number"
      ? rawPad
      : Array.isArray(rawPad)
        ? rawPad.length === 2
          ? rawPad[0]
          : rawPad[2]
        : undefined);
  const padLeft =
    style.paddingLeft ??
    (typeof rawPad === "number"
      ? rawPad
      : Array.isArray(rawPad)
        ? rawPad.length === 2
          ? rawPad[1]
          : rawPad[3]
        : undefined);
  if (padLeft !== undefined) attrs.push(`lIns="${pxToEmu(padLeft)}"`);
  if (padTop !== undefined) attrs.push(`tIns="${pxToEmu(padTop)}"`);
  if (padRight !== undefined) attrs.push(`rIns="${pxToEmu(padRight)}"`);
  if (padBottom !== undefined) attrs.push(`bIns="${pxToEmu(padBottom)}"`);
  if (style.verticalAlign !== undefined) {
    attrs.push(`anchor="${ooxmlAnchor(style.verticalAlign)}"`);
  }
  return attrs.length === 0 ? "<a:bodyPr/>" : `<a:bodyPr ${attrs.join(" ")}/>`;
}

function placeholderShapesXml(
  placeholders: Readonly<Record<string, LayoutPlaceholderSpec>> | undefined,
  themeContext: ThemeContext,
): string {
  if (placeholders === undefined) return "";
  return Object.entries(placeholders)
    .map(([key, spec], index) => {
      const id = 2 + index;
      const name = xmlEscape(spec.name ?? key);
      const phXml = placeholderRefXml({
        type: spec.type,
        ...(spec.explicitType === undefined
          ? {}
          : { explicitType: spec.explicitType }),
        ...(spec.idx === undefined ? {} : { idx: spec.idx }),
        ...(spec.sz === undefined ? {} : { sz: spec.sz }),
        ...(spec.orient === undefined ? {} : { orient: spec.orient }),
        ...(spec.hasCustomPrompt === undefined
          ? {}
          : { hasCustomPrompt: spec.hasCustomPrompt }),
      });
      const hasBox =
        spec.x !== undefined &&
        spec.y !== undefined &&
        spec.width !== undefined &&
        spec.height !== undefined;
      const spPr = hasBox
        ? `<p:spPr><a:xfrm><a:off x="${pxToEmu(spec.x)}" y="${pxToEmu(spec.y)}"/><a:ext cx="${pxToEmu(spec.width)}" cy="${pxToEmu(spec.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
        : "<p:spPr/>";
      const mergedLvl1 =
        styleAsLevel1(spec.style) === undefined &&
        spec.listStyles?.[1] === undefined
          ? undefined
          : { ...styleAsLevel1(spec.style), ...spec.listStyles?.[1] };
      const effectiveListStyles: PlaceholderListStyles | undefined =
        mergedLvl1 === undefined && spec.listStyles === undefined
          ? undefined
          : {
              ...spec.listStyles,
              ...(mergedLvl1 === undefined ? {} : { 1: mergedLvl1 }),
            };
      const lstStyleInner = listStylesInnerXml(
        effectiveListStyles,
        themeContext,
      );
      const lstStyle =
        lstStyleInner === ""
          ? "<a:lstStyle/>"
          : `<a:lstStyle>${lstStyleInner}</a:lstStyle>`;
      const bodyPr = placeholderBodyPrXml(spec.style);
      const p =
        spec.promptText === undefined
          ? "<a:p><a:endParaRPr/></a:p>"
          : `<a:p><a:r><a:t xml:space="preserve">${xmlEscape(spec.promptText)}</a:t></a:r></a:p>`;
      return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phXml}</p:nvPr></p:nvSpPr>${spPr}<p:txBody>${bodyPr}${lstStyle}${p}</p:txBody></p:sp>`;
    })
    .join("");
}

function masterTextStylesXml(
  textStyles: MasterTextStyles | undefined,
  themeContext: ThemeContext,
): string {
  if (textStyles === undefined) {
    return "<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>";
  }
  const titleInner = listStylesInnerXml(textStyles.titleStyle, themeContext);
  const bodyInner = listStylesInnerXml(textStyles.bodyStyle, themeContext);
  const otherInner = listStylesInnerXml(textStyles.otherStyle, themeContext);
  const titleTag =
    titleInner === ""
      ? "<p:titleStyle/>"
      : `<p:titleStyle>${titleInner}</p:titleStyle>`;
  const bodyTag =
    bodyInner === ""
      ? "<p:bodyStyle/>"
      : `<p:bodyStyle>${bodyInner}</p:bodyStyle>`;
  const otherTag =
    otherInner === ""
      ? "<p:otherStyle/>"
      : `<p:otherStyle>${otherInner}</p:otherStyle>`;
  return `<p:txStyles>${titleTag}${bodyTag}${otherTag}</p:txStyles>`;
}

function effectiveThemeWithOverride(
  baseTheme: Theme,
  override: ThemeOverride | undefined,
): Theme {
  if (override === undefined) return baseTheme;
  const mergedColors = { ...baseTheme.colors };
  if (override.colors !== undefined) {
    for (const slot of nativeColorSlots) {
      const value = override.colors[slot];
      if (value !== undefined) mergedColors[slot] = value;
    }
  }
  const mergedFonts = {
    ...baseTheme.fonts,
    ...(override.fonts?.major === undefined
      ? {}
      : { major: override.fonts.major }),
    ...(override.fonts?.minor === undefined
      ? {}
      : { minor: override.fonts.minor }),
  };
  return {
    ...baseTheme,
    colors: mergedColors,
    fonts: mergedFonts,
    ...(override.formatScheme === undefined
      ? baseTheme.formatScheme === undefined
        ? {}
        : { formatScheme: baseTheme.formatScheme }
      : { formatScheme: override.formatScheme }),
  };
}

export function hasThemeOverrideContent(
  override: ThemeOverride | undefined,
): boolean {
  if (override === undefined) return false;
  return (
    (override.colors !== undefined &&
      Object.keys(override.colors).length > 0) ||
    (override.fonts !== undefined &&
      (override.fonts.major !== undefined ||
        override.fonts.minor !== undefined)) ||
    override.formatScheme !== undefined
  );
}

export function themeOverrideXml(
  override: ThemeOverride,
  baseTheme: Theme,
): string {
  let colorsXml = "";
  if (override.colors !== undefined) {
    const merged = {} as Record<NativeColorSlot, string>;
    for (const slot of nativeColorSlots) {
      merged[slot] = override.colors[slot] ?? baseTheme.colors[slot];
    }
    colorsXml = colorSchemeXml(baseTheme.name, merged);
  }
  const fontsXml =
    override.fonts === undefined
      ? ""
      : fontSchemeXml(
          baseTheme.name,
          override.fonts.major ?? baseTheme.fonts.major,
          override.fonts.minor ?? baseTheme.fonts.minor,
        );
  const formatXml =
    override.formatScheme === undefined
      ? ""
      : formatSchemeXml(override.formatScheme);
  return xmlDocument(
    `<a:themeOverride xmlns:a="${namespaces.a}">${colorsXml}${fontsXml}${formatXml}</a:themeOverride>`,
  );
}

export function colorMapAttributesXml(map: ColorMap): string {
  return schemeColorSlots.map((slot) => `${slot}="${map[slot]}"`).join(" ");
}

export function colorMapOverrideXml(override?: ColorMapOverride): string {
  if (override === undefined || override === "master") {
    return "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>";
  }
  return `<p:clrMapOvr><a:overrideClrMapping ${colorMapAttributesXml(override)}/></p:clrMapOvr>`;
}

function headerFooterXml(hf?: HeaderFooter): string {
  if (hf === undefined) return "";
  const attrs = [
    hf.slideNumber === undefined
      ? ""
      : ` sldNum="${hf.slideNumber ? "1" : "0"}"`,
    hf.header === undefined ? "" : ` hdr="${hf.header ? "1" : "0"}"`,
    hf.footer === undefined ? "" : ` ftr="${hf.footer ? "1" : "0"}"`,
    hf.dateTime === undefined ? "" : ` dt="${hf.dateTime ? "1" : "0"}"`,
  ].join("");
  return `<p:hf${attrs}/>`;
}

function masterXml(
  spec: EmittedSlideMasterSpec,
  layoutIds: readonly { readonly id: number; readonly rId: string }[],
  themeContext: ThemeContext,
): string {
  const colorMap = `<p:clrMap ${colorMapAttributesXml(spec.colorMap ?? defaultColorMap)}/>`;
  const preserveAttr =
    spec.preserve === undefined
      ? ""
      : ` preserve="${spec.preserve ? "1" : "0"}"`;
  const nameAttr =
    spec.name === undefined ? "" : ` name="${xmlEscape(spec.name)}"`;
  const layoutIdList = layoutIds
    .map(
      ({ id, rId }) => `<p:sldLayoutId id="${id}" r:id="${xmlEscape(rId)}"/>`,
    )
    .join("");
  const shapes = `${placeholderShapesXml(spec.placeholders, themeContext)}${spec.shapeXml ?? ""}`;
  const txStyles = masterTextStylesXml(spec.textStyles, themeContext);
  return xmlDocument(
    `<p:sldMaster xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${preserveAttr}><p:cSld${nameAttr}>${spec.backgroundXml ?? ""}${shapeTreeXml(shapes)}</p:cSld>${colorMap}<p:sldLayoutIdLst>${layoutIdList}</p:sldLayoutIdLst>${spec.transitionXml ?? ""}${headerFooterXml(spec.headerFooter)}${txStyles}</p:sldMaster>`,
  );
}

function layoutXml(
  spec: EmittedSlideLayoutSpec,
  themeContext: ThemeContext,
): string {
  const typeAttr =
    spec.type === undefined ? "" : ` type="${xmlEscape(spec.type)}"`;
  const preserveAttr =
    spec.preserve === undefined
      ? ""
      : ` preserve="${spec.preserve ? "1" : "0"}"`;
  const matchingNameAttr =
    spec.matchingName === undefined
      ? ""
      : ` matchingName="${xmlEscape(spec.matchingName)}"`;
  const userDrawnAttr =
    spec.userDrawn === undefined
      ? ""
      : ` userDrawn="${spec.userDrawn ? "1" : "0"}"`;
  const showMasterSpAttr =
    spec.showMasterShapes === false ? ' showMasterSp="0"' : "";
  const showMasterPhAnimAttr =
    spec.showMasterPlaceholderAnimations === false
      ? ' showMasterPhAnim="0"'
      : "";
  const nameAttr =
    spec.name === undefined ? "" : ` name="${xmlEscape(spec.name)}"`;
  const shapes = `${placeholderShapesXml(spec.placeholders, themeContext)}${spec.shapeXml ?? ""}`;
  return xmlDocument(
    `<p:sldLayout xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${typeAttr}${preserveAttr}${matchingNameAttr}${userDrawnAttr}${showMasterSpAttr}${showMasterPhAnimAttr}><p:cSld${nameAttr}>${spec.backgroundXml ?? ""}${shapeTreeXml(shapes)}</p:cSld>${colorMapOverrideXml(spec.colorMapOverride)}${spec.transitionXml ?? ""}${headerFooterXml(spec.headerFooter)}</p:sldLayout>`,
  );
}

function slide(
  slideId: string,
  shapeXml: string,
  slideXml: string,
  backgroundXml: string,
  hidden: boolean,
  colorMapOverride?: ColorMapOverride,
  showMasterShapes?: boolean,
): string {
  const showMasterSpAttr =
    showMasterShapes === false ? ' showMasterSp="0"' : "";
  return xmlDocument(
    `<p:sld xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${showMasterSpAttr}${hidden ? ' show="false"' : ""}><p:cSld name="${xmlEscape(slideId)}">${backgroundXml}${shapeTreeXml(shapeXml)}</p:cSld>${colorMapOverrideXml(colorMapOverride)}${slideXml}</p:sld>`,
  );
}

function sectionId(index: number, name: string): string {
  const hex = createHash("sha1")
    .update(`section:${index}:${name}`)
    .digest("hex")
    .toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}}`;
}

function sectionListXml(
  sections: readonly {
    readonly name: string;
    readonly slideIndexes: readonly number[];
  }[],
): string {
  if (sections.length === 0) return "";
  const entries = sections
    .map(
      ({ name, slideIndexes }, index) =>
        `<p14:section name="${xmlEscape(name)}" id="${sectionId(index, name)}"><p14:sldIdLst>${slideIndexes
          .map((slideIndex) => `<p14:sldId id="${255 + slideIndex}"/>`)
          .join("")}</p14:sldIdLst></p14:section>`,
    )
    .join("");
  return `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${entries}</p14:sectionLst></p:ext></p:extLst>`;
}

function coreProperties(title: string): string {
  return xmlDocument(
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEscape(title)}</dc:title></cp:coreProperties>`,
  );
}

export function countAllocatedMasterThemeParts(
  mastersInput: readonly EmittedSlideMasterSpec[] | undefined,
  fallbackTheme: Theme = defaultTheme,
): number {
  const masters =
    mastersInput !== undefined && mastersInput.length > 0
      ? mastersInput
      : defaultEmittedMasters;
  const seen = new Set<Theme>();
  let count = 0;
  for (const spec of masters) {
    const masterTheme = spec.theme ?? fallbackTheme;
    if (masterTheme !== defaultTheme && seen.has(masterTheme)) continue;
    if (masterTheme !== defaultTheme) seen.add(masterTheme);
    count += 1;
  }
  return count;
}

export function pptxParts(
  size: { readonly width: number; readonly height: number },
  slides: readonly {
    readonly id: string;
    readonly layoutKey?: string;
    readonly colorMapOverride?: ColorMapOverride;
    readonly themeOverride?: ThemeOverride;
    readonly hidden?: boolean;
    readonly showMasterShapes?: boolean;
    readonly shapeXml: string;
    readonly backgroundXml?: string;
    readonly slideXml?: string;
    readonly relationships: readonly PptxRelationship[];
  }[],
  resources: readonly PptxPackageResource[] = [],
  packageParts: readonly PptxPackagePart[] = [],
  packageRelationships: readonly PptxPackageRelationshipSet[] = [],
  presentationContent: {
    readonly beforeSlides?: string;
    readonly afterNotesSize?: string;
  } = {},
  presentationAttributes: readonly { name: string; value: string }[] = [],
  themeValue: Theme = defaultTheme,
  title?: string,
  sections: readonly {
    readonly name: string;
    readonly slideIndexes: readonly number[];
  }[] = [],
  mastersInput?: readonly EmittedSlideMasterSpec[],
): Record<string, string | Uint8Array> {
  const masters =
    mastersInput !== undefined && mastersInput.length > 0
      ? mastersInput
      : defaultEmittedMasters;

  // Allocate globally unique >= 2147483648 identifiers across every master and layout.
  let nextMasterLayoutId = minimumSlideLayoutId;
  let globalLayoutNumber = 0;
  const allocatedThemePartByTheme = new Map<
    Theme,
    { readonly partPath: string; readonly relTarget: string }
  >();
  const emittedThemeParts: {
    readonly partPath: string;
    readonly xml: string;
  }[] = [];
  const emittedThemeOverrideParts: {
    readonly partPath: string;
    readonly xml: string;
  }[] = [];

  const allocateThemePart = (
    masterTheme: Theme,
  ): { readonly partPath: string; readonly relTarget: string } => {
    if (masterTheme !== defaultTheme) {
      const existing = allocatedThemePartByTheme.get(masterTheme);
      if (existing !== undefined) return existing;
    }
    const themeNumber = emittedThemeParts.length + 1;
    const partPath = `ppt/theme/theme${themeNumber}.xml`;
    const allocation = {
      partPath,
      relTarget: `../theme/theme${themeNumber}.xml`,
    };
    emittedThemeParts.push({ partPath, xml: themeXml(masterTheme) });
    if (masterTheme !== defaultTheme) {
      allocatedThemePartByTheme.set(masterTheme, allocation);
    }
    return allocation;
  };

  const allocateThemeOverridePart = (
    override: ThemeOverride,
    baseTheme: Theme,
  ): string => {
    const overrideNumber = emittedThemeOverrideParts.length + 1;
    const partPath = `ppt/theme/themeOverride${overrideNumber}.xml`;
    emittedThemeOverrideParts.push({
      partPath,
      xml: themeOverrideXml(override, baseTheme),
    });
    return `../theme/themeOverride${overrideNumber}.xml`;
  };

  const allocatedMasters = masters.map((spec, masterIndex) => {
    const masterNumber = masterIndex + 1;
    const masterId = nextMasterLayoutId++;
    const masterPart = `ppt/slideMasters/slideMaster${masterNumber}.xml`;
    const masterTheme = spec.theme ?? themeValue;
    const themeAllocation = allocateThemePart(masterTheme);
    const layouts = spec.layouts.map((layoutSpec, layoutIndex) => {
      const layoutNumber = ++globalLayoutNumber;
      const layoutId = nextMasterLayoutId++;
      const layoutPart = `ppt/slideLayouts/slideLayout${layoutNumber}.xml`;
      const themeOverrideTarget = hasThemeOverrideContent(
        layoutSpec.themeOverride,
      )
        ? allocateThemeOverridePart(layoutSpec.themeOverride!, masterTheme)
        : undefined;
      return {
        spec: layoutSpec,
        layoutNumber,
        layoutId,
        layoutPart,
        masterRelId: `rId${layoutIndex + 1}`,
        themeOverrideTarget,
      };
    });
    return {
      spec,
      masterNumber,
      masterId,
      masterPart,
      masterTheme,
      themeRelTarget: themeAllocation.relTarget,
      presentationRelId: `rId${masterNumber}`,
      layouts,
    };
  });

  const layoutPartByKey = new Map<string, string>();
  const masterThemeByLayoutKey = new Map<string, Theme>();
  for (const allocatedMaster of allocatedMasters) {
    for (const allocatedLayout of allocatedMaster.layouts) {
      layoutPartByKey.set(allocatedLayout.spec.key, allocatedLayout.layoutPart);
      masterThemeByLayoutKey.set(
        allocatedLayout.spec.key,
        effectiveThemeWithOverride(
          allocatedMaster.masterTheme,
          allocatedLayout.spec.themeOverride,
        ),
      );
    }
  }
  const fallbackLayoutPart =
    allocatedMasters[0]?.layouts[0]?.layoutPart ??
    "ppt/slideLayouts/slideLayout1.xml";
  const fallbackMasterTheme = allocatedMasters[0]?.masterTheme ?? themeValue;

  const slideThemeOverrideTargets = slides.map((slideSpec) => {
    if (!hasThemeOverrideContent(slideSpec.themeOverride)) return undefined;
    const baseTheme =
      (slideSpec.layoutKey === undefined
        ? undefined
        : masterThemeByLayoutKey.get(slideSpec.layoutKey)) ??
      fallbackMasterTheme;
    return allocateThemeOverridePart(slideSpec.themeOverride!, baseTheme);
  });

  const masterRelationships: PptxRelationship[] = allocatedMasters.map(
    ({ masterNumber, presentationRelId }) => ({
      id: presentationRelId,
      type: `${namespaces.r}/slideMaster`,
      target: `slideMasters/slideMaster${masterNumber}.xml`,
    }),
  );
  const masterIdsXml = allocatedMasters
    .map(
      ({ masterId, presentationRelId }) =>
        `<p:sldMasterId id="${masterId}" r:id="${presentationRelId}"/>`,
    )
    .join("");

  const slideRelStart = allocatedMasters.length + 1;
  const slideRelationships = slides.map((_, index) => ({
    id: `rId${slideRelStart + index}`,
    type: `${namespaces.r}/slide`,
    target: `slides/slide${index + 1}.xml`,
  }));
  const slideIds = slides
    .map(
      (_, index) =>
        `<p:sldId id="${256 + index}" r:id="rId${slideRelStart + index}"/>`,
    )
    .join("");
  const slideIdListXml =
    slides.length === 0 ? "" : `<p:sldIdLst>${slideIds}</p:sldIdLst>`;
  const presentationAttributeXml = presentationAttributes
    .map(({ name, value }) => ` ${name}="${xmlEscape(value)}"`)
    .join("");
  const extraPresentationRelationships =
    packageRelationships.find(
      ({ sourcePath }) => sourcePath === "ppt/presentation.xml",
    )?.relationships ?? [];
  const presentation = xmlDocument(
    `<p:presentation xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${presentationAttributeXml}><p:sldMasterIdLst>${masterIdsXml}</p:sldMasterIdLst>${presentationContent.beforeSlides ?? ""}${slideIdListXml}<p:sldSz cx="${pxToEmu(size.width)}" cy="${pxToEmu(size.height)}"/><p:notesSz cx="6858000" cy="9144000"/>${presentationContent.afterNotesSize ?? ""}${sectionListXml(sections)}</p:presentation>`,
  );
  const parts: Record<string, string | Uint8Array> = {
    "[Content_Types].xml": contentTypes(
      allocatedMasters.length,
      globalLayoutNumber,
      emittedThemeParts.length,
      emittedThemeOverrideParts.length,
      slides.length,
      resources,
      packageParts,
      title !== undefined,
    ),
    "_rels/.rels": relationships([
      {
        id: "rId1",
        type: `${namespaces.r}/officeDocument`,
        target: "ppt/presentation.xml",
      },
      ...(title === undefined
        ? []
        : [
            {
              id: "rId2",
              type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
              target: "docProps/core.xml",
            },
          ]),
    ]),
    "ppt/presentation.xml": presentation,
    "ppt/_rels/presentation.xml.rels": relationships([
      ...masterRelationships,
      ...slideRelationships,
      ...extraPresentationRelationships,
    ]),
  };

  const defaultThemeCtx = createThemeContext(themeValue);
  for (const themePart of emittedThemeParts) {
    parts[themePart.partPath] = themePart.xml;
  }
  for (const overridePart of emittedThemeOverrideParts) {
    parts[overridePart.partPath] = overridePart.xml;
  }
  for (const allocatedMaster of allocatedMasters) {
    const masterThemeCtx =
      allocatedMaster.spec.themeContext ??
      (allocatedMaster.spec.theme === undefined
        ? defaultThemeCtx
        : createThemeContext(
            allocatedMaster.spec.theme,
            undefined,
            allocatedMaster.spec.colorMap === undefined
              ? {}
              : { colorMap: allocatedMaster.spec.colorMap },
          ));
    parts[allocatedMaster.masterPart] = masterXml(
      allocatedMaster.spec,
      allocatedMaster.layouts.map((item) => ({
        id: item.layoutId,
        rId: item.masterRelId,
      })),
      masterThemeCtx,
    );
    parts[
      `ppt/slideMasters/_rels/slideMaster${allocatedMaster.masterNumber}.xml.rels`
    ] = relationships([
      ...allocatedMaster.layouts.map((item) => ({
        id: item.masterRelId,
        type: `${namespaces.r}/slideLayout`,
        target: `../slideLayouts/slideLayout${item.layoutNumber}.xml`,
      })),
      {
        id: `rId${allocatedMaster.layouts.length + 1}`,
        type: `${namespaces.r}/theme`,
        target: allocatedMaster.themeRelTarget,
      },
      ...(allocatedMaster.spec.extraRelationships ?? []),
    ]);
    for (const allocatedLayout of allocatedMaster.layouts) {
      const layoutThemeCtx =
        allocatedLayout.spec.themeOverride !== undefined ||
        allocatedLayout.spec.colorMapOverride !== undefined
          ? masterThemeCtx.withNativeOverrides(
              allocatedLayout.spec.themeOverride,
              allocatedLayout.spec.colorMapOverride,
            )
          : masterThemeCtx;
      parts[allocatedLayout.layoutPart] = layoutXml(
        allocatedLayout.spec,
        layoutThemeCtx,
      );
      const layoutRels: PptxRelationship[] = [
        {
          id: "rId1",
          type: `${namespaces.r}/slideMaster`,
          target: `../slideMasters/slideMaster${allocatedMaster.masterNumber}.xml`,
        },
        ...(allocatedLayout.themeOverrideTarget === undefined
          ? []
          : [
              {
                id: "rId2",
                type: `${namespaces.r}/themeOverride`,
                target: allocatedLayout.themeOverrideTarget,
              },
            ]),
        ...(allocatedLayout.spec.extraRelationships ?? []),
      ];
      parts[
        `ppt/slideLayouts/_rels/slideLayout${allocatedLayout.layoutNumber}.xml.rels`
      ] = relationships(layoutRels);
    }
  }

  slides.forEach(
    (
      {
        shapeXml,
        slideXml,
        backgroundXml,
        hidden,
        showMasterShapes,
        layoutKey,
        colorMapOverride,
      },
      index,
    ) => {
      const targetLayoutPart =
        (layoutKey === undefined
          ? undefined
          : layoutPartByKey.get(layoutKey)) ?? fallbackLayoutPart;
      const targetLayoutRel = `../${targetLayoutPart.slice("ppt/".length)}`;
      parts[`ppt/slides/slide${index + 1}.xml`] = slide(
        slides[index]!.id,
        shapeXml,
        slideXml ?? "",
        backgroundXml ?? "",
        hidden ?? false,
        colorMapOverride,
        showMasterShapes,
      );
      const slideOverrideTarget = slideThemeOverrideTargets[index];
      const slideRels: PptxRelationship[] = [
        {
          id: "rId1",
          type: `${namespaces.r}/slideLayout`,
          target: targetLayoutRel,
        },
        ...slides[index]!.relationships,
        ...(slideOverrideTarget === undefined
          ? []
          : [
              {
                id: `rId${slides[index]!.relationships.length + 2}`,
                type: `${namespaces.r}/themeOverride`,
                target: slideOverrideTarget,
              },
            ]),
      ];
      parts[`ppt/slides/_rels/slide${index + 1}.xml.rels`] =
        relationships(slideRels);
    },
  );
  if (title !== undefined) parts["docProps/core.xml"] = coreProperties(title);
  for (const resource of resources) {
    parts[`ppt/media/${resource.name}`] = resource.bytes;
  }
  for (const part of packageParts) {
    parts[part.path] = part.content;
  }
  for (const relationshipSet of packageRelationships) {
    if (relationshipSet.sourcePath === "ppt/presentation.xml") continue;
    const marker = relationshipSet.sourcePath.lastIndexOf("/");
    const directory =
      marker === -1 ? "" : relationshipSet.sourcePath.slice(0, marker + 1);
    const name = relationshipSet.sourcePath.slice(marker + 1);
    parts[`${directory}_rels/${name}.rels`] = relationships(
      relationshipSet.relationships,
    );
  }
  return parts;
}
