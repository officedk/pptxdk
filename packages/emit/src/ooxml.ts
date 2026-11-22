import { createHash } from "node:crypto";

import type {
  LayoutPlaceholderSpec,
  MasterPlaceholderSpec,
  MasterTextStyles,
  NativePlaceholderRef,
  PlaceholderLevelStyle,
  PlaceholderListStyles,
} from "@pptxdk/core";
import {
  createThemeContext,
  defaultTheme,
  isFont,
  nativeColorSlots,
  type FontSize,
  type NativeColorSlot,
  type ResolvedColor,
  type TextAlignment,
  type TextStyle,
  type Theme,
  type ThemeContext,
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
  const element = value.slot === undefined ? "srgbClr" : "schemeClr";
  const color =
    value.slot === undefined ? value.hex : nativeSchemeReference[value.slot];
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
};

export type EmittedSlideMasterSpec = {
  readonly key: string;
  readonly name?: string;
  readonly preserve?: boolean;
  readonly theme?: ThemeContext;
  readonly placeholders?: Readonly<Record<string, MasterPlaceholderSpec>>;
  readonly textStyles?: MasterTextStyles;
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
    ...Array.from({ length: masterCount }, (_, index) => [
      `/ppt/theme/theme${index + 1}.xml`,
      "application/vnd.openxmlformats-officedocument.theme+xml",
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

/**
 * Serializes a resolved deck theme as a native theme part.
 *
 * Capability modules that own a master part (for example the notes master) can write a second
 * theme part from the same author theme, keeping a distinct OPC identity while preserving its
 * colors and fonts.
 */
export function themeXml(value: Theme): string {
  const name = xmlEscape(value.name);
  const colors = `<a:clrScheme name="${name}">${nativeColorSlots
    .map((slot) => {
      const color = value.colors[slot];
      if (!/^#[0-9a-f]{6}$/iu.test(color)) {
        throw new Error(`Theme.colors.${slot} must be a #RRGGBB color`);
      }
      const schemeName =
        slot === "hlink" ? "hlink" : slot === "folHlink" ? "folHlink" : slot;
      return `<a:${schemeName}><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:${schemeName}>`;
    })
    .join("")}</a:clrScheme>`;
  const major = isFont(value.fonts.major)
    ? value.fonts.major.family
    : value.fonts.major;
  const minor = isFont(value.fonts.minor)
    ? value.fonts.minor.family
    : value.fonts.minor;
  const majorEscaped = xmlEscape(major);
  const minorEscaped = xmlEscape(minor);
  const fonts = `<a:fontScheme name="${name}"><a:majorFont><a:latin typeface="${majorEscaped}"/><a:ea typeface="${majorEscaped}"/><a:cs typeface="${majorEscaped}"/></a:majorFont><a:minorFont><a:latin typeface="${minorEscaped}"/><a:ea typeface="${minorEscaped}"/><a:cs typeface="${minorEscaped}"/></a:minorFont></a:fontScheme>`;
  const fill = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const line = `<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  const effect = `<a:effectStyle><a:effectLst/></a:effectStyle>`;
  const format = `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst><a:lnStyleLst>${line}${line}${line}</a:lnStyleLst><a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst><a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>`;
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

function placeholderRefXml(ref: NativePlaceholderRef): string {
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

function masterXml(
  spec: EmittedSlideMasterSpec,
  layoutIds: readonly { readonly id: number; readonly rId: string }[],
  themeContext: ThemeContext,
): string {
  const colorMap =
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
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
  const shapes = placeholderShapesXml(spec.placeholders, themeContext);
  const txStyles = masterTextStylesXml(spec.textStyles, themeContext);
  return xmlDocument(
    `<p:sldMaster xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${preserveAttr}><p:cSld${nameAttr}>${shapeTreeXml(shapes)}</p:cSld>${colorMap}<p:sldLayoutIdLst>${layoutIdList}</p:sldLayoutIdLst>${txStyles}</p:sldMaster>`,
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
  const shapes = placeholderShapesXml(spec.placeholders, themeContext);
  return xmlDocument(
    `<p:sldLayout xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${typeAttr}${preserveAttr}${matchingNameAttr}${userDrawnAttr}${showMasterSpAttr}${showMasterPhAnimAttr}><p:cSld${nameAttr}>${shapeTreeXml(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
}

function slide(
  slideId: string,
  shapeXml: string,
  slideXml: string,
  backgroundXml: string,
  hidden: boolean,
): string {
  return xmlDocument(
    `<p:sld xmlns:a="${namespaces.a}" xmlns:r="${namespaces.r}" xmlns:p="${namespaces.p}"${hidden ? ' show="false"' : ""}><p:cSld name="${xmlEscape(slideId)}">${backgroundXml}${shapeTreeXml(shapeXml)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${slideXml}</p:sld>`,
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

export function pptxParts(
  size: { readonly width: number; readonly height: number },
  slides: readonly {
    readonly id: string;
    readonly layoutKey?: string;
    readonly hidden?: boolean;
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
  const allocatedMasters = masters.map((spec, masterIndex) => {
    const masterNumber = masterIndex + 1;
    const masterId = nextMasterLayoutId++;
    const masterPart = `ppt/slideMasters/slideMaster${masterNumber}.xml`;
    const layouts = spec.layouts.map((layoutSpec, layoutIndex) => {
      const layoutNumber = ++globalLayoutNumber;
      const layoutId = nextMasterLayoutId++;
      const layoutPart = `ppt/slideLayouts/slideLayout${layoutNumber}.xml`;
      return {
        spec: layoutSpec,
        layoutNumber,
        layoutId,
        layoutPart,
        masterRelId: `rId${layoutIndex + 1}`,
      };
    });
    return {
      spec,
      masterNumber,
      masterId,
      masterPart,
      presentationRelId: `rId${masterNumber}`,
      layouts,
    };
  });

  const layoutPartByKey = new Map<string, string>();
  for (const allocatedMaster of allocatedMasters) {
    for (const allocatedLayout of allocatedMaster.layouts) {
      layoutPartByKey.set(allocatedLayout.spec.key, allocatedLayout.layoutPart);
    }
  }
  const fallbackLayoutPart =
    allocatedMasters[0]?.layouts[0]?.layoutPart ??
    "ppt/slideLayouts/slideLayout1.xml";

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
  const serializedTheme = themeXml(themeValue);
  for (const allocatedMaster of allocatedMasters) {
    const masterThemeCtx = allocatedMaster.spec.theme ?? defaultThemeCtx;
    // PowerPoint rejects presentations where multiple SlideMasters share one theme part.
    parts[`ppt/theme/theme${allocatedMaster.masterNumber}.xml`] =
      serializedTheme;
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
        target: `../theme/theme${allocatedMaster.masterNumber}.xml`,
      },
    ]);
    for (const allocatedLayout of allocatedMaster.layouts) {
      parts[allocatedLayout.layoutPart] = layoutXml(
        allocatedLayout.spec,
        masterThemeCtx,
      );
      parts[
        `ppt/slideLayouts/_rels/slideLayout${allocatedLayout.layoutNumber}.xml.rels`
      ] = relationships([
        {
          id: "rId1",
          type: `${namespaces.r}/slideMaster`,
          target: `../slideMasters/slideMaster${allocatedMaster.masterNumber}.xml`,
        },
      ]);
    }
  }

  slides.forEach(
    ({ shapeXml, slideXml, backgroundXml, hidden, layoutKey }, index) => {
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
      );
      parts[`ppt/slides/_rels/slide${index + 1}.xml.rels`] = relationships([
        {
          id: "rId1",
          type: `${namespaces.r}/slideLayout`,
          target: targetLayoutRel,
        },
        ...slides[index]!.relationships,
      ]);
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
