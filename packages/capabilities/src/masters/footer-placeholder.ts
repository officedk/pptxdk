import { schemeColorNameSchema } from "@pptxdk/core";
import {
  type BoundSlideCapability,
  type CapabilityInput,
  type ColorValue,
  type FooterPlaceholderSpec,
  type FooterPlaceholders,
  type HexColor,
  type Rect,
  type HeaderFooter,
  type PlaceholderSize,
  type ResolvedFooter,
  type ResolvedFooterSlot,
  type SealedCanvasMechanism,
  type SlideCapabilityModule,
  type SlideCapabilityParseContext,
  type SlideEmitContext,
  type SlideEmitOutput,
  type SlideFooterValue,
  type SlideMasterInheritanceContext,
  type SlideNumberValue,
  type SlidePaintContext,
  type SlidePaintOutput,
  type TextAlignment,
  type TextStyle,
  type ThemeContext,
  footerPlaceholdersSchema,
  resolvePlaceholderBoxFromLayers,
  slideFooterValueSchema,
  slideNumberValueSchema,
  textStyleFrom,
} from "@pptxdk/core/capability";
import { pxToEmu, xmlEscape } from "@pptxdk/emit";
import {
  A_NS,
  extentEmuFragmentSchema,
  hexRgb6Schema,
  offsetEmuFragmentSchema,
  ooxmlBoolSchema,
  ooxmlNonNegativeIntSchema,
  P_NS,
  readOoxmlRunAttrStyles,
  readRunFontSlots,
  XML_NS,
  defineCanvasMechanism,
  emuToPx,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  emitTextBody,
  parseTextContent,
  renderTextOutlinesSvg,
  resolveTextBody,
  type ResolvedTextBody,
} from "../shared/text-service.ts";
import { footerPlaceholderCodegen } from "./masters-codegen.ts";

export type FooterPlaceholderSlotKey = "dateTime" | "footer" | "slideNumber";

type FooterPlaceholderOoxmlType = "dt" | "ftr" | "sldNum";

const slotToOoxmlType: Readonly<
  Record<FooterPlaceholderSlotKey, FooterPlaceholderOoxmlType>
> = {
  dateTime: "dt",
  footer: "ftr",
  slideNumber: "sldNum",
};

const ooxmlTypeToSlot: Readonly<
  Record<FooterPlaceholderOoxmlType, FooterPlaceholderSlotKey>
> = {
  dt: "dateTime",
  ftr: "footer",
  sldNum: "slideNumber",
};

const slotDefaultMeta: Readonly<
  Record<
    FooterPlaceholderSlotKey,
    {
      readonly idx: number;
      readonly size: PlaceholderSize;
      readonly namePrefix: string;
    }
  >
> = {
  dateTime: { idx: 2, size: "half", namePrefix: "Date Placeholder" },
  footer: { idx: 3, size: "quarter", namePrefix: "Footer Placeholder" },
  slideNumber: {
    idx: 4,
    size: "quarter",
    namePrefix: "Slide Number Placeholder",
  },
};

const slotOrder: readonly FooterPlaceholderSlotKey[] = [
  "dateTime",
  "footer",
  "slideNumber",
];

function defaultSlotBox(
  slot: FooterPlaceholderSlotKey,
  dimensions: { readonly width: number; readonly height: number },
): Rect {
  const y = Math.round(dimensions.height * 0.88);
  const height = Math.max(18, Math.round(dimensions.height * 0.06));
  switch (slot) {
    case "dateTime":
      return {
        x: Math.round(dimensions.width * 0.05),
        y,
        width: Math.round(dimensions.width * 0.25),
        height,
      };
    case "footer":
      return {
        x: Math.round(dimensions.width * 0.32),
        y,
        width: Math.round(dimensions.width * 0.36),
        height,
      };
    case "slideNumber":
      return {
        x: Math.round(dimensions.width * 0.72),
        y,
        width: Math.round(dimensions.width * 0.23),
        height,
      };
  }
}

function hasExplicitGeometry(spec: FooterPlaceholderSpec | undefined): boolean {
  if (spec === undefined) return false;
  const style = spec.style as Readonly<Record<string, unknown>> | undefined;
  if (style === undefined) return false;
  return (
    style["left"] !== undefined ||
    style["top"] !== undefined ||
    style["right"] !== undefined ||
    style["bottom"] !== undefined ||
    style["inset"] !== undefined ||
    style["width"] !== undefined ||
    style["height"] !== undefined
  );
}

function mergeFooterStyle(
  ...styles: readonly (Readonly<Record<string, unknown>> | undefined)[]
): Readonly<Record<string, unknown>> | undefined {
  const merged: Record<string, unknown> = {};
  for (const style of styles) {
    if (style === undefined) continue;
    for (const [key, value] of Object.entries(style)) {
      if (value !== undefined) {
        merged[key] = value;
      }
    }
  }
  return Object.keys(merged).length === 0 ? undefined : merged;
}

function resolveBoxFromSpecs(
  slot: FooterPlaceholderSlotKey,
  dimensions: { readonly width: number; readonly height: number },
  ...specs: readonly (FooterPlaceholderSpec | undefined)[]
): Rect {
  const fallback = defaultSlotBox(slot, dimensions);
  const resolved = resolvePlaceholderBoxFromLayers(specs, dimensions);
  return {
    x: typeof resolved.x === "number" ? resolved.x : fallback.x,
    y: typeof resolved.y === "number" ? resolved.y : fallback.y,
    width: typeof resolved.width === "number" ? resolved.width : fallback.width,
    height:
      typeof resolved.height === "number" ? resolved.height : fallback.height,
  };
}

type ResolvedFooterItem = {
  readonly slot: FooterPlaceholderSlotKey;
  readonly phType: FooterPlaceholderOoxmlType;
  readonly idx: number;
  readonly size: PlaceholderSize;
  readonly box: Rect;
  readonly emitGeometry: boolean;
  readonly name: string;
  readonly text: ResolvedTextBody;
};

type ResolvedFooterPlaceholderModel = {
  readonly items: readonly ResolvedFooterItem[];
};

function buildFooterTextBody(args: {
  readonly slot: FooterPlaceholderSlotKey;
  readonly content: string | undefined;
  readonly style: Readonly<Record<string, unknown>> | undefined;
  readonly box: Rect;
  readonly theme: ThemeContext;
  readonly slideNumber: number;
}): ResolvedTextBody {
  const { slot, content, style, box, theme, slideNumber } = args;
  const children: unknown[] =
    slot === "slideNumber"
      ? [{ type: "slide-number", props: {}, children: [] }]
      : content === undefined || content === ""
        ? []
        : [{ type: "#text", text: content }];
  const parsed = parseTextContent(
    children,
    style === undefined ? undefined : textStyleFrom(style, theme),
    theme,
    { slideNumber },
    `footerPlaceholders.${slot}.style`,
  );
  return resolveTextBody(
    parsed.paragraphs,
    parsed.style,
    theme,
    parsed.style.verticalAlign ?? "middle",
    {
      width: box.width,
      height: box.height,
      insets: { top: 0, right: 0, bottom: 0, left: 0 },
    },
    "text-box",
  );
}

function emitFooterPlaceholderItemXml(
  item: ResolvedFooterItem,
  shapeId: number,
  emitContext: SlideEmitContext,
): {
  readonly xml: string;
  readonly presentationContributions: ResolvedFooterItem["text"] extends never
    ? never
    : ReturnType<typeof emitTextBody>["presentationContributions"];
} {
  const spPrXml = item.emitGeometry
    ? `<p:spPr><a:xfrm><a:off x="${pxToEmu(item.box.x)}" y="${pxToEmu(item.box.y)}"/><a:ext cx="${pxToEmu(item.box.width)}" cy="${pxToEmu(item.box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`
    : "<p:spPr/>";
  const embedFonts = emitContext.options.embedFonts;
  const emittedBody = emitTextBody(item.text, emitContext.theme, {
    createFieldId: emitContext.createTextFieldId,
    ...(embedFonts === undefined ? {} : { embedFonts }),
  });
  const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="${xmlEscape(item.name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${item.phType}" sz="${item.size}" idx="${item.idx}"/></p:nvPr></p:nvSpPr>${spPrXml}${emittedBody.xml}</p:sp>`;
  return {
    xml,
    presentationContributions: emittedBody.presentationContributions,
  };
}

function emitFooterPlaceholderModel(
  model: ResolvedFooterPlaceholderModel,
  emitContext: SlideEmitContext,
): SlideEmitOutput {
  if (model.items.length === 0) return { xml: "" };
  let fallbackShapeId = 2;
  const nextId = (): number =>
    emitContext.allocateObjectId?.() ?? fallbackShapeId++;
  let xml = "";
  const contributions = [];
  for (const item of model.items) {
    const out = emitFooterPlaceholderItemXml(item, nextId(), emitContext);
    xml += out.xml;
    contributions.push(...out.presentationContributions);
  }
  return {
    xml,
    ...(contributions.length === 0
      ? {}
      : { presentationContributions: contributions }),
  };
}

function paintFooterPlaceholderModel(
  model: ResolvedFooterPlaceholderModel,
  paintContext: SlidePaintContext,
): SlidePaintOutput {
  if (model.items.length === 0) return { svg: "" };
  const svgParts: string[] = [];
  for (const item of model.items) {
    if (item.text.text === "") continue;
    const glyphs = renderTextOutlinesSvg({
      paragraphs: item.text.paragraphs,
      box: item.box,
      insets: { top: 0, right: 0, bottom: 0, left: 0 },
      verticalAlign: item.text.verticalAlign,
      theme: paintContext.theme,
      measuredLayout: item.text.measuredLayout,
      idPrefix: `ftr-s${paintContext.slideNumber}-${item.slot}`,
    });
    if (glyphs !== "") svgParts.push(glyphs);
  }
  return { svg: svgParts.join("") };
}

function normalizeSlideSlotOverride(
  slot: FooterPlaceholderSlotKey,
  value: SlideFooterValue | SlideNumberValue | undefined,
):
  | {
      readonly explicit: false;
    }
  | {
      readonly explicit: true;
      readonly visible: boolean;
      readonly text?: string;
      readonly style?: Readonly<Record<string, unknown>>;
      readonly idx?: number;
      readonly size?: PlaceholderSize;
    } {
  if (value === undefined) return { explicit: false };
  if (typeof value === "boolean") {
    return { explicit: true, visible: value };
  }
  if (typeof value === "string") {
    return { explicit: true, visible: true, text: value };
  }
  return {
    explicit: true,
    visible: value.visible ?? true,
    ...(typeof value === "object" && "text" in value && value.text !== undefined
      ? { text: value.text }
      : {}),
    ...(value.style === undefined ? {} : { style: value.style }),
    ...(value.idx === undefined ? {} : { idx: value.idx }),
    ...(value.size === undefined ? {} : { size: value.size }),
  };
}

function isHeaderFooterEnabled(
  slot: FooterPlaceholderSlotKey,
  masterHf?: HeaderFooter,
  layoutHf?: HeaderFooter,
): boolean {
  const layoutFlag = layoutHf?.[slot];
  if (layoutFlag !== undefined) return layoutFlag;
  const masterFlag = masterHf?.[slot];
  if (masterFlag !== undefined) return masterFlag;
  return true;
}

function resolveSlideFooterModel(args: {
  readonly slideProps: Readonly<Record<string, unknown>>;
  readonly inheritance?: SlideMasterInheritanceContext;
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly theme: ThemeContext;
  readonly slideNumber: number;
}): ResolvedFooterPlaceholderModel {
  const { slideProps, inheritance, dimensions, theme, slideNumber } = args;
  const masterPlaceholders = inheritance?.master?.footerPlaceholders;
  const layoutPlaceholders = inheritance?.layout?.footerPlaceholders;
  const masterHf = inheritance?.master?.headerFooter;
  const layoutHf = inheritance?.layout?.headerFooter;

  const parsedDateTime =
    slideProps["dateTime"] === undefined
      ? undefined
      : slideFooterValueSchema.parse(slideProps["dateTime"]);
  const parsedFooter =
    slideProps["footer"] === undefined
      ? undefined
      : slideFooterValueSchema.parse(slideProps["footer"]);
  const parsedSlideNumber =
    slideProps["slideNumber"] === undefined
      ? undefined
      : slideNumberValueSchema.parse(slideProps["slideNumber"]);

  const items: ResolvedFooterItem[] = [];
  for (const slot of slotOrder) {
    const rawSlideValue =
      slot === "dateTime"
        ? parsedDateTime
        : slot === "footer"
          ? parsedFooter
          : parsedSlideNumber;
    const slideOverride = normalizeSlideSlotOverride(slot, rawSlideValue);
    const masterSpec = masterPlaceholders?.[slot];
    const layoutSpec = layoutPlaceholders?.[slot];

    const hfEnabled = isHeaderFooterEnabled(slot, masterHf, layoutHf);
    let active = false;
    if (slideOverride.explicit) {
      active = slideOverride.visible && hfEnabled;
    } else if (hfEnabled) {
      if (layoutSpec?.visible !== undefined) {
        active = layoutSpec.visible;
      } else if (layoutSpec !== undefined) {
        active = masterSpec?.visible ?? true;
      } else if (masterSpec !== undefined) {
        active = masterSpec.visible ?? true;
      }
    }
    if (!active) continue;

    const defaults = slotDefaultMeta[slot];
    const slideStyleSpec: FooterPlaceholderSpec | undefined =
      slideOverride.explicit && slideOverride.style !== undefined
        ? { style: slideOverride.style }
        : undefined;
    const box = resolveBoxFromSpecs(
      slot,
      dimensions,
      masterSpec,
      layoutSpec,
      slideStyleSpec,
    );
    const style = mergeFooterStyle(
      masterSpec?.style,
      layoutSpec?.style,
      slideOverride.explicit ? slideOverride.style : undefined,
    );
    const content =
      slot === "slideNumber"
        ? undefined
        : ((slideOverride.explicit && slideOverride.text !== undefined
            ? slideOverride.text
            : (layoutSpec?.text ?? masterSpec?.text)) ?? "");
    const idx =
      (slideOverride.explicit ? slideOverride.idx : undefined) ??
      layoutSpec?.idx ??
      masterSpec?.idx ??
      defaults.idx;
    const size =
      (slideOverride.explicit ? slideOverride.size : undefined) ??
      layoutSpec?.size ??
      masterSpec?.size ??
      defaults.size;
    const text = buildFooterTextBody({
      slot,
      content,
      style,
      box,
      theme,
      slideNumber,
    });
    items.push({
      slot,
      phType: slotToOoxmlType[slot],
      idx,
      size,
      box,
      // PowerPoint resolves slide placeholder transforms from SlideLayout (not skipping to SlideMaster); emit inline geometry unless the layout explicitly defines box geometry without a slide-level override.
      emitGeometry:
        !hasExplicitGeometry(layoutSpec) || slideStyleSpec !== undefined,
      name: `${defaults.namePrefix} ${idx}`,
      text,
    });
  }
  return { items };
}

function resolveContainerFooterModel(args: {
  readonly targetType: "slide-master" | "slide-layout";
  readonly footerPlaceholders: FooterPlaceholders;
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly theme: ThemeContext;
}): ResolvedFooterPlaceholderModel {
  const { targetType, footerPlaceholders, dimensions, theme } = args;
  const items: ResolvedFooterItem[] = [];
  for (const slot of slotOrder) {
    const spec = footerPlaceholders[slot];
    if (spec === undefined) continue;
    const defaults = slotDefaultMeta[slot];
    const box = resolveBoxFromSpecs(slot, dimensions, spec);
    const idx = spec.idx ?? defaults.idx;
    const size = spec.size ?? defaults.size;
    const text = buildFooterTextBody({
      slot,
      content: spec.text,
      style: spec.style,
      box,
      theme,
      slideNumber: 1,
    });
    items.push({
      slot,
      phType: slotToOoxmlType[slot],
      idx,
      size,
      box,
      emitGeometry: targetType === "slide-master" || hasExplicitGeometry(spec),
      name: `${defaults.namePrefix} ${idx}`,
      text,
    });
  }
  return { items };
}

// ---------------------------------------------------------------------------
// Import: local Zod XML schemas and reader for `<p:sp><p:nvSpPr><p:nvPr><p:ph type="dt|ftr|sldNum"/>`
// ---------------------------------------------------------------------------

const PlaceholderAttrSchema = z.object({
  type: z.enum(["dt", "ftr", "sldNum"]),
  sz: z.enum(["full", "half", "quarter"]).optional(),
  idx: ooxmlNonNegativeIntSchema.optional(),
});

const NonVisualIdAttrSchema = z.object({
  id: ooxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const SpLocksAttrSchema = z.object({
  noGrp: ooxmlBoolSchema.optional(),
});

const PrstGeomAttrSchema = z.object({
  prst: z.string().min(1),
});

const BodyPrAttrSchema = z.object({
  vert: z.string().optional(),
  wrap: z.enum(["square", "none"]).optional(),
  rtlCol: ooxmlBoolSchema.optional(),
  lIns: ooxmlNonNegativeIntSchema.optional(),
  tIns: ooxmlNonNegativeIntSchema.optional(),
  rIns: ooxmlNonNegativeIntSchema.optional(),
  bIns: ooxmlNonNegativeIntSchema.optional(),
  anchor: z.enum(["t", "ctr", "b"]).optional(),
});

const LvlPPrAttrSchema = z.object({
  algn: z.enum(["l", "ctr", "r", "just"]).optional(),
});

const FieldAttrSchema = z.object({
  id: z.string().min(1),
  type: z.string().min(1),
});

const percentValAttrSchema = z.object({
  val: ooxmlNonNegativeIntSchema,
});

const SrgbAttrSchema = z.object({
  val: hexRgb6Schema,
});

const SchemeClrAttrSchema = z.object({
  val: schemeColorNameSchema,
});

const algnMap: Readonly<Record<"l" | "ctr" | "r" | "just", TextAlignment>> = {
  l: "left",
  ctr: "center",
  r: "right",
  just: "justify",
};

const anchorMap: Readonly<
  Record<"t" | "ctr" | "b", NonNullable<TextStyle["verticalAlign"]>>
> = {
  t: "top",
  ctr: "middle",
  b: "bottom",
};

function directDomChild(
  parent: XmlElementCursor["element"],
  namespaceUri: string,
  localName: string,
): XmlElementCursor["element"] | undefined {
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElementCursor["element"];
    if (
      (el.namespaceURI ?? "") === namespaceUri &&
      (el.localName ?? el.tagName) === localName
    ) {
      return el;
    }
  }
  return undefined;
}

function readFooterPlaceholderOoxmlType(
  cursor: XmlElementCursor,
): FooterPlaceholderOoxmlType | undefined {
  const nvSpPr = directDomChild(cursor.element, P_NS, "nvSpPr");
  if (nvSpPr === undefined) return undefined;
  const nvPr = directDomChild(nvSpPr, P_NS, "nvPr");
  if (nvPr === undefined) return undefined;
  const ph = directDomChild(nvPr, P_NS, "ph");
  if (ph === undefined) return undefined;
  const rawType = ph.getAttribute("type");
  if (rawType === "dt" || rawType === "ftr" || rawType === "sldNum") {
    return rawType;
  }
  return undefined;
}

function readFooterXmlColor(
  fillCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ColorValue | undefined {
  const srgb = fillCursor.child(A_NS, "srgbClr");
  const scheme =
    srgb === undefined ? fillCursor.child(A_NS, "schemeClr") : undefined;
  const colorCursor = srgb ?? scheme;
  if (colorCursor === undefined) return undefined;

  let slot: z.infer<typeof schemeColorNameSchema> | undefined;
  let hex: HexColor | undefined;
  if (srgb !== undefined) {
    const parsed = SrgbAttrSchema.safeParse(srgb.attrsObject(["val"]));
    if (!parsed.success) return undefined;
    hex = `#${parsed.data.val.toUpperCase()}`;
  } else if (scheme !== undefined) {
    const parsed = SchemeClrAttrSchema.safeParse(scheme.attrsObject(["val"]));
    if (!parsed.success) return undefined;
    slot = parsed.data.val;
  }

  let lumMod: number | undefined;
  let lumOff: number | undefined;
  let alpha: number | undefined;
  for (const mod of colorCursor.remainingElementChildren()) {
    if (mod.namespaceUri !== A_NS) continue;
    const parsed = percentValAttrSchema.safeParse(mod.attrsObject(["val"]));
    if (!parsed.success) continue;
    const fraction = Math.min(1, Math.max(0, parsed.data.val / 100_000));
    if (mod.localName === "tint") {
      lumMod = fraction;
      lumOff = 1 - fraction;
    } else if (mod.localName === "lumMod") {
      lumMod = fraction;
    } else if (mod.localName === "lumOff") {
      lumOff = fraction;
    } else if (mod.localName === "alpha") {
      alpha = fraction;
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: mod.xpath },
        namespaceUri: mod.namespaceUri,
        localName: mod.localName,
        mechanismHint: `color-transform:${mod.localName}`,
      });
    }
  }

  const baseColor: ColorValue | undefined =
    slot !== undefined
      ? lumMod !== undefined
        ? {
            slot,
            lumMod,
            ...(lumOff === undefined ? {} : { lumOff }),
          }
        : slot
      : hex;
  if (baseColor === undefined) return undefined;
  return alpha !== undefined && alpha < 1
    ? { color: baseColor, alpha }
    : baseColor;
}

function readFooterRunStyle(
  rPrCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): TextStyle {
  const attrStyles = readOoxmlRunAttrStyles({
    rPrCursor,
    location,
    diagnostics,
    allowBaseline: false,
    allowNoneTextDecoration: false,
    diagnoseUnsupportedEnums: false,
  });
  const solidFill = rPrCursor.child(A_NS, "solidFill");
  const color =
    solidFill === undefined
      ? undefined
      : readFooterXmlColor(solidFill, location, diagnostics);
  const { fontFamily, eastAsianFontFamily } = readRunFontSlots(rPrCursor);
  return {
    ...(attrStyles.fontSize === undefined
      ? {}
      : { fontSize: attrStyles.fontSize }),
    ...(attrStyles.fontWeight === undefined
      ? {}
      : { fontWeight: attrStyles.fontWeight }),
    ...(attrStyles.fontStyle === undefined
      ? {}
      : { fontStyle: attrStyles.fontStyle }),
    ...(attrStyles.lang === undefined ? {} : { lang: attrStyles.lang }),
    ...(attrStyles.letterSpacing === undefined
      ? {}
      : { letterSpacing: attrStyles.letterSpacing }),
    ...(color === undefined ? {} : { color }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(eastAsianFontFamily === undefined ? {} : { eastAsianFontFamily }),
  };
}

function readSpPrTransformBox(
  spPr: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): Rect | undefined {
  const prstGeom = spPr.child(A_NS, "prstGeom");
  if (prstGeom !== undefined) {
    PrstGeomAttrSchema.safeParse(prstGeom.attrsObject(["prst"]));
    prstGeom.ignoreEmptyChild(A_NS, "avLst");
  }
  const xfrm = spPr.child(A_NS, "xfrm");
  if (xfrm === undefined) return undefined;
  const off = xfrm.child(A_NS, "off");
  const ext = xfrm.child(A_NS, "ext");
  if (off === undefined || ext === undefined) return undefined;
  const parsedOff = offsetEmuFragmentSchema.safeParse(
    off.attrsObject(["x", "y"]),
  );
  const parsedExt = extentEmuFragmentSchema.safeParse(
    ext.attrsObject(["cx", "cy"]),
  );
  if (!parsedOff.success || !parsedExt.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: xfrm.xpath },
      attributeOrElement: "a:xfrm",
    });
    return undefined;
  }
  return {
    x: emuToPx(parsedOff.data.x),
    y: emuToPx(parsedOff.data.y),
    width: Math.max(1, emuToPx(parsedExt.data.cx)),
    height: Math.max(1, emuToPx(parsedExt.data.cy)),
  };
}

function readFooterPlaceholderImport(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, containerNode } = context;
  const location: ImportLocation = {
    part,
    xpath: cursor.xpath,
    ...(containerNode !== undefined &&
    (containerNode.kind === "slide-master" ||
      containerNode.kind === "slide-layout")
      ? {}
      : { slideIndex }),
  };
  const diagnostics: ImportDiagnostic[] = [];

  const nvSpPr = cursor.child(P_NS, "nvSpPr");
  if (nvSpPr === undefined) return {};
  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    NonVisualIdAttrSchema.safeParse(cNvPr.attrsObject(["id", "name"]));
  }
  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    const spLocks = cNvSpPr.child(A_NS, "spLocks");
    if (spLocks !== undefined) {
      SpLocksAttrSchema.safeParse(spLocks.attrsObject(["noGrp"]));
    }
  }
  const nvPr = nvSpPr.child(P_NS, "nvPr");
  const ph = nvPr?.child(P_NS, "ph");
  if (ph === undefined) return {};
  const parsedPh = PlaceholderAttrSchema.safeParse(
    ph.attrsObject(["type", "sz", "idx"]),
  );
  if (!parsedPh.success) return {};

  const slot = ooxmlTypeToSlot[parsedPh.data.type];
  const spPr = cursor.child(P_NS, "spPr");
  const box =
    spPr === undefined
      ? undefined
      : readSpPrTransformBox(spPr, location, diagnostics);

  let style: TextStyle | undefined;
  const textParts: string[] = [];

  const txBody = cursor.child(P_NS, "txBody");
  if (txBody !== undefined) {
    const bodyPr = txBody.child(A_NS, "bodyPr");
    if (bodyPr !== undefined) {
      const parsedBody = BodyPrAttrSchema.safeParse(
        bodyPr.attrsObject([
          "vert",
          "wrap",
          "rtlCol",
          "lIns",
          "tIns",
          "rIns",
          "bIns",
          "anchor",
        ]),
      );
      if (parsedBody.success && parsedBody.data.anchor !== undefined) {
        style = mergeFooterStyle(style, {
          verticalAlign: anchorMap[parsedBody.data.anchor],
        });
      }
      bodyPr.ignoreEmptyChild(A_NS, "noAutofit");
      bodyPr.ignoreEmptyChild(A_NS, "normAutofit");
      bodyPr.ignoreEmptyChild(A_NS, "spAutoFit");
    }

    const lstStyle = txBody.child(A_NS, "lstStyle");
    if (lstStyle !== undefined) {
      const lvl1 = lstStyle.child(A_NS, "lvl1pPr");
      if (lvl1 !== undefined) {
        const parsedLvl1 = LvlPPrAttrSchema.safeParse(
          lvl1.attrsObject(["algn"]),
        );
        if (parsedLvl1.success && parsedLvl1.data.algn !== undefined) {
          style = mergeFooterStyle(style, {
            align: algnMap[parsedLvl1.data.algn],
          });
        }
        const defRPr = lvl1.child(A_NS, "defRPr");
        if (defRPr !== undefined) {
          style = mergeFooterStyle(
            style,
            readFooterRunStyle(defRPr, location, diagnostics),
          );
        }
      }
    }

    for (const pCursor of txBody.children(A_NS, "p")) {
      const pPr = pCursor.child(A_NS, "pPr");
      if (pPr !== undefined) {
        const parsedPPr = LvlPPrAttrSchema.safeParse(pPr.attrsObject(["algn"]));
        if (parsedPPr.success && parsedPPr.data.algn !== undefined) {
          style = mergeFooterStyle(style, {
            align: algnMap[parsedPPr.data.algn],
          });
        }
        const defRPr = pPr.child(A_NS, "defRPr");
        if (defRPr !== undefined) {
          style = mergeFooterStyle(
            style,
            readFooterRunStyle(defRPr, location, diagnostics),
          );
        }
      }
      let paragraphText = "";
      let hasRunOrField = false;
      for (const child of pCursor.remainingElementChildren()) {
        if (child.namespaceUri === A_NS && child.localName === "r") {
          hasRunOrField = true;
          const rPr = child.child(A_NS, "rPr");
          if (rPr !== undefined) {
            style = mergeFooterStyle(
              style,
              readFooterRunStyle(rPr, location, diagnostics),
            );
          }
          const t = child.child(A_NS, "t");
          if (t !== undefined) {
            t.consumeAttr("space", XML_NS);
            paragraphText += t.textContent;
          }
        } else if (child.namespaceUri === A_NS && child.localName === "fld") {
          hasRunOrField = true;
          FieldAttrSchema.safeParse(child.attrsObject(["id", "type"]));
          const rPr = child.child(A_NS, "rPr");
          if (rPr !== undefined) {
            style = mergeFooterStyle(
              style,
              readFooterRunStyle(rPr, location, diagnostics),
            );
          }
          child.ignoreEmptyChild(A_NS, "pPr");
          const t = child.child(A_NS, "t");
          if (t !== undefined) {
            t.consumeAttr("space", XML_NS);
            paragraphText += t.textContent;
          }
        } else if (
          child.namespaceUri === A_NS &&
          child.localName === "endParaRPr"
        ) {
          const endStyle = readFooterRunStyle(child, location, diagnostics);
          // Empty paragraphs (such as master/layout footer slots without default text) store run defaults in endParaRPr.
          if (!hasRunOrField) {
            style = mergeFooterStyle(style, endStyle);
          }
        }
      }
      textParts.push(paragraphText);
    }
  }

  const combinedText = textParts.join("\n");
  if (
    containerNode !== undefined &&
    (containerNode.kind === "slide-master" ||
      containerNode.kind === "slide-layout")
  ) {
    const existingRaw = containerNode.props["footerPlaceholders"];
    const existingParsed =
      existingRaw === undefined
        ? {}
        : footerPlaceholdersSchema.parse(existingRaw);
    const specStyle = mergeFooterStyle(
      box === undefined
        ? undefined
        : {
            position: "absolute",
            left: box.x,
            top: box.y,
            width: box.width,
            height: box.height,
          },
      style,
    ) as FooterPlaceholderSpec["style"];
    const spec: FooterPlaceholderSpec = {
      ...(parsedPh.data.idx === undefined ? {} : { idx: parsedPh.data.idx }),
      ...(parsedPh.data.sz === undefined ? {} : { size: parsedPh.data.sz }),
      ...(slot !== "slideNumber" && combinedText !== ""
        ? { text: combinedText }
        : {}),
      ...(specStyle === undefined ? {} : { style: specStyle }),
      // Imported master/layout footer templates only appear on slides that explicitly realize them.
      visible: false,
    };
    const nextFooterPlaceholders: FooterPlaceholders = {
      ...existingParsed,
      [slot]: spec,
    };
    containerNode.set({ footerPlaceholders: nextFooterPlaceholders });
  } else if (containerNode !== undefined && containerNode.kind === "slide") {
    if (slot === "slideNumber") {
      const value: SlideNumberValue =
        style === undefined &&
        parsedPh.data.idx === undefined &&
        parsedPh.data.sz === undefined
          ? true
          : {
              ...(style === undefined ? {} : { style }),
              ...(parsedPh.data.idx === undefined
                ? {}
                : { idx: parsedPh.data.idx }),
              ...(parsedPh.data.sz === undefined
                ? {}
                : { size: parsedPh.data.sz }),
            };
      containerNode.set({ slideNumber: value });
    } else {
      const value: SlideFooterValue =
        style === undefined &&
        parsedPh.data.idx === undefined &&
        parsedPh.data.sz === undefined
          ? combinedText
          : {
              ...(combinedText !== "" ? { text: combinedText } : {}),
              ...(style === undefined ? {} : { style }),
              ...(parsedPh.data.idx === undefined
                ? {}
                : { idx: parsedPh.data.idx }),
              ...(parsedPh.data.sz === undefined
                ? {}
                : { size: parsedPh.data.sz }),
            };
      containerNode.set({ [slot]: value });
    }
  }

  return diagnostics.length === 0 ? {} : { diagnostics };
}

export const footerPlaceholderShapeImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches(cursor: XmlElementCursor): boolean {
    return readFooterPlaceholderOoxmlType(cursor) !== undefined;
  },
  read: readFooterPlaceholderImport,
};

/** Native canvas mechanism for `dt`, `ftr`, and `sldNum` metadata placeholders (`footer-placeholder`). */
export const footerPlaceholderMechanism = defineCanvasMechanism({
  key: "footer-placeholder",
  role: "canvas",
  emit: emitFooterPlaceholderModel,
  paint: paintFooterPlaceholderModel,
  import: footerPlaceholderShapeImportHandler,
});

/** Slide capability resolving, painting, emitting, and importing `dateTime`, `footer`, `slideNumber`, and master/layout `footerPlaceholders`. */
export const footerPlaceholderCapability: SlideCapabilityModule<
  BoundSlideCapability,
  readonly [SealedCanvasMechanism<"footer-placeholder">]
> = {
  declaration: "footer-placeholder",
  slideProp: "footer",
  slideProps: ["footer", "dateTime", "slideNumber"],
  slot: "shape-tree",
  always: true,
  mechanisms: [footerPlaceholderMechanism] as const,
  codegen: footerPlaceholderCodegen,
  bindSlide(
    input: CapabilityInput,
    context: SlideCapabilityParseContext,
  ): BoundSlideCapability {
    const model =
      input.type === "slide-master" || input.type === "slide-layout"
        ? resolveContainerFooterModel({
            targetType: input.type,
            footerPlaceholders: footerPlaceholdersSchema.parse(
              input.props["footerPlaceholders"] ?? {},
            ),
            dimensions: context.dimensions,
            theme: context.theme,
          })
        : resolveSlideFooterModel({
            slideProps: input.props,
            ...(context.inheritance === undefined
              ? {}
              : { inheritance: context.inheritance }),
            dimensions: context.dimensions,
            theme: context.theme,
            slideNumber: context.slideNumber,
          });
    return {
      resolveSlide() {
        const claims = model.items.flatMap((item) => item.text.resourceClaims);
        const fontUsages = model.items.flatMap((item) =>
          item.text.fontUsages.map((usage) => ({
            ...usage,
            path: `footerPlaceholders.${item.slot}`,
          })),
        );
        let footer: ResolvedFooter | undefined;
        if (input.type === "slide" && model.items.length > 0) {
          const slots: {
            dateTime?: ResolvedFooterSlot;
            footer?: ResolvedFooterSlot;
            slideNumber?: ResolvedFooterSlot;
          } = {};
          for (const item of model.items) {
            if (item.text.textMetrics === undefined) continue;
            slots[item.slot] = {
              box: item.box,
              text: item.text.text,
              textMetrics: item.text.textMetrics,
            };
          }
          if (Object.keys(slots).length > 0) {
            footer = slots;
          }
        }
        return {
          extensions: [],
          ...(footer === undefined ? {} : { footer }),
          ...(claims.length === 0 ? {} : { resourceClaims: claims }),
          ...(fontUsages.length === 0 ? {} : { fontUsages }),
          emit: (emitContext: SlideEmitContext) =>
            footerPlaceholderMechanism.emit(model, emitContext),
          ...(model.items.length === 0
            ? {}
            : {
                paint: (paintContext: SlidePaintContext) =>
                  footerPlaceholderMechanism.paint(model, paintContext),
              }),
        };
      },
    };
  },
} satisfies SlideCapabilityModule;
