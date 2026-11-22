import {
  allocatePlaceholderAuthorKeys,
  assertValidLang,
  layoutPlaceholderSpecSchema,
  masterPlaceholderSpecSchema,
  masterTextStylesSchema,
  placeholderOrientSchema,
  placeholderRoleSchema,
  placeholderSzSchema,
  slideLayoutTypeSchema,
  type ColorMap,
  type ColorMapOverride,
  type ColorValue,
  type ContentPlaceholderRole,
  type HeaderFooter,
  type LayoutPlaceholderSpec,
  type MasterPlaceholderMap,
  type MasterPlaceholderSpec,
  type MasterTextStyles,
  type NativeColorSlot,
  type PlaceholderLevelStyle,
  type PlaceholderListStyles,
  type PlaceholderMap,
  type PlaceholderStyle,
  type SlideLayoutProps,
  type SlideMasterProps,
  type TextAlignment,
  type TextList,
  type Theme,
  type ThemeColor,
  type ThemeOverride,
} from "@pptxdk/core";
import {
  createSlideLayoutNode,
  createSlideMasterNode,
  type CommonNode,
  type DeckContainer,
  type SlideLayoutContainer,
} from "@pptxdk/core/document";
import { z } from "zod";

import { readAlternateContent } from "#/alternate-content.ts";
import type { ImportDiagnostic } from "#/diagnostics.ts";
import type {
  DeferredSlideLink,
  OmittedSlideObjectInfo,
  ShapeTreeImportHandler,
  SlideImportHandler,
} from "#/mechanism.ts";
import {
  A_NS,
  P_NS,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  REL_THEME_OVERRIDE,
  R_NS,
  XML_NS,
  emuToPx,
} from "#/namespaces.ts";
import {
  resolvePartPath,
  type ParsedRelationshipEntry,
  type PptxPackage,
} from "#/opc.ts";
import {
  inspectUnsupportedGraphicFrame,
  peekShapeTreeObjectIdentity,
  type RelationshipIndex,
} from "#/resources.ts";
import {
  isNonDefaultColorMap,
  readColorMapElement,
  readColorMapOverride,
  readThemeOverridePart,
  readThemePart,
} from "#/themes.ts";
import {
  ExtentEmuFragmentSchema,
  HexRgb6Schema,
  OffsetEmuFragmentSchema,
  OoxmlBoolSchema,
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  OoxmlPositiveIntSchema,
} from "#/xml-fragments.ts";
import { XmlElementCursor, parseXmlPart } from "#/xml-reader.ts";

const minimumMasterLayoutId = 2_147_483_648;
const maximumMasterLayoutId = 4_294_967_295;

const SlideMasterAttrsSchema = z.object({
  preserve: OoxmlBoolSchema.optional(),
});

const SlideLayoutAttrsSchema = z.object({
  type: slideLayoutTypeSchema.optional(),
  preserve: OoxmlBoolSchema.optional(),
  matchingName: z.string().optional(),
  userDrawn: OoxmlBoolSchema.optional(),
  showMasterSp: OoxmlBoolSchema.optional(),
  showMasterPhAnim: OoxmlBoolSchema.optional(),
});

const HeaderFooterAttrsSchema = z.object({
  sldNum: OoxmlBoolSchema.optional(),
  hdr: OoxmlBoolSchema.optional(),
  ftr: OoxmlBoolSchema.optional(),
  dt: OoxmlBoolSchema.optional(),
});

function readHeaderFooterElement(
  parentCursor: XmlElementCursor,
  partPath: string,
  diagnostics: ImportDiagnostic[],
): HeaderFooter | undefined {
  const hf = parentCursor.child(P_NS, "hf");
  if (hf === undefined) return undefined;
  const raw = hf.attrsObject(["sldNum", "hdr", "ftr", "dt"]);
  const parsed = HeaderFooterAttrsSchema.safeParse(raw);
  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: partPath, xpath: hf.xpath },
      attributeOrElement: "p:hf",
      rawValue: JSON.stringify(raw),
    });
    return undefined;
  }
  const result: HeaderFooter = {
    ...(parsed.data.dt === undefined ? {} : { dateTime: parsed.data.dt }),
    ...(parsed.data.ftr === undefined ? {} : { footer: parsed.data.ftr }),
    ...(parsed.data.hdr === undefined ? {} : { header: parsed.data.hdr }),
    ...(parsed.data.sldNum === undefined
      ? {}
      : { slideNumber: parsed.data.sldNum }),
  };
  return Object.keys(result).length === 0 ? {} : result;
}

const defaultColorMap: Readonly<Record<string, NativeColorSlot>> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

const directSchemeToSlot: Readonly<Record<string, NativeColorSlot>> = {
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

function readColorModifierVal(
  colorCursor: XmlElementCursor,
  localName: "alpha" | "lumMod" | "lumOff" | "tint",
): number | undefined {
  const mod = colorCursor.child(A_NS, localName);
  if (mod === undefined) return undefined;
  const parsed = OoxmlNonNegativeIntSchema.safeParse(mod.attr("val"));
  return parsed.success ? Math.min(1, parsed.data / 100_000) : undefined;
}

function readStyleColor(
  containerCursor: XmlElementCursor,
  clrMap: Readonly<Record<string, NativeColorSlot>>,
): ColorValue | undefined {
  const srgb = containerCursor.child(A_NS, "srgbClr");
  const scheme =
    srgb === undefined ? containerCursor.child(A_NS, "schemeClr") : undefined;
  const colorCursor = srgb ?? scheme;
  if (colorCursor === undefined) return undefined;

  let baseColor: ThemeColor | undefined;
  if (srgb !== undefined) {
    const parsed = HexRgb6Schema.safeParse(srgb.attr("val"));
    if (!parsed.success) return undefined;
    baseColor = `#${parsed.data.toUpperCase()}`;
  } else if (scheme !== undefined) {
    const val = scheme.attr("val");
    const slot =
      val === undefined ? undefined : (clrMap[val] ?? directSchemeToSlot[val]);
    if (slot === undefined) return undefined;
    const rawLumMod = readColorModifierVal(colorCursor, "lumMod");
    const rawLumOff = readColorModifierVal(colorCursor, "lumOff");
    const tint = readColorModifierVal(colorCursor, "tint");
    const lumMod =
      rawLumMod ??
      (rawLumOff === undefined && tint !== undefined ? tint : undefined);
    const lumOff =
      rawLumOff ??
      (rawLumMod === undefined && tint !== undefined ? 1 - tint : undefined);
    baseColor =
      lumMod !== undefined
        ? {
            slot,
            lumMod,
            ...(lumOff === undefined ? {} : { lumOff }),
          }
        : slot;
  }
  if (baseColor === undefined) return undefined;
  const alpha = readColorModifierVal(colorCursor, "alpha");
  return alpha !== undefined ? { color: baseColor, alpha } : baseColor;
}

function readDefRPr(
  rPr: XmlElementCursor,
  clrMap: Readonly<Record<string, NativeColorSlot>>,
): Partial<PlaceholderLevelStyle> {
  const raw = rPr.attrsObject([
    "lang",
    "sz",
    "b",
    "i",
    "u",
    "strike",
    "cap",
    "spc",
    "kern",
    "dirty",
    "smtClean",
    "baseline",
  ]);
  let lang: string | undefined;
  if (raw["lang"] !== undefined) {
    try {
      lang = assertValidLang(raw["lang"]);
    } catch {
      lang = undefined;
    }
  }
  const szParsed = OoxmlPositiveIntSchema.safeParse(raw["sz"]);
  const fontSize = szParsed.success
    ? (`${szParsed.data / 100}pt` as const)
    : undefined;
  const bParsed = OoxmlBoolSchema.safeParse(raw["b"]);
  const fontWeight = bParsed.success
    ? bParsed.data
      ? ("bold" as const)
      : ("normal" as const)
    : undefined;
  const iParsed = OoxmlBoolSchema.safeParse(raw["i"]);
  const fontStyle = iParsed.success
    ? iParsed.data
      ? ("italic" as const)
      : ("normal" as const)
    : undefined;
  const textDecoration =
    raw["u"] === "sng"
      ? ("underline" as const)
      : raw["strike"] === "sngStrike"
        ? ("line-through" as const)
        : raw["u"] === "none" && raw["strike"] === "noStrike"
          ? ("none" as const)
          : undefined;
  const textTransform =
    raw["cap"] === "all"
      ? ("uppercase" as const)
      : raw["cap"] === "none"
        ? ("none" as const)
        : undefined;
  const spcParsed = OoxmlIntSchema.safeParse(raw["spc"]);
  const letterSpacing = spcParsed.success ? spcParsed.data / 75 : undefined;

  const solidFill = rPr.child(A_NS, "solidFill");
  const color =
    solidFill === undefined ? undefined : readStyleColor(solidFill, clrMap);
  rPr.ignoreEmptyChild(A_NS, "noFill");

  const latin = rPr.child(A_NS, "latin");
  const ea = rPr.child(A_NS, "ea");
  const cs = rPr.child(A_NS, "cs");
  const fontCursor = latin ?? ea ?? cs;
  let fontFamily: string | undefined;
  if (fontCursor !== undefined) {
    const typeface = fontCursor.attr("typeface");
    fontCursor.consumeAttr("pitchFamily");
    fontCursor.consumeAttr("charset");
    fontCursor.consumeAttr("panose");
    if (typeface !== undefined && typeface.trim() !== "") {
      if (
        typeface === "+mj-lt" ||
        typeface === "+mj-ea" ||
        typeface === "+mj-cs"
      ) {
        fontFamily = "major";
      } else if (
        typeface === "+mn-lt" ||
        typeface === "+mn-ea" ||
        typeface === "+mn-cs"
      ) {
        fontFamily = "minor";
      } else {
        fontFamily = typeface;
      }
    }
  }
  for (const extra of [ea, cs]) {
    if (extra !== undefined && extra !== fontCursor) {
      extra.consumeAttr("typeface");
      extra.consumeAttr("pitchFamily");
      extra.consumeAttr("charset");
      extra.consumeAttr("panose");
    }
  }
  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(fontWeight === undefined ? {} : { fontWeight }),
    ...(fontStyle === undefined ? {} : { fontStyle }),
    ...(color === undefined ? {} : { color }),
    ...(letterSpacing === undefined ? {} : { letterSpacing }),
    ...(textTransform === undefined ? {} : { textTransform }),
    ...(textDecoration === undefined ? {} : { textDecoration }),
    ...(lang === undefined ? {} : { lang }),
  };
}

function readLevelParagraphProperties(
  pPr: XmlElementCursor,
  levelIndex0: number,
  clrMap: Readonly<Record<string, NativeColorSlot>>,
): PlaceholderLevelStyle {
  const rawAttrs = pPr.attrsObject([
    "marL",
    "indent",
    "algn",
    "defTabSz",
    "rtl",
    "eaLnBrk",
    "latinLnBrk",
    "hangingPunct",
    "fontAlgn",
  ]);
  const marLParsed = OoxmlNonNegativeIntSchema.safeParse(rawAttrs["marL"]);
  const marginLeft = marLParsed.success ? emuToPx(marLParsed.data) : undefined;
  const indentParsed = OoxmlIntSchema.safeParse(rawAttrs["indent"]);
  const indent = indentParsed.success ? emuToPx(indentParsed.data) : undefined;
  const algnMap: Readonly<Record<string, TextAlignment>> = {
    l: "left",
    ctr: "center",
    r: "right",
    just: "justify",
  };
  const align =
    rawAttrs["algn"] === undefined ? undefined : algnMap[rawAttrs["algn"]];

  const spcBef = pPr.child(A_NS, "spcBef");
  const spcBefPts = spcBef?.child(A_NS, "spcPts");
  const spcBefPct = spcBef?.child(A_NS, "spcPct");
  spcBefPct?.consumeAttr("val");
  const spcBefVal = OoxmlNonNegativeIntSchema.safeParse(spcBefPts?.attr("val"));
  const spacingBefore = spcBefVal.success ? spcBefVal.data / 75 : undefined;

  const spcAft = pPr.child(A_NS, "spcAft");
  const spcAftPts = spcAft?.child(A_NS, "spcPts");
  const spcAftPct = spcAft?.child(A_NS, "spcPct");
  spcAftPct?.consumeAttr("val");
  const spcAftVal = OoxmlNonNegativeIntSchema.safeParse(spcAftPts?.attr("val"));
  const spacingAfter = spcAftVal.success ? spcAftVal.data / 75 : undefined;

  const lnSpc = pPr.child(A_NS, "lnSpc");
  const lnSpcPts = lnSpc?.child(A_NS, "spcPts");
  const lnSpcPct = lnSpc?.child(A_NS, "spcPct");
  lnSpcPct?.consumeAttr("val");
  const lnSpcVal = OoxmlPositiveIntSchema.safeParse(lnSpcPts?.attr("val"));
  const lineHeight = lnSpcVal.success
    ? (`${lnSpcVal.data / 100}pt` as const)
    : undefined;

  pPr.ignoreEmptyChild(A_NS, "buNone");
  const buFont = pPr.child(A_NS, "buFont");
  if (buFont !== undefined) {
    buFont.consumeAttr("typeface");
    buFont.consumeAttr("pitchFamily");
    buFont.consumeAttr("charset");
    buFont.consumeAttr("panose");
  }
  const buChar = pPr.child(A_NS, "buChar");
  const buAutoNum = pPr.child(A_NS, "buAutoNum");
  let list: TextList | undefined;
  if (buChar !== undefined) {
    const ch = buChar.attr("char");
    if (ch !== undefined && ch.length > 0) {
      list = {
        type: "bullet",
        character: ch,
        ...(levelIndex0 > 0 ? { level: levelIndex0 } : {}),
      };
    }
  } else if (buAutoNum !== undefined) {
    buAutoNum.consumeAttr("type");
    const startParsed = OoxmlPositiveIntSchema.safeParse(
      buAutoNum.attr("startAt"),
    );
    list = {
      type: "number",
      style: "arabicPeriod",
      ...(startParsed.success ? { startAt: startParsed.data } : {}),
      ...(levelIndex0 > 0 ? { level: levelIndex0 } : {}),
    };
  }

  const defRPr = pPr.child(A_NS, "defRPr");
  const defStyle = defRPr === undefined ? {} : readDefRPr(defRPr, clrMap);

  return {
    ...defStyle,
    ...(list === undefined ? {} : { list }),
    ...(marginLeft === undefined ? {} : { marginLeft }),
    ...(indent === undefined ? {} : { indent }),
    ...(spacingBefore === undefined ? {} : { spacingBefore }),
    ...(align === undefined ? {} : { align }),
    ...(lineHeight === undefined ? {} : { lineHeight }),
    ...(spacingAfter === undefined ? {} : { spacingAfter }),
  };
}

function readPlaceholderListStyles(
  container: XmlElementCursor | undefined,
  clrMap: Readonly<Record<string, NativeColorSlot>>,
): PlaceholderListStyles | undefined {
  if (container === undefined) return undefined;
  const levels: Record<number, PlaceholderLevelStyle> = {};
  let count = 0;
  for (const levelNum of [1, 2, 3, 4, 5, 6, 7, 8, 9] as const) {
    const pPr = container.child(A_NS, `lvl${levelNum}pPr`);
    if (pPr === undefined) continue;
    const style = readLevelParagraphProperties(pPr, levelNum - 1, clrMap);
    levels[levelNum] = style;
    count += 1;
  }
  return count === 0 ? undefined : levels;
}

function readMasterTextStyles(
  txStyles: XmlElementCursor | undefined,
  clrMap: Readonly<Record<string, NativeColorSlot>>,
): MasterTextStyles | undefined {
  if (txStyles === undefined) return undefined;
  const titleStyleCursor = txStyles.child(P_NS, "titleStyle");
  const bodyStyleCursor = txStyles.child(P_NS, "bodyStyle");
  const otherStyleCursor = txStyles.child(P_NS, "otherStyle");
  const titleStyle = readPlaceholderListStyles(titleStyleCursor, clrMap);
  const bodyStyle = readPlaceholderListStyles(bodyStyleCursor, clrMap);
  const otherStyle = readPlaceholderListStyles(otherStyleCursor, clrMap);
  const candidate: MasterTextStyles = {
    ...(titleStyle === undefined ? {} : { titleStyle }),
    ...(bodyStyle === undefined ? {} : { bodyStyle }),
    ...(otherStyle === undefined ? {} : { otherStyle }),
  };
  if (Object.keys(candidate).length === 0) return undefined;
  const parsed = masterTextStylesSchema.safeParse(candidate);
  return parsed.success ? parsed.data : undefined;
}

function directPlaceholderChild(el: XmlElementCursor["element"]): boolean {
  for (let child = el.firstChild; child !== null; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const childEl = child as XmlElementCursor["element"];
    if (
      (childEl.namespaceURI ?? "") === P_NS &&
      (childEl.localName ?? childEl.tagName) === "nvSpPr"
    ) {
      for (let nv = childEl.firstChild; nv !== null; nv = nv.nextSibling) {
        if (nv.nodeType !== 1) continue;
        const nvEl = nv as XmlElementCursor["element"];
        if (
          (nvEl.namespaceURI ?? "") === P_NS &&
          (nvEl.localName ?? nvEl.tagName) === "nvPr"
        ) {
          for (let ph = nvEl.firstChild; ph !== null; ph = ph.nextSibling) {
            if (ph.nodeType !== 1) continue;
            const phEl = ph as XmlElementCursor["element"];
            if (
              (phEl.namespaceURI ?? "") === P_NS &&
              (phEl.localName ?? phEl.tagName) === "ph"
            ) {
              return true;
            }
          }
        }
      }
    }
  }
  return false;
}

function readPlaceholderSpList(
  spTree: XmlElementCursor | undefined,
  clrMap: Readonly<Record<string, NativeColorSlot>>,
  isMaster: boolean,
): {
  readonly placeholders?: PlaceholderMap | MasterPlaceholderMap;
  readonly placeholderKeyBySpid: ReadonlyMap<number, string>;
} {
  const placeholderKeyBySpid = new Map<number, string>();
  if (spTree === undefined) return { placeholderKeyBySpid };
  const rawSpecs: LayoutPlaceholderSpec[] = [];
  const rawSpids: (number | undefined)[] = [];
  for (const sp of spTree.matchingChildren(
    (el) =>
      (el.namespaceURI ?? "") === P_NS &&
      (el.localName ?? el.tagName) === "sp" &&
      directPlaceholderChild(el),
  )) {
    const nvSpPr = sp.child(P_NS, "nvSpPr");
    const cNvPr = nvSpPr?.child(P_NS, "cNvPr");
    const rawId = cNvPr?.attr("id");
    const parsedSpid = OoxmlNonNegativeIntSchema.safeParse(rawId);
    const spid =
      parsedSpid.success && parsedSpid.data > 0 ? parsedSpid.data : undefined;
    const rawName = cNvPr?.attr("name");
    const name =
      rawName !== undefined && rawName.trim() !== "" ? rawName : undefined;
    const cNvSpPr = nvSpPr?.child(P_NS, "cNvSpPr");
    if (cNvSpPr !== undefined) {
      cNvSpPr.consumeAttr("txBox");
      const spLocks = cNvSpPr.child(A_NS, "spLocks");
      spLocks?.consumeAttr("noGrp");
    }
    const nvPr = nvSpPr?.child(P_NS, "nvPr");
    const ph = nvPr?.child(P_NS, "ph");
    if (ph === undefined) continue;

    const rawType = ph.attr("type");
    const explicitType = rawType !== undefined;
    const typeRes = placeholderRoleSchema.safeParse(rawType ?? "obj");
    const phType: ContentPlaceholderRole = typeRes.success
      ? typeRes.data
      : "obj";
    const idxParsed = OoxmlNonNegativeIntSchema.safeParse(ph.attr("idx"));
    const idx = idxParsed.success ? idxParsed.data : undefined;
    const szParsed = placeholderSzSchema.safeParse(ph.attr("sz"));
    const orientParsed = placeholderOrientSchema.safeParse(ph.attr("orient"));
    const customPromptParsed = OoxmlBoolSchema.safeParse(
      ph.attr("hasCustomPrompt"),
    );

    let x: number | undefined;
    let y: number | undefined;
    let width: number | undefined;
    let height: number | undefined;
    let backgroundColor: ColorValue | undefined;
    const spPr = sp.child(P_NS, "spPr");
    if (spPr !== undefined) {
      const xfrm = spPr.child(A_NS, "xfrm");
      if (xfrm !== undefined) {
        xfrm.consumeAttr("rot");
        xfrm.consumeAttr("flipH");
        xfrm.consumeAttr("flipV");
        const off = xfrm.child(A_NS, "off");
        const ext = xfrm.child(A_NS, "ext");
        const parsedOff =
          off === undefined
            ? { success: false as const }
            : OffsetEmuFragmentSchema.safeParse(off.attrsObject(["x", "y"]));
        const parsedExt =
          ext === undefined
            ? { success: false as const }
            : ExtentEmuFragmentSchema.safeParse(ext.attrsObject(["cx", "cy"]));
        if (parsedOff.success && parsedExt.success) {
          x = emuToPx(parsedOff.data.x);
          y = emuToPx(parsedOff.data.y);
          width = emuToPx(parsedExt.data.cx);
          height = emuToPx(parsedExt.data.cy);
        }
      }
      const prstGeom = spPr.child(A_NS, "prstGeom");
      prstGeom?.consumeAttr("prst");
      prstGeom?.ignoreEmptyChild(A_NS, "avLst");
      spPr.ignoreEmptyChild(A_NS, "noFill");
      const solidFill = spPr.child(A_NS, "solidFill");
      if (solidFill !== undefined) {
        backgroundColor = readStyleColor(solidFill, clrMap);
      }
    }

    const styleProps: Record<string, unknown> =
      backgroundColor === undefined ? {} : { backgroundColor };
    let listStyles: PlaceholderListStyles | undefined;
    let promptText: string | undefined;
    const txBody = sp.child(P_NS, "txBody");
    if (txBody !== undefined) {
      const bodyPr = txBody.child(A_NS, "bodyPr");
      if (bodyPr !== undefined) {
        const b = bodyPr.attrsObject([
          "wrap",
          "rtlCol",
          "lIns",
          "tIns",
          "rIns",
          "bIns",
          "anchor",
          "anchorCtr",
          "numCol",
          "spcCol",
        ]);
        if (b["wrap"] === "none") styleProps["textWrap"] = "nowrap";
        else if (b["wrap"] === "square") styleProps["textWrap"] = "normal";
        if (b["anchor"] === "t") styleProps["verticalAlign"] = "top";
        else if (b["anchor"] === "ctr") styleProps["verticalAlign"] = "middle";
        else if (b["anchor"] === "b") styleProps["verticalAlign"] = "bottom";
        const lIns = OoxmlNonNegativeIntSchema.safeParse(b["lIns"]);
        const tIns = OoxmlNonNegativeIntSchema.safeParse(b["tIns"]);
        const rIns = OoxmlNonNegativeIntSchema.safeParse(b["rIns"]);
        const bIns = OoxmlNonNegativeIntSchema.safeParse(b["bIns"]);
        if (lIns.success && lIns.data !== 0)
          styleProps["paddingLeft"] = emuToPx(lIns.data);
        if (tIns.success && tIns.data !== 0)
          styleProps["paddingTop"] = emuToPx(tIns.data);
        if (rIns.success && rIns.data !== 0)
          styleProps["paddingRight"] = emuToPx(rIns.data);
        if (bIns.success && bIns.data !== 0)
          styleProps["paddingBottom"] = emuToPx(bIns.data);
        bodyPr.ignoreEmptyChild(A_NS, "noAutofit");
        bodyPr.ignoreEmptyChild(A_NS, "spAutoFit");
        const normAutofit = bodyPr.child(A_NS, "normAutofit");
        normAutofit?.consumeAttr("fontScale");
        normAutofit?.consumeAttr("lnSpcReduction");
      }

      const lstStyle = txBody.child(A_NS, "lstStyle");
      listStyles = readPlaceholderListStyles(lstStyle, clrMap);

      const promptParagraphs: string[] = [];
      let promptLvl1Fallback: Partial<PlaceholderLevelStyle> | undefined;
      for (const p of txBody.children(A_NS, "p")) {
        const pPr = p.child(A_NS, "pPr");
        const pStyle =
          pPr === undefined
            ? undefined
            : readLevelParagraphProperties(pPr, 0, clrMap);
        const runTexts: string[] = [];
        for (const r of p.children(A_NS, "r")) {
          const rPr = r.child(A_NS, "rPr");
          const rStyle =
            rPr === undefined ? undefined : readDefRPr(rPr, clrMap);
          if (
            promptLvl1Fallback === undefined &&
            (pStyle !== undefined || rStyle !== undefined)
          ) {
            promptLvl1Fallback = { ...pStyle, ...rStyle };
          }
          const t = r.child(A_NS, "t");
          if (t !== undefined) {
            t.consumeAttr("space", XML_NS);
            runTexts.push(t.textContent);
          }
        }
        const fld = p.child(A_NS, "fld");
        if (fld !== undefined) {
          fld.consumeAttr("id");
          fld.consumeAttr("type");
          const rPr = fld.child(A_NS, "rPr");
          if (rPr !== undefined) readDefRPr(rPr, clrMap);
          fld.ignoreEmptyChild(A_NS, "pPr");
          const t = fld.child(A_NS, "t");
          if (t !== undefined) {
            t.consumeAttr("space", XML_NS);
            runTexts.push(t.textContent);
          }
        }
        const endPara = p.child(A_NS, "endParaRPr");
        if (endPara !== undefined) readDefRPr(endPara, clrMap);
        if (runTexts.length > 0) promptParagraphs.push(runTexts.join(""));
      }
      if (promptParagraphs.length > 0) {
        promptText = promptParagraphs.join("\n");
      }
      if (
        listStyles === undefined &&
        promptLvl1Fallback !== undefined &&
        Object.keys(promptLvl1Fallback).length > 0
      ) {
        listStyles = { 1: promptLvl1Fallback };
      }
    }

    const style =
      Object.keys(styleProps).length === 0
        ? undefined
        : (styleProps as PlaceholderStyle);
    const candidate: LayoutPlaceholderSpec = {
      type: phType,
      ...(explicitType ? {} : { explicitType: false }),
      ...(idx === undefined ? {} : { idx }),
      ...(szParsed.success ? { sz: szParsed.data } : {}),
      ...(orientParsed.success ? { orient: orientParsed.data } : {}),
      ...(customPromptParsed.success && customPromptParsed.data
        ? { hasCustomPrompt: true }
        : {}),
      ...(name === undefined ? {} : { name }),
      ...(x === undefined ? {} : { x }),
      ...(y === undefined ? {} : { y }),
      ...(width === undefined ? {} : { width }),
      ...(height === undefined ? {} : { height }),
      ...(style === undefined ? {} : { style }),
      ...(listStyles === undefined ? {} : { listStyles }),
      ...(promptText === undefined ? {} : { promptText }),
    };
    const schema = isMaster
      ? masterPlaceholderSpecSchema
      : layoutPlaceholderSpecSchema;
    const validated = schema.safeParse(candidate);
    rawSpecs.push(validated.success ? validated.data : candidate);
    rawSpids.push(spid);
  }
  if (rawSpecs.length === 0) return { placeholderKeyBySpid };
  const keys = allocatePlaceholderAuthorKeys(rawSpecs);
  const map: Record<string, LayoutPlaceholderSpec | MasterPlaceholderSpec> = {};
  for (let i = 0; i < rawSpecs.length; i += 1) {
    const key = keys[i]!;
    map[key] = rawSpecs[i]!;
    const spid = rawSpids[i];
    if (spid !== undefined) {
      placeholderKeyBySpid.set(spid, key);
    }
  }
  return { placeholders: Object.freeze(map), placeholderKeyBySpid };
}

/** Allocates a deterministic unique author id within `used`, disambiguating collisions with `-2`, `-3`, ... */
export function allocateScopedAuthorId(
  requested: string | undefined,
  used: Set<string>,
  fallbackBase: string,
): string {
  const base =
    requested !== undefined && requested.trim().length > 0
      ? requested.trim()
      : fallbackBase;
  if (!used.has(base)) {
    used.add(base);
    return base;
  }
  let suffix = 2;
  while (used.has(`${base}-${suffix}`)) {
    suffix += 1;
  }
  const unique = `${base}-${suffix}`;
  used.add(unique);
  return unique;
}

/** Consumes standard identity `<p:nvGrpSpPr>` and `<p:grpSpPr>` transform boilerplate on an `spTree`. */
export function consumeIdentityShapeTreeBoilerplate(
  spTree: XmlElementCursor,
): void {
  const nvGrpSpPr = spTree.child(P_NS, "nvGrpSpPr");
  if (nvGrpSpPr !== undefined) {
    const cNvPr = nvGrpSpPr.child(P_NS, "cNvPr");
    if (cNvPr !== undefined) {
      cNvPr.consumeAttr("id");
      cNvPr.consumeAttr("name");
    }
    nvGrpSpPr.ignoreEmptyChild(P_NS, "cNvGrpSpPr");
    nvGrpSpPr.ignoreEmptyChild(P_NS, "nvPr");
  }
  const grpSpPr = spTree.child(P_NS, "grpSpPr");
  if (grpSpPr !== undefined) {
    const xfrm = grpSpPr.child(A_NS, "xfrm");
    if (xfrm !== undefined) {
      const off = xfrm.child(A_NS, "off");
      if (off !== undefined && off.attr("x") === "0" && off.attr("y") === "0") {
        // standard identity offset
      }
      const ext = xfrm.child(A_NS, "ext");
      if (
        ext !== undefined &&
        ext.attr("cx") === "0" &&
        ext.attr("cy") === "0"
      ) {
        // standard identity extent
      }
      const chOff = xfrm.child(A_NS, "chOff");
      if (
        chOff !== undefined &&
        chOff.attr("x") === "0" &&
        chOff.attr("y") === "0"
      ) {
        // standard identity child offset
      }
      const chExt = xfrm.child(A_NS, "chExt");
      if (
        chExt !== undefined &&
        chExt.attr("cx") === "0" &&
        chExt.attr("cy") === "0"
      ) {
        // standard identity child extent
      }
    }
  }
}

/** Records `unconsumed-xml` diagnostics for all unconsumed attributes/elements under a root cursor. */
export function appendUnconsumedDiagnostics(
  cursor: XmlElementCursor,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): void {
  for (const item of cursor.collectUnconsumed()) {
    diagnostics.push({
      code: "unconsumed-xml",
      category: "unsupported",
      disposition: "omitted",
      kind: item.kind,
      location: {
        ...item.location,
        ...(slideIndex === undefined ? {} : { slideIndex }),
      },
      namespaceUri: item.namespaceUri,
      localName: item.localName,
    });
  }
}

function validateMasterLayoutNumericId(
  rawId: string | undefined,
  part: string,
  xpath: string,
  diagnostics: ImportDiagnostic[],
): void {
  if (rawId === undefined) return;
  const parsed = OoxmlNonNegativeIntSchema.safeParse(rawId);
  if (
    !parsed.success ||
    parsed.data < minimumMasterLayoutId ||
    parsed.data > maximumMasterLayoutId
  ) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { part, xpath: `${xpath}/@id` },
      attributeOrElement: "@id",
      rawValue: rawId,
    });
  }
}

const p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const creationIdExtensionUri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";

/** Consumes benign PowerPoint `<p:cSld>/<p:extLst>` creationId metadata without altering author models. */
export function consumeCommonSlideCreationId(cSld: XmlElementCursor): void {
  const extLst = cSld.child(P_NS, "extLst");
  if (extLst === undefined) return;
  for (const ext of extLst.children(P_NS, "ext")) {
    if (ext.attr("uri") !== creationIdExtensionUri) continue;
    const creationId = ext.child(p14Namespace, "creationId");
    if (creationId !== undefined) {
      creationId.consumeAttr("val");
    }
  }
}

function stripSlideIndexFromDiagnostics(
  entries: readonly ImportDiagnostic[] | undefined,
): ImportDiagnostic[] {
  if (entries === undefined || entries.length === 0) return [];
  return entries.map((entry) => {
    const { slideIndex: _ignored, ...locationRest } = entry.location;
    return {
      ...entry,
      location: locationRest,
    };
  });
}

const supportedShapeTreeNamespaces = new Set<string>([P_NS, A_NS, R_NS]);

function readSingleShapeTreeItem(
  itemCursor: XmlElementCursor,
  pkg: PptxPackage,
  partPath: string,
  slideIndex: number | undefined,
  relIndex: RelationshipIndex,
  shapeTreeHandlers: readonly ShapeTreeImportHandler[],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void,
  diagnostics: ImportDiagnostic[],
  containerNode?: CommonNode,
  layoutPlaceholders?: PlaceholderMap,
  clrMap?: Readonly<Record<string, NativeColorSlot>>,
  registerImportedObject?: (spid: number, node: CommonNode) => void,
  omittedSpids?: Map<number, OmittedSlideObjectInfo>,
): CommonNode | undefined {
  const identity = peekShapeTreeObjectIdentity(itemCursor.element);
  const match = shapeTreeHandlers.find(
    (handler) =>
      itemCursor.namespaceUri === P_NS &&
      handler.localName === itemCursor.localName &&
      (handler.matches === undefined || handler.matches(itemCursor)),
  );
  if (match === undefined) {
    if (omittedSpids !== undefined && identity.objectId !== undefined) {
      omittedSpids.set(identity.objectId, {
        ...(identity.objectName === undefined
          ? {}
          : { objectName: identity.objectName }),
        reason: "unsupported",
      });
    }
    const frameInfo =
      itemCursor.namespaceUri === P_NS &&
      itemCursor.localName === "graphicFrame"
        ? inspectUnsupportedGraphicFrame(
            itemCursor,
            partPath,
            relIndex.byId,
            relIndex.entries,
          )
        : undefined;
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: {
        part: partPath,
        xpath: itemCursor.xpath,
        ...(slideIndex === undefined ? {} : { slideIndex }),
        ...identity,
      },
      namespaceUri: itemCursor.namespaceUri,
      localName: itemCursor.localName,
      ...(frameInfo?.mechanismHint === undefined
        ? {}
        : { mechanismHint: frameInfo.mechanismHint }),
      ...(frameInfo === undefined || frameInfo.cascaded.length === 0
        ? {}
        : { cascaded: frameInfo.cascaded }),
    });
    itemCursor.suppressUnconsumed();
    return undefined;
  }
  const outcome = match.read({
    cursor: itemCursor,
    part: partPath,
    slideIndex: slideIndex ?? 0,
    package: pkg,
    ...(containerNode === undefined ? {} : { containerNode }),
    relationships: relIndex.byId,
    registerDeferredSlideLink,
    ...(registerImportedObject === undefined ? {} : { registerImportedObject }),
    ...(layoutPlaceholders === undefined ? {} : { layoutPlaceholders }),
    ...(clrMap === undefined ? {} : { clrMap }),
  });
  if (outcome.diagnostics !== undefined) {
    diagnostics.push(
      ...(slideIndex === undefined
        ? stripSlideIndexFromDiagnostics(outcome.diagnostics)
        : outcome.diagnostics),
    );
  }
  if (outcome.node === undefined) {
    if (omittedSpids !== undefined && identity.objectId !== undefined) {
      omittedSpids.set(identity.objectId, {
        ...(identity.objectName === undefined
          ? {}
          : { objectName: identity.objectName }),
        reason: "corrupt",
      });
    }
    itemCursor.suppressUnconsumed();
    return undefined;
  }
  if (registerImportedObject !== undefined && identity.objectId !== undefined) {
    registerImportedObject(identity.objectId, outcome.node);
  }
  return outcome.node;
}

/** Shared `<p:spTree>` child reader for slides, slide layouts, and slide masters. */
export function readShapeTreeItems(args: {
  readonly spTree: XmlElementCursor;
  readonly pkg: PptxPackage;
  readonly partPath: string;
  readonly slideIndex: number | undefined;
  readonly relIndex: RelationshipIndex;
  readonly shapeTreeHandlers: readonly ShapeTreeImportHandler[];
  readonly registerDeferredSlideLink: (link: DeferredSlideLink) => void;
  readonly appendNode: (node: CommonNode) => void;
  readonly diagnostics: ImportDiagnostic[];
  readonly containerNode?: CommonNode;
  readonly layoutPlaceholders?: PlaceholderMap;
  readonly clrMap?: Readonly<Record<string, NativeColorSlot>>;
  readonly registerImportedObject?: (spid: number, node: CommonNode) => void;
  readonly omittedSpids?: Map<number, OmittedSlideObjectInfo>;
}): void {
  const {
    spTree,
    pkg,
    partPath,
    slideIndex,
    relIndex,
    shapeTreeHandlers,
    registerDeferredSlideLink,
    appendNode,
    diagnostics,
    containerNode,
    layoutPlaceholders,
    clrMap,
    registerImportedObject,
    omittedSpids,
  } = args;
  consumeIdentityShapeTreeBoilerplate(spTree);
  for (const itemCursor of spTree.remainingElementChildren()) {
    if (
      itemCursor.namespaceUri === MC_NS_IMPORT_PLACEHOLDER &&
      itemCursor.localName === "AlternateContent"
    ) {
      const selection = readAlternateContent(itemCursor, (ns) =>
        supportedShapeTreeNamespaces.has(ns),
      );
      if (selection.branch === "choice") {
        for (const childCursor of selection.children) {
          const node = readSingleShapeTreeItem(
            childCursor,
            pkg,
            partPath,
            slideIndex,
            relIndex,
            shapeTreeHandlers,
            registerDeferredSlideLink,
            diagnostics,
            containerNode,
            layoutPlaceholders,
            clrMap,
            registerImportedObject,
            omittedSpids,
          );
          if (node !== undefined) appendNode(node);
        }
      } else if (selection.branch === "fallback") {
        let adoptedFallback = false;
        for (const childCursor of selection.children) {
          const identity = peekShapeTreeObjectIdentity(childCursor.element);
          const node = readSingleShapeTreeItem(
            childCursor,
            pkg,
            partPath,
            slideIndex,
            relIndex,
            shapeTreeHandlers,
            registerDeferredSlideLink,
            diagnostics,
            containerNode,
            layoutPlaceholders,
            clrMap,
            registerImportedObject,
            omittedSpids,
          );
          if (node !== undefined) {
            adoptedFallback = true;
            appendNode(node);
            diagnostics.push({
              code: "compatibility-fallback",
              category: "unsupported",
              disposition: "fallback",
              location: {
                part: partPath,
                xpath: itemCursor.xpath,
                ...(slideIndex === undefined ? {} : { slideIndex }),
                ...identity,
              },
              preferredMechanism: selection.preferredMechanism,
              selectedFallback: node.kind,
              loss: {
                lost: "native-mechanism",
                retainedAs: node.kind,
              },
            });
          }
        }
        if (!adoptedFallback) {
          itemCursor.suppressUnconsumed();
        }
      } else {
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: {
            part: partPath,
            xpath: itemCursor.xpath,
            ...(slideIndex === undefined ? {} : { slideIndex }),
          },
          namespaceUri: itemCursor.namespaceUri,
          localName: itemCursor.localName,
          mechanismHint: selection.preferredMechanism,
        });
        itemCursor.suppressUnconsumed();
      }
      continue;
    }

    const node = readSingleShapeTreeItem(
      itemCursor,
      pkg,
      partPath,
      slideIndex,
      relIndex,
      shapeTreeHandlers,
      registerDeferredSlideLink,
      diagnostics,
      containerNode,
      layoutPlaceholders,
      clrMap,
      registerImportedObject,
      omittedSpids,
    );
    if (node !== undefined) appendNode(node);
  }
}

const MC_NS_IMPORT_PLACEHOLDER =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";

function readSlideLayoutPart(
  pkg: PptxPackage,
  layoutPartPath: string,
  ownerMasterPartPath: string,
  layoutOrder: number,
  usedLayoutIds: Set<string>,
  masterClrMap: Readonly<Record<string, NativeColorSlot>>,
  masterLevelSlideHandlers: readonly SlideImportHandler[],
  shapeTreeHandlers: readonly ShapeTreeImportHandler[],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void,
  diagnostics: ImportDiagnostic[],
): SlideLayoutContainer | undefined {
  const xmlText = pkg.readTextPart(layoutPartPath);
  if (xmlText === undefined) return undefined;

  let layoutEl: ReturnType<typeof parseXmlPart>;
  try {
    layoutEl = parseXmlPart(layoutPartPath, xmlText);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: layoutPartPath },
    });
    return undefined;
  }

  if (
    (layoutEl.namespaceURI ?? "") !== P_NS ||
    (layoutEl.localName ?? layoutEl.tagName) !== "sldLayout"
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: layoutPartPath },
    });
    return undefined;
  }

  const layoutCursor = XmlElementCursor.createRoot(
    layoutEl,
    layoutPartPath,
    "/p:sldLayout",
  );
  const rawAttrs = layoutCursor.attrsObject([
    "type",
    "preserve",
    "matchingName",
    "userDrawn",
    "showMasterSp",
    "showMasterPhAnim",
  ]);
  const parsedAttrs = SlideLayoutAttrsSchema.safeParse(rawAttrs);
  let layoutType: SlideLayoutProps["type"];
  let preserve: boolean | undefined;
  let matchingName: string | undefined;
  let userDrawn: boolean | undefined;
  let showMasterShapes: boolean | undefined;
  let showMasterPlaceholderAnimations: boolean | undefined;

  if (parsedAttrs.success) {
    layoutType = parsedAttrs.data.type;
    preserve = parsedAttrs.data.preserve;
    matchingName = parsedAttrs.data.matchingName;
    userDrawn = parsedAttrs.data.userDrawn;
    showMasterShapes = parsedAttrs.data.showMasterSp;
    showMasterPlaceholderAnimations = parsedAttrs.data.showMasterPhAnim;
  } else {
    for (const issue of parsedAttrs.error.issues) {
      const key = String(issue.path[0] ?? "sldLayout");
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: {
          part: layoutPartPath,
          xpath: `${layoutCursor.xpath}/@${key}`,
        },
        attributeOrElement: `@${key}`,
        rawValue: rawAttrs[key] ?? "",
      });
    }
    const typeRes = slideLayoutTypeSchema.safeParse(rawAttrs["type"]);
    if (typeRes.success) layoutType = typeRes.data;
    const preserveRes = OoxmlBoolSchema.safeParse(rawAttrs["preserve"]);
    if (preserveRes.success) preserve = preserveRes.data;
    if (rawAttrs["matchingName"] !== undefined) {
      matchingName = rawAttrs["matchingName"];
    }
    const userDrawnRes = OoxmlBoolSchema.safeParse(rawAttrs["userDrawn"]);
    if (userDrawnRes.success) userDrawn = userDrawnRes.data;
    const showSpRes = OoxmlBoolSchema.safeParse(rawAttrs["showMasterSp"]);
    if (showSpRes.success) showMasterShapes = showSpRes.data;
    const showPhAnimRes = OoxmlBoolSchema.safeParse(
      rawAttrs["showMasterPhAnim"],
    );
    if (showPhAnimRes.success) {
      showMasterPlaceholderAnimations = showPhAnimRes.data;
    }
  }

  const cSld = layoutCursor.child(P_NS, "cSld");
  let layoutName: string | undefined;
  if (cSld === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: layoutPartPath, xpath: layoutCursor.xpath },
    });
    return undefined;
  }
  layoutName = cSld.attr("name");
  consumeCommonSlideCreationId(cSld);
  const spTree = cSld.child(P_NS, "spTree");
  let placeholders: PlaceholderMap | undefined;
  let layoutPlaceholderKeyBySpid: ReadonlyMap<number, string> | undefined;
  if (spTree !== undefined) {
    consumeIdentityShapeTreeBoilerplate(spTree);
    const placeholderRead = readPlaceholderSpList(spTree, masterClrMap, false);
    placeholders = placeholderRead.placeholders;
    layoutPlaceholderKeyBySpid = placeholderRead.placeholderKeyBySpid;
  }
  const clrMapOvr = layoutCursor.child(P_NS, "clrMapOvr");
  let colorMapOverride: ColorMapOverride | undefined;
  if (clrMapOvr !== undefined) {
    const parsedOvr = readColorMapOverride(
      clrMapOvr,
      layoutPartPath,
      undefined,
      diagnostics,
    );
    if (parsedOvr !== undefined && parsedOvr !== "master") {
      colorMapOverride = parsedOvr;
    }
  }
  const layoutHeaderFooter = readHeaderFooterElement(
    layoutCursor,
    layoutPartPath,
    diagnostics,
  );

  const layoutRels = pkg.readRelationships(layoutPartPath);
  diagnostics.push(...layoutRels.diagnostics);
  const backMasterRel = layoutRels.entries.find(
    (entry) =>
      entry.Type === REL_SLIDE_MASTER && entry.TargetMode !== "External",
  );
  if (backMasterRel !== undefined) {
    const resolvedMaster = resolvePartPath(
      layoutPartPath,
      backMasterRel.Target,
    );
    if (resolvedMaster !== ownerMasterPartPath) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: layoutPartPath, xpath: backMasterRel.xpath },
        relationshipId: backMasterRel.Id,
      });
    }
  }

  let themeOverride: ThemeOverride | undefined;
  const themeOverrideRel = layoutRels.entries.find(
    (entry) =>
      entry.Type === REL_THEME_OVERRIDE && entry.TargetMode !== "External",
  );
  if (themeOverrideRel !== undefined) {
    const overridePart = resolvePartPath(
      layoutPartPath,
      themeOverrideRel.Target,
    );
    if (overridePart === undefined || !pkg.hasPart(overridePart)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: layoutPartPath, xpath: themeOverrideRel.xpath },
        relationshipId: themeOverrideRel.Id,
        targetPart: overridePart ?? themeOverrideRel.Target,
      });
    } else {
      themeOverride = readThemeOverridePart(
        pkg,
        overridePart,
        undefined,
        diagnostics,
      );
    }
  }

  const layoutId = allocateScopedAuthorId(
    layoutName,
    usedLayoutIds,
    `layout-${layoutOrder}`,
  );
  const props: SlideLayoutProps = {
    id: layoutId,
    ...(layoutName === undefined ? {} : { name: layoutName }),
    ...(layoutType === undefined ? {} : { type: layoutType }),
    ...(preserve === undefined ? {} : { preserve }),
    ...(matchingName === undefined ? {} : { matchingName }),
    ...(userDrawn === undefined ? {} : { userDrawn }),
    ...(showMasterShapes === undefined ? {} : { showMasterShapes }),
    ...(showMasterPlaceholderAnimations === undefined
      ? {}
      : { showMasterPlaceholderAnimations }),
    ...(placeholders === undefined ? {} : { placeholders }),
    ...(colorMapOverride === undefined ? {} : { colorMapOverride }),
    ...(themeOverride === undefined ? {} : { themeOverride }),
    ...(layoutHeaderFooter === undefined
      ? {}
      : { headerFooter: layoutHeaderFooter }),
  };
  const layoutNode = createSlideLayoutNode(props);
  if (spTree !== undefined && shapeTreeHandlers.length > 0) {
    const relIndex: RelationshipIndex = {
      entries: layoutRels.entries,
      byId: new Map(layoutRels.entries.map((entry) => [entry.Id, entry])),
      diagnostics: [],
    };
    readShapeTreeItems({
      spTree,
      pkg,
      partPath: layoutPartPath,
      slideIndex: undefined,
      relIndex,
      shapeTreeHandlers,
      registerDeferredSlideLink,
      appendNode: (node) => layoutNode.append(node),
      diagnostics,
      containerNode: layoutNode,
      clrMap: masterClrMap,
    });
  }

  for (const slideHandler of masterLevelSlideHandlers) {
    if (
      slideHandler.matches !== undefined &&
      !slideHandler.matches(layoutCursor)
    ) {
      continue;
    }
    const outcome = slideHandler.read({
      slideCursor: layoutCursor,
      slideNode: layoutNode,
      part: layoutPartPath,
      ...(slideHandler.slot === "transition" ? {} : { slideIndex: 0 }),
      package: pkg,
      ...(layoutPlaceholderKeyBySpid === undefined
        ? {}
        : { placeholderKeyBySpid: layoutPlaceholderKeyBySpid }),
    });
    diagnostics.push(...stripSlideIndexFromDiagnostics(outcome.diagnostics));
  }

  appendUnconsumedDiagnostics(layoutCursor, undefined, diagnostics);
  return layoutNode;
}

function readSlideMasterPart(
  pkg: PptxPackage,
  masterPartPath: string,
  masterOrder: number,
  usedDeckIds: Set<string>,
  themesByPartPath: Map<string, Theme>,
  layoutNodeByPartPath: Map<string, SlideLayoutContainer>,
  clrMapByLayoutPartPath: Map<
    string,
    Readonly<Record<string, NativeColorSlot>>
  >,
  masterLevelSlideHandlers: readonly SlideImportHandler[],
  shapeTreeHandlers: readonly ShapeTreeImportHandler[],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void,
  diagnostics: ImportDiagnostic[],
) {
  const xmlText = pkg.readTextPart(masterPartPath);
  if (xmlText === undefined) return undefined;

  let masterEl: ReturnType<typeof parseXmlPart>;
  try {
    masterEl = parseXmlPart(masterPartPath, xmlText);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath },
    });
    return undefined;
  }

  if (
    (masterEl.namespaceURI ?? "") !== P_NS ||
    (masterEl.localName ?? masterEl.tagName) !== "sldMaster"
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath },
    });
    return undefined;
  }

  const masterCursor = XmlElementCursor.createRoot(
    masterEl,
    masterPartPath,
    "/p:sldMaster",
  );
  const rawAttrs = masterCursor.attrsObject(["preserve"]);
  const parsedAttrs = SlideMasterAttrsSchema.safeParse(rawAttrs);
  let preserve: boolean | undefined;
  if (!parsedAttrs.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: {
        part: masterPartPath,
        xpath: `${masterCursor.xpath}/@preserve`,
      },
      attributeOrElement: "@preserve",
      rawValue: rawAttrs["preserve"] ?? "",
    });
  } else {
    preserve = parsedAttrs.data.preserve;
  }

  const cSld = masterCursor.child(P_NS, "cSld");
  if (cSld === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath, xpath: masterCursor.xpath },
    });
    return undefined;
  }
  const masterName = cSld.attr("name");
  consumeCommonSlideCreationId(cSld);
  const clrMap = masterCursor.child(P_NS, "clrMap");
  let colorMap: ColorMap | undefined;
  let masterClrMap: Readonly<Record<string, NativeColorSlot>> = defaultColorMap;
  if (clrMap !== undefined) {
    const parsedMap = readColorMapElement(
      clrMap,
      masterPartPath,
      undefined,
      diagnostics,
    );
    masterClrMap = parsedMap;
    if (isNonDefaultColorMap(parsedMap)) {
      colorMap = parsedMap;
    }
  }

  const spTree = cSld.child(P_NS, "spTree");
  let masterPlaceholders: MasterPlaceholderMap | undefined;
  let masterPlaceholderKeyBySpid: ReadonlyMap<number, string> | undefined;
  if (spTree !== undefined) {
    consumeIdentityShapeTreeBoilerplate(spTree);
    const placeholderRead = readPlaceholderSpList(spTree, masterClrMap, true);
    masterPlaceholders = placeholderRead.placeholders;
    masterPlaceholderKeyBySpid = placeholderRead.placeholderKeyBySpid;
  }

  const txStylesCursor = masterCursor.child(P_NS, "txStyles");
  let textStyles: MasterTextStyles | undefined;
  if (txStylesCursor !== undefined) {
    txStylesCursor.ignoreEmptyChild(P_NS, "titleStyle");
    txStylesCursor.ignoreEmptyChild(P_NS, "bodyStyle");
    txStylesCursor.ignoreEmptyChild(P_NS, "otherStyle");
    textStyles = readMasterTextStyles(txStylesCursor, masterClrMap);
  }
  const masterHeaderFooter = readHeaderFooterElement(
    masterCursor,
    masterPartPath,
    diagnostics,
  );

  const masterRels = pkg.readRelationships(masterPartPath);
  diagnostics.push(...masterRels.diagnostics);
  const masterRelById = new Map<string, ParsedRelationshipEntry>();
  for (const rel of masterRels.entries) {
    masterRelById.set(rel.Id, rel);
  }

  let masterTheme: Theme | undefined;
  const themeRel = masterRels.entries.find(
    (entry) => entry.Type === REL_THEME && entry.TargetMode !== "External",
  );
  if (themeRel !== undefined) {
    const themePartPath = resolvePartPath(masterPartPath, themeRel.Target);
    if (themePartPath === undefined || !pkg.hasPart(themePartPath)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: masterPartPath, xpath: themeRel.xpath },
        relationshipId: themeRel.Id,
        targetPart: themePartPath ?? themeRel.Target,
      });
    } else {
      masterTheme = readThemePart(
        pkg,
        themePartPath,
        themesByPartPath,
        diagnostics,
      );
    }
  }

  const masterId = allocateScopedAuthorId(
    masterName,
    usedDeckIds,
    `master-${masterOrder}`,
  );
  const masterProps: SlideMasterProps = {
    id: masterId,
    ...(masterName === undefined ? {} : { name: masterName }),
    ...(preserve === undefined ? {} : { preserve }),
    ...(masterPlaceholders === undefined
      ? {}
      : { placeholders: masterPlaceholders }),
    ...(textStyles === undefined ? {} : { textStyles }),
    ...(masterTheme === undefined ? {} : { theme: masterTheme }),
    ...(colorMap === undefined ? {} : { colorMap }),
    ...(masterHeaderFooter === undefined
      ? {}
      : { headerFooter: masterHeaderFooter }),
  };
  const masterNode = createSlideMasterNode(masterProps);
  if (spTree !== undefined && shapeTreeHandlers.length > 0) {
    readShapeTreeItems({
      spTree,
      pkg,
      partPath: masterPartPath,
      slideIndex: undefined,
      relIndex: {
        entries: masterRels.entries,
        byId: masterRelById,
        diagnostics: [],
      },
      shapeTreeHandlers,
      registerDeferredSlideLink,
      appendNode: (node) => masterNode.append(node),
      diagnostics,
      containerNode: masterNode,
      clrMap: masterClrMap,
    });
  }
  for (const slideHandler of masterLevelSlideHandlers) {
    if (
      slideHandler.matches !== undefined &&
      !slideHandler.matches(masterCursor)
    ) {
      continue;
    }
    const outcome = slideHandler.read({
      slideCursor: masterCursor,
      slideNode: masterNode,
      part: masterPartPath,
      ...(slideHandler.slot === "transition" ? {} : { slideIndex: 0 }),
      package: pkg,
      ...(masterPlaceholderKeyBySpid === undefined
        ? {}
        : { placeholderKeyBySpid: masterPlaceholderKeyBySpid }),
    });
    diagnostics.push(...stripSlideIndexFromDiagnostics(outcome.diagnostics));
  }

  const usedLayoutIds = new Set<string>();
  const sldLayoutIdLst = masterCursor.child(P_NS, "sldLayoutIdLst");
  if (sldLayoutIdLst !== undefined) {
    const layoutIdEntries = sldLayoutIdLst.children(P_NS, "sldLayoutId");
    for (let index = 0; index < layoutIdEntries.length; index += 1) {
      const layoutIdCursor = layoutIdEntries[index]!;
      validateMasterLayoutNumericId(
        layoutIdCursor.attr("id"),
        masterPartPath,
        layoutIdCursor.xpath,
        diagnostics,
      );
      const rId = layoutIdCursor.attr("id", R_NS);
      if (rId === undefined) {
        diagnostics.push({
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: masterPartPath, xpath: layoutIdCursor.xpath },
          relationshipId: "",
        });
        continue;
      }
      const rel = masterRelById.get(rId);
      if (
        rel === undefined ||
        rel.Type !== REL_SLIDE_LAYOUT ||
        rel.TargetMode === "External"
      ) {
        diagnostics.push({
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: masterPartPath, xpath: layoutIdCursor.xpath },
          relationshipId: rId,
        });
        continue;
      }
      const layoutPart = resolvePartPath(masterPartPath, rel.Target);
      if (layoutPart === undefined || !pkg.hasPart(layoutPart)) {
        diagnostics.push({
          code: "missing-target-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: masterPartPath, xpath: layoutIdCursor.xpath },
          relationshipId: rId,
          targetPart: layoutPart ?? rel.Target,
        });
        continue;
      }
      const layoutNode = readSlideLayoutPart(
        pkg,
        layoutPart,
        masterPartPath,
        index + 1,
        usedLayoutIds,
        masterClrMap,
        masterLevelSlideHandlers,
        shapeTreeHandlers,
        registerDeferredSlideLink,
        diagnostics,
      );
      if (layoutNode !== undefined) {
        masterNode.append(layoutNode);
        layoutNodeByPartPath.set(layoutPart, layoutNode);
        clrMapByLayoutPartPath.set(layoutPart, masterClrMap);
      }
    }
  }

  appendUnconsumedDiagnostics(masterCursor, undefined, diagnostics);
  return masterNode;
}

/**
 * Traverses `<p:sldMasterIdLst>` and `<p:sldLayoutIdLst>` relationships to reconstruct explicit
 * `SlideMaster` and `SlideLayout` ownership (including idle layouts) onto `deckNode`.
 */
export function importMastersAndLayouts(
  presCursor: XmlElementCursor,
  pkg: PptxPackage,
  presRelById: ReadonlyMap<string, ParsedRelationshipEntry>,
  deckNode: DeckContainer,
  usedDeckIds: Set<string>,
  diagnostics: ImportDiagnostic[],
  slideHandlers: readonly SlideImportHandler[] = [],
  shapeTreeHandlers: readonly ShapeTreeImportHandler[] = [],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void = () => {},
  clrMapByLayoutPartPath: Map<
    string,
    Readonly<Record<string, NativeColorSlot>>
  > = new Map(),
): ReadonlyMap<string, SlideLayoutContainer> {
  const masterLevelSlideHandlers = slideHandlers.filter(
    (handler) =>
      handler.slot === "background" ||
      handler.slot === "header-footer" ||
      handler.slot === "transition" ||
      handler.slot === "timing",
  );
  const layoutNodeByPartPath = new Map<string, SlideLayoutContainer>();
  const themesByPartPath = new Map<string, Theme>();
  const masterIdLst = presCursor.child(P_NS, "sldMasterIdLst");
  if (masterIdLst === undefined) {
    return layoutNodeByPartPath;
  }
  const masterIdEntries = masterIdLst.children(P_NS, "sldMasterId");
  for (let index = 0; index < masterIdEntries.length; index += 1) {
    const masterIdCursor = masterIdEntries[index]!;
    validateMasterLayoutNumericId(
      masterIdCursor.attr("id"),
      pkg.presentationPath,
      masterIdCursor.xpath,
      diagnostics,
    );
    const rId = masterIdCursor.attr("id", R_NS);
    if (rId === undefined) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: masterIdCursor.xpath },
        relationshipId: "",
      });
      continue;
    }
    const rel = presRelById.get(rId);
    if (
      rel === undefined ||
      rel.Type !== REL_SLIDE_MASTER ||
      rel.TargetMode === "External"
    ) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: masterIdCursor.xpath },
        relationshipId: rId,
      });
      continue;
    }
    const masterPart = resolvePartPath(pkg.presentationPath, rel.Target);
    if (masterPart === undefined || !pkg.hasPart(masterPart)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: masterIdCursor.xpath },
        relationshipId: rId,
        targetPart: masterPart ?? rel.Target,
      });
      continue;
    }
    const masterNode = readSlideMasterPart(
      pkg,
      masterPart,
      index + 1,
      usedDeckIds,
      themesByPartPath,
      layoutNodeByPartPath,
      clrMapByLayoutPartPath,
      masterLevelSlideHandlers,
      shapeTreeHandlers,
      registerDeferredSlideLink,
      diagnostics,
    );
    if (masterNode !== undefined) {
      deckNode.append(masterNode);
    }
  }
  return layoutNodeByPartPath;
}
