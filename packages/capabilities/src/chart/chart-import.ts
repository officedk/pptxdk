import {
  bcp47LangSchema,
  type ColorValue,
  type LinkTarget,
  type NativeColorSlot,
  type Shadow,
} from "@pptxdk/core/capability";
import {
  A_NS,
  C_NS,
  ExtentEmuFragmentSchema,
  HexRgb6Schema,
  MC_NS,
  OffsetEmuFragmentSchema,
  OoxmlBoolSchema,
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  P_NS,
  REL_CHART,
  REL_HYPERLINK,
  R_NS,
  TransformAttributesFragmentSchema,
  XmlElementCursor,
  emuToPx,
  parseXmlPart,
  resolvePartPath,
  type ImportDiagnostic,
  type PptxPackage,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
} from "@pptxdk/import";
import { z } from "zod";

import { createChartNode } from "./chart-declaration.ts";
import {
  chartAxisPositionFromNative,
  chartDataLabelPositionFromNative,
  chartErrorBarValueTypeFromNative,
  chartLegendPositionFromNative,
  chartLineDashFromNative,
  chartTrendlineTypeFromNative,
} from "./chart-labels.ts";
import type {
  ChartAxes,
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartCustomErrorBars,
  ChartDataLabelPosition,
  ChartDataLabelSettings,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartErrorBarInclude,
  ChartErrorBars,
  ChartGroup,
  ChartKind,
  ChartLine,
  ChartLineDash,
  ChartMarker,
  ChartPoint,
  ChartProps,
  ChartSeries,
  ChartTextStyle,
  ChartTickLabelPosition,
  ChartTickMark,
  ChartTitle,
  ChartTrendline,
  IndependentCategorySeries,
  ScatterChartPoint,
  ScatterChartSeries,
  ScatterDataLabels,
  ScatterErrorBars,
  ScatterLine,
} from "./chart-schemas.ts";

const C14_NS = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
const C15_NS = "http://schemas.microsoft.com/office/drawing/2012/chart";
const WPS_ET_NS = "https://web.wps.cn/et/2018/main";

const P14_MOD_ID_EXT_URI = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}";
const C15_CAT_RANGE_EXT_URI = "{02D57815-91ED-43cb-92C2-25804820EDAC}";
const C15_DLBL_RANGE_FLAG_EXT_URI = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}";
const C16_UNIQUE_ID_EXT_URI = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}";
const WPS_CHART_PROPS_EXT_URI = "{0b15fc19-7d7d-44ad-8c2d-2c3a37ce22c3}";

const NonVisualIdSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const GraphicDataAttrSchema = z.object({
  uri: z.string().min(1),
});

const OoxmlFloatSchema = z
  .string()
  .transform((value) => Number(value.trim()))
  .refine((value) => Number.isFinite(value));

const SchemeColorSchema = z.object({
  val: z.enum([
    "tx1",
    "bg1",
    "tx2",
    "bg2",
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
  ]),
});

const schemeToNativeSlot: Readonly<
  Record<z.infer<typeof SchemeColorSchema>["val"], NativeColorSlot>
> = {
  tx1: "dk1",
  bg1: "lt1",
  tx2: "dk2",
  bg2: "lt2",
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

const BarDirSchema = z.object({ val: z.enum(["col", "bar"]) });
const GroupingSchema = z.object({
  val: z.enum(["clustered", "stacked", "percentStacked", "standard"]),
});
const ScatterStyleSchema = z.object({
  val: z.enum([
    "none",
    "line",
    "lineMarker",
    "marker",
    "smooth",
    "smoothMarker",
  ]),
});
const CrossesSchema = z.object({ val: z.enum(["autoZero", "max", "min"]) });
const TickMarkSchema = z.object({
  val: z.enum(["none", "in", "out", "cross"]),
});
const TickLblPosSchema = z.object({
  val: z.enum(["nextTo", "high", "low", "none"]),
});
const DispBlanksAsSchema = z.object({
  val: z.enum(["gap", "zero", "span"]),
});
const DLblPosSchema = z.object({
  val: z.enum([
    "ctr",
    "inEnd",
    "inBase",
    "outEnd",
    "bestFit",
    "l",
    "r",
    "b",
    "t",
  ]),
});
const MarkerSymbolSchema = z.object({
  val: z.enum([
    "auto",
    "circle",
    "dash",
    "diamond",
    "dot",
    "none",
    "plus",
    "square",
    "star",
    "triangle",
    "x",
  ]),
});
const ErrBarTypeSchema = z.object({
  val: z.enum(["both", "plus", "minus"]),
});

type XmlElement = XmlElementCursor["element"];

function directElementChild(
  parent: XmlElement,
  namespaceUri: string,
  localName: string,
): XmlElement | undefined {
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if (
      (el.namespaceURI ?? "") === namespaceUri &&
      (el.localName ?? el.tagName) === localName
    ) {
      return el;
    }
  }
  return undefined;
}

/** Side-effect-free predicate matching `<p:graphicFrame>` elements containing DrawingML charts. */
export function isChartGraphicFrame(cursor: XmlElementCursor): boolean {
  const graphic = directElementChild(cursor.element, A_NS, "graphic");
  if (graphic === undefined) return false;
  const graphicData = directElementChild(graphic, A_NS, "graphicData");
  if (graphicData === undefined) return false;
  return graphicData.getAttribute("uri") === C_NS;
}

function consumeAllAttributes(cursor: XmlElementCursor): void {
  const attrs = cursor.element.attributes;
  for (let i = 0; i < attrs.length; i += 1) {
    const attr = attrs.item(i);
    if (attr === null) continue;
    const ns = attr.namespaceURI ?? "";
    const local = attr.localName ?? attr.name;
    if (
      ns === "http://www.w3.org/2000/xmlns/" ||
      attr.name === "xmlns" ||
      attr.prefix === "xmlns"
    ) {
      continue;
    }
    cursor.consumeAttr(local, ns);
  }
}

/** Recursively consumes a known benign Office/WPS boilerplate element and its descendants. */
function consumeSubtree(cursor: XmlElementCursor): void {
  consumeAllAttributes(cursor);
  for (const child of cursor.remainingElementChildren()) {
    consumeSubtree(child);
  }
}

function readValAttr(cursor: XmlElementCursor | undefined): string | undefined {
  return cursor?.attr("val");
}

function readBoolVal(
  cursor: XmlElementCursor | undefined,
): boolean | undefined {
  const raw = readValAttr(cursor);
  if (raw === undefined) return undefined;
  const parsed = OoxmlBoolSchema.safeParse(raw);
  return parsed.success ? parsed.data : undefined;
}

function readIntVal(cursor: XmlElementCursor | undefined): number | undefined {
  const raw = readValAttr(cursor);
  if (raw === undefined) return undefined;
  const parsed = OoxmlIntSchema.safeParse(raw);
  return parsed.success ? parsed.data : undefined;
}

function readFloatVal(
  cursor: XmlElementCursor | undefined,
): number | undefined {
  const raw = readValAttr(cursor);
  if (raw === undefined) return undefined;
  const parsed = OoxmlFloatSchema.safeParse(raw);
  return parsed.success ? parsed.data : undefined;
}

function readColorModifier(
  colorCursor: XmlElementCursor,
  localName: "alpha" | "lumMod" | "lumOff",
): number | undefined {
  const mod = colorCursor.child(A_NS, localName);
  if (mod === undefined) return undefined;
  const val = readIntVal(mod);
  if (val === undefined || val < 0 || val > 100_000) return undefined;
  return val / 100_000;
}

function readSolidFillColor(
  fillCursor: XmlElementCursor,
): ColorValue | undefined {
  const srgb = fillCursor.child(A_NS, "srgbClr");
  const scheme =
    srgb === undefined ? fillCursor.child(A_NS, "schemeClr") : undefined;
  const colorCursor = srgb ?? scheme;
  if (colorCursor === undefined) {
    consumeSubtree(fillCursor);
    return undefined;
  }

  let baseString: string | undefined;
  let schemeSlot: NativeColorSlot | undefined;
  if (srgb !== undefined) {
    const raw = srgb.attrsObject(["val"]);
    const parsed = HexRgb6Schema.safeParse(raw["val"]);
    if (parsed.success) {
      baseString = `#${parsed.data.toUpperCase()}`;
    }
  } else if (scheme !== undefined) {
    const raw = scheme.attrsObject(["val"]);
    const parsed = SchemeColorSchema.safeParse(raw);
    if (parsed.success) {
      schemeSlot = schemeToNativeSlot[parsed.data.val];
    }
  }
  const lumMod = readColorModifier(colorCursor, "lumMod");
  const lumOff = readColorModifier(colorCursor, "lumOff");
  const alpha = readColorModifier(colorCursor, "alpha");
  consumeSubtree(colorCursor);

  const baseColor: ColorValue | undefined =
    schemeSlot !== undefined && lumMod !== undefined
      ? {
          slot: schemeSlot,
          lumMod: Math.min(1, lumMod),
          ...(lumOff === undefined ? {} : { lumOff: Math.min(1, lumOff) }),
        }
      : (schemeSlot ?? baseString);

  if (baseColor === undefined) return undefined;
  return alpha === undefined ? baseColor : { color: baseColor, alpha };
}

function readLineFromSpPr(spPr: XmlElementCursor): {
  readonly noLine: boolean;
  readonly line?: ChartLine;
} {
  const ln = spPr.child(A_NS, "ln");
  if (ln === undefined) return { noLine: false };
  const rawW = ln.attr("w");
  ln.consumeAttr("cap");
  const noFill = ln.child(A_NS, "noFill");
  if (noFill !== undefined) {
    consumeSubtree(ln);
    return { noLine: true };
  }
  const solidFill = ln.child(A_NS, "solidFill");
  const color =
    solidFill === undefined ? undefined : readSolidFillColor(solidFill);
  const prstDash = ln.child(A_NS, "prstDash");
  let dash: ChartLineDash | undefined;
  if (prstDash !== undefined) {
    const rawDash = prstDash.attrsObject(["val"])["val"];
    if (rawDash !== undefined) {
      dash = chartLineDashFromNative[rawDash];
    }
  }
  consumeSubtree(ln);
  let width: number | undefined;
  if (rawW !== undefined) {
    const parsedW = OoxmlIntSchema.safeParse(rawW);
    if (parsedW.success && parsedW.data > 0) {
      width = emuToPx(parsedW.data);
    }
  }
  if (color === undefined && width === undefined && dash === undefined) {
    return { noLine: false };
  }
  return {
    noLine: false,
    line: {
      ...(color === undefined ? {} : { color }),
      ...(width === undefined ? {} : { width }),
      ...(dash === undefined ? {} : { dash }),
    },
  };
}

function readSeriesSpPr(
  spPr: XmlElementCursor,
  ignoreLine: boolean,
): { readonly color?: ColorValue; readonly line?: ChartLine } {
  const solidFill = spPr.child(A_NS, "solidFill");
  const color =
    solidFill === undefined ? undefined : readSolidFillColor(solidFill);
  const { line } = ignoreLine ? { line: undefined } : readLineFromSpPr(spPr);
  consumeSubtree(spPr);
  return {
    ...(color === undefined ? {} : { color }),
    ...(line === undefined ? {} : { line }),
  };
}

function resolveChartTypefaceToken(
  typeface: string,
): NonNullable<ChartTextStyle["fontFamily"]> {
  if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
    return "minor";
  }
  if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
    return "major";
  }
  return typeface;
}

function readChartOuterShadow(rPrCursor: XmlElementCursor): Shadow | undefined {
  const effectLst = rPrCursor.child(A_NS, "effectLst");
  if (effectLst === undefined) return undefined;
  const outerShdw = effectLst.child(A_NS, "outerShdw");
  if (outerShdw === undefined) {
    consumeSubtree(effectLst);
    return undefined;
  }
  const blurRadRaw = outerShdw.attr("blurRad");
  const distRaw = outerShdw.attr("dist");
  const dirRaw = outerShdw.attr("dir");
  outerShdw.consumeAttr("algn");
  outerShdw.consumeAttr("rotWithShape");
  const color = readSolidFillColor(outerShdw);
  consumeSubtree(outerShdw);
  consumeSubtree(effectLst);

  const blurRad =
    blurRadRaw === undefined
      ? undefined
      : OoxmlNonNegativeIntSchema.safeParse(blurRadRaw);
  const dist =
    distRaw === undefined
      ? undefined
      : OoxmlNonNegativeIntSchema.safeParse(distRaw);
  const dir =
    dirRaw === undefined ? undefined : OoxmlIntSchema.safeParse(dirRaw);
  return {
    ...(color === undefined ? {} : { color }),
    ...(blurRad?.success ? { blur: emuToPx(blurRad.data) } : {}),
    ...(dist?.success ? { distance: emuToPx(dist.data) } : {}),
    ...(dir?.success ? { angle: dir.data / 60_000 } : {}),
  };
}

function readChartRunProperties(rPrCursor: XmlElementCursor): ChartTextStyle {
  const rawLang = rPrCursor.attr("lang");
  const parsedLang =
    rawLang === undefined ? undefined : bcp47LangSchema.safeParse(rawLang);
  const lang = parsedLang?.success ? parsedLang.data : undefined;

  const szRaw = rPrCursor.attr("sz");
  const szParsed =
    szRaw === undefined ? undefined : OoxmlIntSchema.safeParse(szRaw);
  const fontSize: ChartTextStyle["fontSize"] =
    szParsed?.success && szParsed.data >= 100 && szParsed.data <= 400_000
      ? `${szParsed.data / 100}pt`
      : undefined;

  const bRaw = rPrCursor.attr("b");
  const bParsed =
    bRaw === undefined ? undefined : OoxmlBoolSchema.safeParse(bRaw);
  const fontWeight: ChartTextStyle["fontWeight"] = bParsed?.success
    ? bParsed.data
      ? "bold"
      : "normal"
    : undefined;

  const iRaw = rPrCursor.attr("i");
  const iParsed =
    iRaw === undefined ? undefined : OoxmlBoolSchema.safeParse(iRaw);
  const fontStyle: ChartTextStyle["fontStyle"] = iParsed?.success
    ? iParsed.data
      ? "italic"
      : "normal"
    : undefined;

  const uRaw = rPrCursor.attr("u");
  const strikeRaw = rPrCursor.attr("strike");
  let textDecoration: ChartTextStyle["textDecoration"];
  if (uRaw === "sng") textDecoration = "underline";
  else if (strikeRaw === "sngStrike") textDecoration = "line-through";
  else if (uRaw === "none" && strikeRaw === "noStrike") textDecoration = "none";

  const capRaw = rPrCursor.attr("cap");
  let textTransform: ChartTextStyle["textTransform"];
  if (capRaw === "all") textTransform = "uppercase";
  else if (capRaw === "none") textTransform = "none";

  const spcRaw = rPrCursor.attr("spc");
  const spcParsed =
    spcRaw === undefined ? undefined : OoxmlIntSchema.safeParse(spcRaw);
  const letterSpacing =
    spcParsed?.success &&
    spcParsed.data >= -400_000 &&
    spcParsed.data <= 400_000
      ? spcParsed.data / 75
      : undefined;

  const solidFill = rPrCursor.child(A_NS, "solidFill");
  const color =
    solidFill === undefined ? undefined : readSolidFillColor(solidFill);
  if (solidFill !== undefined) consumeSubtree(solidFill);
  const textShadow = readChartOuterShadow(rPrCursor);

  const latin = rPrCursor.child(A_NS, "latin");
  const ea = rPrCursor.child(A_NS, "ea");
  const cs = rPrCursor.child(A_NS, "cs");
  const latinTypeface = latin?.attr("typeface");
  const eaTypeface = ea?.attr("typeface");
  const csTypeface = cs?.attr("typeface");
  if (latin !== undefined) consumeSubtree(latin);
  if (ea !== undefined) consumeSubtree(ea);
  if (cs !== undefined) consumeSubtree(cs);
  const typeface = latinTypeface ?? eaTypeface ?? csTypeface;
  let fontFamily: ChartTextStyle["fontFamily"];
  if (typeface !== undefined && typeface.trim() !== "") {
    fontFamily = resolveChartTypefaceToken(typeface);
  }
  consumeSubtree(rPrCursor);

  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(fontWeight === undefined ? {} : { fontWeight }),
    ...(fontStyle === undefined ? {} : { fontStyle }),
    ...(color === undefined ? {} : { color }),
    ...(letterSpacing === undefined ? {} : { letterSpacing }),
    ...(textTransform === undefined ? {} : { textTransform }),
    ...(textDecoration === undefined ? {} : { textDecoration }),
    ...(textShadow === undefined ? {} : { textShadow }),
    ...(lang === undefined ? {} : { lang }),
  };
}

function readChartTextProperties(
  txPrCursor: XmlElementCursor,
): ChartTextStyle | undefined {
  let verticalAlign: ChartTextStyle["verticalAlign"];
  const bodyPr = txPrCursor.child(A_NS, "bodyPr");
  if (bodyPr !== undefined) {
    const anchor = bodyPr.attr("anchor");
    if (anchor === "t") verticalAlign = "top";
    else if (anchor === "ctr") verticalAlign = "middle";
    else if (anchor === "b") verticalAlign = "bottom";
    consumeSubtree(bodyPr);
  }

  let align: ChartTextStyle["align"];
  let lineHeight: ChartTextStyle["lineHeight"];
  let spacingAfter: ChartTextStyle["spacingAfter"];
  let defStyle: ChartTextStyle = {};
  let runStyle: ChartTextStyle = {};
  let endLang: string | undefined;

  const p = txPrCursor.child(A_NS, "p");
  if (p !== undefined) {
    const pPr = p.child(A_NS, "pPr");
    if (pPr !== undefined) {
      const algn = pPr.attr("algn");
      if (algn === "l") align = "left";
      else if (algn === "ctr") align = "center";
      else if (algn === "r") align = "right";
      else if (algn === "just") align = "justify";

      const lnSpc = pPr.child(A_NS, "lnSpc");
      if (lnSpc !== undefined) {
        const spcPts = readIntVal(lnSpc.child(A_NS, "spcPts"));
        if (spcPts !== undefined && spcPts > 0 && spcPts <= 158_400) {
          lineHeight = `${spcPts / 100}pt`;
        }
        consumeSubtree(lnSpc);
      }

      const spcAft = pPr.child(A_NS, "spcAft");
      if (spcAft !== undefined) {
        const spcPts = readIntVal(spcAft.child(A_NS, "spcPts"));
        if (spcPts !== undefined && spcPts >= 0 && spcPts <= 158_400) {
          spacingAfter = spcPts / 75;
        }
        consumeSubtree(spcAft);
      }

      const defRPr = pPr.child(A_NS, "defRPr");
      if (defRPr !== undefined) {
        defStyle = readChartRunProperties(defRPr);
      }
      consumeSubtree(pPr);
    }

    const r = p.child(A_NS, "r");
    if (r !== undefined) {
      const rPr = r.child(A_NS, "rPr");
      if (rPr !== undefined) {
        runStyle = readChartRunProperties(rPr);
      }
      consumeSubtree(r);
    }

    const endParaRPr = p.child(A_NS, "endParaRPr");
    if (endParaRPr !== undefined) {
      const rawLang = endParaRPr.attr("lang");
      const parsedLang =
        rawLang === undefined ? undefined : bcp47LangSchema.safeParse(rawLang);
      if (parsedLang?.success) endLang = parsedLang.data;
      consumeSubtree(endParaRPr);
    }
    consumeSubtree(p);
  }
  consumeSubtree(txPrCursor);

  const merged: ChartTextStyle = {
    ...defStyle,
    ...runStyle,
    ...(endLang !== undefined &&
    defStyle.lang === undefined &&
    runStyle.lang === undefined
      ? { lang: endLang }
      : {}),
    ...(verticalAlign === undefined ? {} : { verticalAlign }),
    ...(align === undefined ? {} : { align }),
    ...(lineHeight === undefined ? {} : { lineHeight }),
    ...(spacingAfter === undefined ? {} : { spacingAfter }),
  };
  return Object.keys(merged).length === 0 ? undefined : merged;
}

function fontSizePointsValue(
  value: ChartTextStyle["fontSize"],
): number | undefined {
  if (typeof value === "number") return value * 0.75;
  if (typeof value === "string") {
    if (value.endsWith("pt")) return Number(value.slice(0, -2));
    if (value.endsWith("px")) return Number(value.slice(0, -2)) * 0.75;
  }
  return undefined;
}

function stableJsonEqual(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (
    typeof left !== "object" ||
    typeof right !== "object" ||
    left === null ||
    right === null
  ) {
    return false;
  }
  const leftKeys = Object.keys(left as Record<string, unknown>).filter(
    (k) => (left as Record<string, unknown>)[k] !== undefined,
  );
  const rightKeys = Object.keys(right as Record<string, unknown>).filter(
    (k) => (right as Record<string, unknown>)[k] !== undefined,
  );
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every((key) =>
    stableJsonEqual(
      (left as Record<string, unknown>)[key],
      (right as Record<string, unknown>)[key],
    ),
  );
}

function styleFieldEqual(
  key: keyof ChartTextStyle,
  left: unknown,
  right: unknown,
): boolean {
  if (left === right) return true;
  if (key === "fontSize" || key === "lineHeight") {
    const lPt = fontSizePointsValue(left as ChartTextStyle["fontSize"]);
    const rPt = fontSizePointsValue(right as ChartTextStyle["fontSize"]);
    return lPt !== undefined && rPt !== undefined && Math.abs(lPt - rPt) < 1e-6;
  }
  return stableJsonEqual(left, right);
}

function diffChartTextStyle(
  child: ChartTextStyle | undefined,
  inherited: ChartTextStyle | undefined,
): ChartTextStyle | undefined {
  if (child === undefined) return undefined;
  const baseline: ChartTextStyle = {
    fontWeight: "normal",
    fontStyle: "normal",
    verticalAlign: "top",
    textDecoration: "none",
    textTransform: "none",
    color: "dk1",
    ...inherited,
  };
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(child)) {
    if (value === undefined) continue;
    const baseValue = (baseline as Record<string, unknown>)[key];
    if (!styleFieldEqual(key as keyof ChartTextStyle, value, baseValue)) {
      out[key] = value;
    }
  }
  return Object.keys(out).length === 0 ? undefined : out;
}

function normalizeRootChartStyle(
  rawStyle: ChartTextStyle | undefined,
): ChartTextStyle | undefined {
  if (rawStyle === undefined) return undefined;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(rawStyle)) {
    if (value === undefined) continue;
    if (key === "fontWeight" && value === "normal") continue;
    if (key === "fontStyle" && value === "normal") continue;
    if (key === "verticalAlign" && value === "top") continue;
    if (key === "textDecoration" && value === "none") continue;
    if (key === "textTransform" && value === "none") continue;
    out[key] = value;
  }
  return Object.keys(out).length === 0 ? undefined : out;
}

function readChartTitle(
  titleCursor: XmlElementCursor,
  inheritedStyle: ChartTextStyle | undefined,
): ChartTitle | undefined {
  let titleText: string | undefined;
  let richStyle: ChartTextStyle | undefined;
  const tx = titleCursor.child(C_NS, "tx");
  if (tx !== undefined) {
    const rich = tx.child(C_NS, "rich");
    if (rich !== undefined) {
      const bodyPr = rich.child(A_NS, "bodyPr");
      const anchor = bodyPr?.attr("anchor");
      const verticalAlign =
        anchor === "t"
          ? ("top" as const)
          : anchor === "ctr"
            ? ("middle" as const)
            : anchor === "b"
              ? ("bottom" as const)
              : undefined;
      if (bodyPr !== undefined) consumeSubtree(bodyPr);

      let align: ChartTextStyle["align"];
      let lineHeight: ChartTextStyle["lineHeight"];
      let spacingAfter: ChartTextStyle["spacingAfter"];
      let defRun: ChartTextStyle = {};
      let explicitRun: ChartTextStyle = {};
      const parts: string[] = [];

      for (const [paraIndex, para] of rich.children(A_NS, "p").entries()) {
        if (paraIndex === 0) {
          const pPr = para.child(A_NS, "pPr");
          if (pPr !== undefined) {
            const algn = pPr.attr("algn");
            align =
              algn === "l"
                ? ("left" as const)
                : algn === "ctr"
                  ? ("center" as const)
                  : algn === "r"
                    ? ("right" as const)
                    : algn === "just"
                      ? ("justify" as const)
                      : undefined;
            const lnSpc = pPr.child(A_NS, "lnSpc");
            if (lnSpc !== undefined) {
              const spcPts = readIntVal(lnSpc.child(A_NS, "spcPts"));
              if (spcPts !== undefined && spcPts > 0 && spcPts <= 158_400) {
                lineHeight = `${spcPts / 100}pt`;
              }
              consumeSubtree(lnSpc);
            }
            const spcAft = pPr.child(A_NS, "spcAft");
            if (spcAft !== undefined) {
              const spcPts = readIntVal(spcAft.child(A_NS, "spcPts"));
              if (spcPts !== undefined && spcPts >= 0 && spcPts <= 158_400) {
                spacingAfter = spcPts / 75;
              }
              consumeSubtree(spcAft);
            }
            const defRPr = pPr.child(A_NS, "defRPr");
            if (defRPr !== undefined) {
              defRun = readChartRunProperties(defRPr);
            }
            consumeSubtree(pPr);
          }
        }

        const runs: string[] = [];
        for (const child of para.remainingElementChildren()) {
          if (child.namespaceUri === A_NS && child.localName === "r") {
            if (paraIndex === 0 && Object.keys(explicitRun).length === 0) {
              const rPr = child.child(A_NS, "rPr");
              if (rPr !== undefined) {
                explicitRun = readChartRunProperties(rPr);
              }
            }
            const t = child.child(A_NS, "t");
            if (t !== undefined) runs.push(t.textContent);
            consumeSubtree(child);
          } else if (child.namespaceUri === A_NS && child.localName === "br") {
            runs.push("\n");
            consumeSubtree(child);
          } else {
            consumeSubtree(child);
          }
        }
        consumeSubtree(para);
        parts.push(runs.join(""));
      }

      titleText = parts.join("\n");
      consumeSubtree(rich);
      richStyle = {
        ...defRun,
        ...explicitRun,
        ...(verticalAlign === undefined ? {} : { verticalAlign }),
        ...(align === undefined ? {} : { align }),
        ...(lineHeight === undefined ? {} : { lineHeight }),
        ...(spacingAfter === undefined ? {} : { spacingAfter }),
      };
    } else {
      const strRef = tx.child(C_NS, "strRef");
      if (strRef !== undefined) {
        const strCache = strRef.child(C_NS, "strCache");
        if (strCache !== undefined) {
          const pts = strCache.children(C_NS, "pt");
          const first = pts[0];
          const v = first?.child(C_NS, "v");
          if (v !== undefined) titleText = v.textContent;
        }
        consumeSubtree(strRef);
      } else {
        const v = tx.child(C_NS, "v");
        if (v !== undefined) titleText = v.textContent;
      }
    }
    consumeSubtree(tx);
  }
  const txPr = titleCursor.child(C_NS, "txPr");
  const txPrStyle =
    txPr === undefined ? undefined : readChartTextProperties(txPr);
  consumeSubtree(titleCursor);
  if (titleText === undefined || titleText.trim() === "") return undefined;

  const combinedStyle =
    richStyle !== undefined || txPrStyle !== undefined
      ? { ...txPrStyle, ...richStyle }
      : undefined;
  const styleDiff = diffChartTextStyle(combinedStyle, inheritedStyle);
  return styleDiff === undefined
    ? titleText
    : { text: titleText, style: styleDiff };
}

type StringDataSourceResult = {
  readonly formula?: string;
  readonly values: readonly string[];
};

function readPtCount(container: XmlElementCursor): number | undefined {
  const ptCountCursor = container.child(C_NS, "ptCount");
  if (ptCountCursor === undefined) return undefined;
  const raw = readValAttr(ptCountCursor);
  if (raw === undefined) return undefined;
  const parsed = OoxmlNonNegativeIntSchema.safeParse(raw);
  return parsed.success ? parsed.data : undefined;
}

function readStringCacheOrLit(container: XmlElementCursor): readonly string[] {
  const ptCount = readPtCount(container) ?? 0;
  const pts = container.children(C_NS, "pt");
  const size = Math.max(ptCount, pts.length);
  if (size === 0) {
    consumeSubtree(container);
    return [];
  }
  const result = Array.from<string>({ length: size }).fill("");
  for (const pt of pts) {
    const rawIdx = pt.attr("idx");
    const parsedIdx =
      rawIdx === undefined
        ? undefined
        : OoxmlNonNegativeIntSchema.safeParse(rawIdx);
    const v = pt.child(C_NS, "v");
    const text = v?.textContent ?? "";
    consumeSubtree(pt);
    if (parsedIdx?.success && parsedIdx.data < result.length) {
      result[parsedIdx.data] = text;
    }
  }
  consumeSubtree(container);
  return result;
}

function readNumericCacheOrLitAsStrings(
  container: XmlElementCursor,
): readonly string[] {
  const formatCode = container.child(C_NS, "formatCode");
  if (formatCode !== undefined) consumeSubtree(formatCode);
  return readStringCacheOrLit(container);
}

function readNumericCacheOrLit(
  container: XmlElementCursor,
): readonly (number | null)[] {
  const formatCode = container.child(C_NS, "formatCode");
  if (formatCode !== undefined) consumeSubtree(formatCode);
  const ptCount = readPtCount(container);
  const pts = container.children(C_NS, "pt");
  if (ptCount === undefined && pts.length === 0) {
    consumeSubtree(container);
    return [];
  }
  const length = ptCount ?? pts.length;
  const result = Array.from<number | null>({ length }).fill(null);
  for (const pt of pts) {
    const rawIdx = pt.attr("idx");
    pt.consumeAttr("formatCode");
    const parsedIdx =
      rawIdx === undefined
        ? undefined
        : OoxmlNonNegativeIntSchema.safeParse(rawIdx);
    const v = pt.child(C_NS, "v");
    const rawVal = v?.textContent.trim();
    consumeSubtree(pt);
    if (
      parsedIdx?.success &&
      parsedIdx.data < result.length &&
      rawVal !== undefined &&
      rawVal !== ""
    ) {
      const numeric = Number(rawVal);
      if (Number.isFinite(numeric)) {
        result[parsedIdx.data] = numeric;
      }
    }
  }
  consumeSubtree(container);
  return result;
}

function readStringDataSource(
  sourceCursor: XmlElementCursor,
): StringDataSourceResult {
  const strRef = sourceCursor.child(C_NS, "strRef");
  if (strRef !== undefined) {
    const f = strRef.child(C_NS, "f");
    const formula = f?.textContent;
    const strCache = strRef.child(C_NS, "strCache");
    const values = strCache === undefined ? [] : readStringCacheOrLit(strCache);
    consumeSubtree(strRef);
    return { ...(formula === undefined ? {} : { formula }), values };
  }
  const strLit = sourceCursor.child(C_NS, "strLit");
  if (strLit !== undefined) {
    return { values: readStringCacheOrLit(strLit) };
  }
  const numRef = sourceCursor.child(C_NS, "numRef");
  if (numRef !== undefined) {
    const f = numRef.child(C_NS, "f");
    const formula = f?.textContent;
    const numCache = numRef.child(C_NS, "numCache");
    const values =
      numCache === undefined ? [] : readNumericCacheOrLitAsStrings(numCache);
    consumeSubtree(numRef);
    return { ...(formula === undefined ? {} : { formula }), values };
  }
  const numLit = sourceCursor.child(C_NS, "numLit");
  if (numLit !== undefined) {
    return { values: readNumericCacheOrLitAsStrings(numLit) };
  }
  const multiLvlStrRef = sourceCursor.child(C_NS, "multiLvlStrRef");
  if (multiLvlStrRef !== undefined) {
    const f = multiLvlStrRef.child(C_NS, "f");
    const formula = f?.textContent;
    const cache = multiLvlStrRef.child(C_NS, "multiLvlStrCache");
    let values: readonly string[] = [];
    if (cache !== undefined) {
      const ptCount = readPtCount(cache) ?? 0;
      const lvls = cache.children(C_NS, "lvl");
      if (ptCount > 0 && lvls.length > 0) {
        const arr = Array.from<string>({ length: ptCount }).fill("");
        for (const pt of lvls[0]!.children(C_NS, "pt")) {
          const rawIdx = pt.attr("idx");
          const parsedIdx =
            rawIdx === undefined
              ? undefined
              : OoxmlNonNegativeIntSchema.safeParse(rawIdx);
          const v = pt.child(C_NS, "v");
          if (
            parsedIdx?.success &&
            parsedIdx.data < arr.length &&
            v !== undefined
          ) {
            arr[parsedIdx.data] = v.textContent;
          }
          consumeSubtree(pt);
        }
        values = arr;
      }
      consumeSubtree(cache);
    }
    consumeSubtree(multiLvlStrRef);
    return { ...(formula === undefined ? {} : { formula }), values };
  }
  return { values: [] };
}

function readNumericDataSource(
  sourceCursor: XmlElementCursor,
): readonly (number | null)[] {
  const numRef = sourceCursor.child(C_NS, "numRef");
  if (numRef !== undefined) {
    const f = numRef.child(C_NS, "f");
    if (f !== undefined) consumeSubtree(f);
    const numCache = numRef.child(C_NS, "numCache");
    const values =
      numCache === undefined ? [] : readNumericCacheOrLit(numCache);
    consumeSubtree(numRef);
    return values;
  }
  const numLit = sourceCursor.child(C_NS, "numLit");
  if (numLit !== undefined) {
    return readNumericCacheOrLit(numLit);
  }
  return [];
}

function readSeriesName(
  txCursor: XmlElementCursor | undefined,
): string | undefined {
  if (txCursor === undefined) return undefined;
  const strRef = txCursor.child(C_NS, "strRef");
  if (strRef !== undefined) {
    const f = strRef.child(C_NS, "f");
    if (f !== undefined) consumeSubtree(f);
    const strCache = strRef.child(C_NS, "strCache");
    const values = strCache === undefined ? [] : readStringCacheOrLit(strCache);
    consumeSubtree(strRef);
    consumeSubtree(txCursor);
    return values[0];
  }
  const v = txCursor.child(C_NS, "v");
  const text = v?.textContent;
  consumeSubtree(txCursor);
  return text;
}

function readMarkerElement(
  markerCursor: XmlElementCursor,
): boolean | ChartMarker | undefined {
  const symbolCursor = markerCursor.child(C_NS, "symbol");
  const parsedSymbol =
    symbolCursor === undefined
      ? undefined
      : MarkerSymbolSchema.safeParse(symbolCursor.attrsObject(["val"]));
  const symbol = parsedSymbol?.success ? parsedSymbol.data.val : undefined;
  const size = readIntVal(markerCursor.child(C_NS, "size"));
  const spPr = markerCursor.child(C_NS, "spPr");
  let fill: ColorValue | undefined;
  let line: ChartLine | undefined;
  if (spPr !== undefined) {
    const solidFill = spPr.child(A_NS, "solidFill");
    if (solidFill !== undefined) fill = readSolidFillColor(solidFill);
    const lineRes = readLineFromSpPr(spPr);
    line = lineRes.line;
    consumeSubtree(spPr);
  }
  consumeSubtree(markerCursor);
  if (
    symbol === "none" &&
    size === undefined &&
    fill === undefined &&
    line === undefined
  ) {
    return false;
  }
  if (
    symbol === undefined &&
    size === undefined &&
    fill === undefined &&
    line === undefined
  ) {
    return undefined;
  }
  return {
    ...(symbol === undefined ? {} : { symbol: symbol }),
    ...(size !== undefined && size >= 2 && size <= 72 ? { size } : {}),
    ...(fill === undefined ? {} : { fill }),
    ...(line === undefined ? {} : { line }),
  };
}

type ImportedDataLabelsResult = {
  readonly dataLabels?: false | ChartDataLabelSettings;
  readonly dataLabelsStyle?: ChartTextStyle;
  readonly pointLabels: ReadonlyMap<number, false | ChartDataLabelSettings>;
  readonly pointLabelColors: ReadonlyMap<number, ColorValue>;
  readonly showDataLabelsRange?: boolean;
};

function readDataLabelSettingsBody(
  cursor: XmlElementCursor,
  kind: ChartKind,
  isPointLevel = false,
): {
  readonly settings: ChartDataLabelSettings;
  readonly rawStyle?: ChartTextStyle;
  readonly showDataLabelsRange?: boolean;
} {
  const numFmtCursor = cursor.child(C_NS, "numFmt");
  let numberFormat: string | undefined;
  if (numFmtCursor !== undefined) {
    const code = numFmtCursor.attr("formatCode");
    const slRaw = numFmtCursor.attr("sourceLinked");
    if (
      code !== undefined &&
      code.trim() !== "" &&
      (slRaw === "0" || slRaw === "false" || code !== "General")
    ) {
      numberFormat = code;
    }
    consumeSubtree(numFmtCursor);
  }

  const spPr = cursor.child(C_NS, "spPr");
  if (spPr !== undefined) consumeSubtree(spPr);
  const txPr = cursor.child(C_NS, "txPr");
  const rawStyle =
    txPr === undefined ? undefined : readChartTextProperties(txPr);

  const dLblPos = cursor.child(C_NS, "dLblPos");
  let position: ChartDataLabelPosition | undefined;
  if (dLblPos !== undefined) {
    const parsed = DLblPosSchema.safeParse(dLblPos.attrsObject(["val"]));
    if (parsed.success && kind !== "area" && kind !== "doughnut") {
      const mapped = chartDataLabelPositionFromNative[parsed.data.val];
      if (
        mapped !== undefined &&
        (kind !== "scatter" ||
          ["left", "right", "bottom", "top", "center"].includes(mapped))
      ) {
        position = mapped;
      }
    }
  }

  const legendKey = cursor.child(C_NS, "showLegendKey");
  if (legendKey !== undefined) consumeSubtree(legendKey);

  const showVal = readBoolVal(cursor.child(C_NS, "showVal"));
  const showCatName = readBoolVal(cursor.child(C_NS, "showCatName"));
  const showSerName = readBoolVal(cursor.child(C_NS, "showSerName"));
  const showPercent = readBoolVal(cursor.child(C_NS, "showPercent"));
  const showBubbleSize = cursor.child(C_NS, "showBubbleSize");
  if (showBubbleSize !== undefined) consumeSubtree(showBubbleSize);

  const sepCursor = cursor.child(C_NS, "separator");
  const separator =
    sepCursor !== undefined && sepCursor.textContent.length > 0
      ? sepCursor.textContent
      : undefined;

  const showLeaderLines = readBoolVal(cursor.child(C_NS, "showLeaderLines"));
  const leaderLinesCursor = cursor.child(C_NS, "leaderLines");
  let leaderLineStyle: ChartLine | undefined;
  if (leaderLinesCursor !== undefined) {
    const lSpPr = leaderLinesCursor.child(C_NS, "spPr");
    if (lSpPr !== undefined) {
      leaderLineStyle = readLineFromSpPr(lSpPr).line;
      consumeSubtree(lSpPr);
    }
    consumeSubtree(leaderLinesCursor);
  }

  let showDataLabelsRange: boolean | undefined;
  const extLst = cursor.child(C_NS, "extLst");
  if (extLst !== undefined) {
    for (const ext of extLst.children(C_NS, "ext")) {
      const uri = ext.attr("uri");
      if (uri === C15_DLBL_RANGE_FLAG_EXT_URI) {
        const flag = ext.child(C15_NS, "showDataLabelsRange");
        showDataLabelsRange = readBoolVal(flag);
        consumeSubtree(ext);
      } else if (
        uri === "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" ||
        uri === "{0b15fc19-7d7d-44ad-8c2d-2c3a37ce22c3}" ||
        uri === C16_UNIQUE_ID_EXT_URI
      ) {
        consumeSubtree(ext);
      } else {
        // Check for c15 WPS leaderLines/layout boilerplate inside dLbls/extLst
        const hasC15 =
          directElementChild(ext.element, C15_NS, "layout") !== undefined ||
          directElementChild(ext.element, C15_NS, "showLeaderLines") !==
            undefined ||
          directElementChild(ext.element, C15_NS, "leaderLines") !==
            undefined ||
          directElementChild(ext.element, C15_NS, "showDataLabelsRange") !==
            undefined;
        if (hasC15) {
          consumeSubtree(ext);
        }
      }
    }
  }

  // Consume any remaining c15 elements directly inside dLbls if WPS emitted them directly
  for (const rem of cursor.remainingElementChildren()) {
    if (rem.namespaceUri === C15_NS) {
      consumeSubtree(rem);
    }
  }

  const category =
    kind === "scatter"
      ? undefined
      : showDataLabelsRange === true
        ? true
        : showCatName === true
          ? true
          : showCatName === false && showDataLabelsRange === false
            ? false
            : undefined;
  const xValue = kind === "scatter" && showCatName === true ? true : undefined;
  const value = showVal === true ? true : showVal === false ? false : undefined;
  const percent =
    (kind === "pie" || kind === "doughnut") && showPercent === true
      ? true
      : (kind === "pie" || kind === "doughnut") && showPercent === false
        ? false
        : undefined;
  const seriesName =
    showSerName === true ? true : showSerName === false ? false : undefined;

  const hasVisibleContent =
    value === true ||
    percent === true ||
    category === true ||
    xValue === true ||
    seriesName === true;

  let leaderLines: boolean | ChartLine | undefined;
  if (
    !isPointLevel &&
    (kind === "pie" || kind === "doughnut") &&
    hasVisibleContent
  ) {
    if (leaderLineStyle !== undefined) {
      leaderLines = leaderLineStyle;
    } else if (showLeaderLines !== undefined) {
      leaderLines = showLeaderLines;
    }
  }

  const settings: ChartDataLabelSettings = {
    ...(value === undefined ? {} : { value }),
    ...(percent === undefined ? {} : { percent }),
    ...(category === undefined ? {} : { category }),
    ...(xValue === undefined ? {} : { xValue }),
    ...(seriesName === undefined ? {} : { seriesName }),
    ...(separator === undefined ? {} : { separator }),
    ...(position === undefined ? {} : { position }),
    ...(numberFormat === undefined ? {} : { numberFormat }),
    ...(leaderLines === undefined ? {} : { leaderLines }),
  };

  return {
    settings,
    ...(rawStyle === undefined ? {} : { rawStyle }),
    ...(showDataLabelsRange === undefined ? {} : { showDataLabelsRange }),
  };
}

function readDataLabelsElement(
  dLblsCursor: XmlElementCursor,
  kind: ChartKind,
  isGroupLevel: boolean,
  inheritedChartStyle: ChartTextStyle | undefined,
): ImportedDataLabelsResult {
  const pointLabels = new Map<number, false | ChartDataLabelSettings>();
  const pointRawColors = new Map<number, ColorValue>();
  for (const dLbl of dLblsCursor.children(C_NS, "dLbl")) {
    const idx = readIntVal(dLbl.child(C_NS, "idx"));
    const del = readBoolVal(dLbl.child(C_NS, "delete"));
    if (idx !== undefined) {
      if (del === true) {
        pointLabels.set(idx, false);
      } else {
        const { settings, rawStyle } = readDataLabelSettingsBody(
          dLbl,
          kind,
          true,
        );
        pointLabels.set(idx, settings);
        if (rawStyle?.color !== undefined) {
          pointRawColors.set(idx, rawStyle.color);
        }
      }
    }
    consumeSubtree(dLbl);
  }

  const del = readBoolVal(dLblsCursor.child(C_NS, "delete"));
  if (del === true) {
    consumeSubtree(dLblsCursor);
    return {
      dataLabels: false,
      pointLabels,
      pointLabelColors: new Map(),
    };
  }

  const { settings, rawStyle, showDataLabelsRange } = readDataLabelSettingsBody(
    dLblsCursor,
    kind,
    false,
  );
  const dataLabelsStyle = diffChartTextStyle(rawStyle, inheritedChartStyle);

  // Resolve point labelColor relative to the series/group label color (or chart color).
  const inheritedLabelColor =
    rawStyle?.color ?? inheritedChartStyle?.color ?? "dk1";
  const pointLabelColors = new Map<number, ColorValue>();
  for (const [idx, color] of pointRawColors.entries()) {
    if (!stableJsonEqual(color, inheritedLabelColor)) {
      pointLabelColors.set(idx, color);
    }
  }

  // At group level, an all-false boilerplate <c:dLbls> (e.g. PowerPoint/WPS default)
  // represents no authored group dataLabels unless a visible flag or explicit format/position/style is set.
  const hasAnyEnabledOrCustomized =
    settings.value === true ||
    settings.percent === true ||
    settings.category === true ||
    settings.xValue === true ||
    settings.seriesName === true ||
    settings.separator !== undefined ||
    settings.numberFormat !== undefined ||
    settings.position !== undefined ||
    settings.leaderLines !== undefined ||
    dataLabelsStyle !== undefined;

  if (isGroupLevel && !hasAnyEnabledOrCustomized) {
    return { pointLabels, pointLabelColors };
  }

  const normalized: ChartDataLabelSettings = {
    ...(settings.value === true ? { value: true } : {}),
    ...(settings.percent === true ? { percent: true } : {}),
    ...(settings.category === true ? { category: true } : {}),
    ...(settings.xValue === true ? { xValue: true } : {}),
    ...(settings.seriesName === true ? { seriesName: true } : {}),
    ...(settings.separator === undefined
      ? {}
      : { separator: settings.separator }),
    ...(settings.position === undefined ? {} : { position: settings.position }),
    ...(settings.numberFormat === undefined
      ? {}
      : { numberFormat: settings.numberFormat }),
    ...(settings.leaderLines === undefined
      ? {}
      : { leaderLines: settings.leaderLines }),
  };

  return {
    ...(Object.keys(normalized).length > 0 ? { dataLabels: normalized } : {}),
    ...(dataLabelsStyle === undefined ? {} : { dataLabelsStyle }),
    pointLabels,
    pointLabelColors,
    ...(showDataLabelsRange === undefined ? {} : { showDataLabelsRange }),
  };
}

function diffPointDataLabels(
  pointLabels: false | ChartDataLabelSettings,
  seriesLabels: false | ChartDataLabelSettings | undefined,
): false | ChartDataLabelSettings | undefined {
  if (pointLabels === false) return false;
  if (seriesLabels === undefined || seriesLabels === false) {
    return Object.keys(pointLabels).length === 0 ? undefined : pointLabels;
  }
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(pointLabels)) {
    if (value === undefined) continue;
    const baseVal = (seriesLabels as Record<string, unknown>)[key];
    if (!stableJsonEqual(value, baseVal)) {
      out[key] = value;
    }
  }
  return Object.keys(out).length === 0 ? undefined : out;
}

function readSeriesTrendlines(
  serCursor: XmlElementCursor,
  kind: ChartKind,
  stacking: "stacked" | "percent" | undefined,
  pointCount: number,
  fitPoints: readonly { readonly x: number; readonly y: number }[],
): readonly ChartTrendline[] | undefined {
  const nodes = serCursor.children(C_NS, "trendline");
  if (nodes.length === 0) return undefined;
  const supported =
    kind !== "pie" && kind !== "doughnut" && stacking === undefined;
  const result: ChartTrendline[] = [];

  for (const tl of nodes) {
    const nameEl = tl.child(C_NS, "name");
    if (nameEl !== undefined) consumeSubtree(nameEl);

    const spPr = tl.child(C_NS, "spPr");
    const line = spPr === undefined ? undefined : readLineFromSpPr(spPr).line;
    if (spPr !== undefined) consumeSubtree(spPr);

    const typeRaw = readValAttr(tl.child(C_NS, "trendlineType"));
    const type =
      typeRaw === undefined ? undefined : chartTrendlineTypeFromNative[typeRaw];

    const order = readIntVal(tl.child(C_NS, "order"));
    const period = readIntVal(tl.child(C_NS, "period"));
    const rawForward = readFloatVal(tl.child(C_NS, "forward"));
    const rawBackward = readFloatVal(tl.child(C_NS, "backward"));
    const rawIntercept = readFloatVal(tl.child(C_NS, "intercept"));
    const displayRSquared = readBoolVal(tl.child(C_NS, "dispRSqr"));
    const displayEquation = readBoolVal(tl.child(C_NS, "dispEq"));
    consumeSubtree(tl);

    if (!supported || type === undefined) continue;

    const validForecast = (value: number | undefined): number | undefined => {
      if (value === undefined || !Number.isFinite(value) || value <= 0) {
        return undefined;
      }
      if (
        kind !== "scatter" &&
        Math.abs(value * 2 - Math.round(value * 2)) > 1e-9
      ) {
        return undefined;
      }
      return value;
    };
    const forward = validForecast(rawForward);
    const backward = validForecast(rawBackward);
    const baseFields = {
      ...(line === undefined ? {} : { line }),
      ...(forward === undefined ? {} : { forward }),
      ...(backward === undefined ? {} : { backward }),
    };

    if (type === "movingAverage") {
      if (
        period !== undefined &&
        period >= 2 &&
        period <= 255 &&
        period < pointCount
      ) {
        result.push({ ...baseFields, type: "movingAverage", period });
      }
      continue;
    }

    const displayFields = {
      ...baseFields,
      ...(displayEquation === undefined ? {} : { displayEquation }),
      ...(displayRSquared === undefined ? {} : { displayRSquared }),
    };

    if (type === "polynomial") {
      if (order !== undefined && order >= 2 && order <= 6) {
        result.push({
          ...displayFields,
          type: "polynomial",
          order,
          ...(rawIntercept === undefined ? {} : { intercept: rawIntercept }),
        });
      }
      continue;
    }

    if (type === "logarithmic") {
      if (!fitPoints.some(({ x }) => x <= 0)) {
        result.push({ ...displayFields, type: "logarithmic" });
      }
      continue;
    }

    if (type === "exponential") {
      if (!fitPoints.some(({ y }) => y <= 0)) {
        result.push({
          ...displayFields,
          type: "exponential",
          ...(rawIntercept === undefined ? {} : { intercept: rawIntercept }),
        });
      }
      continue;
    }

    if (type === "power") {
      if (!fitPoints.some(({ x, y }) => x <= 0 || y <= 0)) {
        result.push({ ...displayFields, type: "power" });
      }
      continue;
    }

    // linear
    result.push({
      ...displayFields,
      type: "linear",
      ...(rawIntercept === undefined ? {} : { intercept: rawIntercept }),
    });
  }

  return result.length === 0 ? undefined : result;
}

function alignErrorBarValues(
  values: readonly (number | null)[],
  pointCount: number,
): readonly (number | null)[] {
  return Array.from({ length: pointCount }, (_, idx) => {
    const v = values[idx];
    return v !== undefined && v !== null && Number.isFinite(v) && v >= 0
      ? v
      : null;
  });
}

function readSingleErrorBars(
  errCursor: XmlElementCursor,
  pointCount: number,
): { readonly direction: "x" | "y"; readonly bars?: ChartErrorBars } {
  const errDirRaw = readValAttr(errCursor.child(C_NS, "errDir"));
  const direction: "x" | "y" = errDirRaw === "x" ? "x" : "y";

  const includeParsed = ErrBarTypeSchema.safeParse(
    errCursor.child(C_NS, "errBarType")?.attrsObject(["val"]) ?? {},
  );
  const include: ChartErrorBarInclude = includeParsed.success
    ? includeParsed.data.val
    : "both";

  const errValTypeRaw = readValAttr(errCursor.child(C_NS, "errValType"));
  const barType =
    errValTypeRaw === undefined
      ? undefined
      : chartErrorBarValueTypeFromNative[errValTypeRaw];

  const noEndCap = readBoolVal(errCursor.child(C_NS, "noEndCap"));
  const endCap = noEndCap === true ? false : undefined;

  const plusCursor = errCursor.child(C_NS, "plus");
  const rawPlus =
    plusCursor === undefined ? undefined : readNumericDataSource(plusCursor);
  if (plusCursor !== undefined) consumeSubtree(plusCursor);

  const minusCursor = errCursor.child(C_NS, "minus");
  const rawMinus =
    minusCursor === undefined ? undefined : readNumericDataSource(minusCursor);
  if (minusCursor !== undefined) consumeSubtree(minusCursor);

  const val = readFloatVal(errCursor.child(C_NS, "val"));

  const spPr = errCursor.child(C_NS, "spPr");
  const line = spPr === undefined ? undefined : readLineFromSpPr(spPr).line;
  if (spPr !== undefined) consumeSubtree(spPr);

  consumeSubtree(errCursor);

  if (barType === undefined) return { direction };

  const styleFields = {
    ...(endCap === undefined ? {} : { endCap }),
    ...(line === undefined ? {} : { line }),
  };

  if (barType === "standardError") {
    return {
      direction,
      bars: {
        ...styleFields,
        type: "standardError",
        ...(include === "both" ? {} : { include }),
      },
    };
  }

  if (barType === "custom") {
    const plus =
      rawPlus === undefined
        ? undefined
        : alignErrorBarValues(rawPlus, pointCount);
    const minus =
      rawMinus === undefined
        ? undefined
        : alignErrorBarValues(rawMinus, pointCount);

    let custom: ChartCustomErrorBars | undefined;
    if (include === "plus" && plus !== undefined) {
      custom = { ...styleFields, type: "custom", include: "plus", plus };
    } else if (include === "minus" && minus !== undefined) {
      custom = { ...styleFields, type: "custom", include: "minus", minus };
    } else if (plus !== undefined && minus !== undefined) {
      custom = { ...styleFields, type: "custom", include: "both", plus, minus };
    }
    return { direction, ...(custom === undefined ? {} : { bars: custom }) };
  }

  if (val === undefined || !Number.isFinite(val) || val < 0) {
    return { direction };
  }
  return {
    direction,
    bars: {
      ...styleFields,
      type: barType,
      value: val,
      ...(include === "both" ? {} : { include }),
    },
  };
}

function readCategoryErrorBars(
  serCursor: XmlElementCursor,
  kind: Exclude<ChartKind, "scatter">,
  pointCount: number,
): ChartErrorBars | undefined {
  const nodes = serCursor.children(C_NS, "errBars");
  if (nodes.length === 0) return undefined;
  let result: ChartErrorBars | undefined;
  for (const node of nodes) {
    const parsed = readSingleErrorBars(node, pointCount);
    if (kind !== "pie" && kind !== "doughnut" && result === undefined) {
      result = parsed.bars;
    }
  }
  return result;
}

function readScatterErrorBars(
  serCursor: XmlElementCursor,
  pointCount: number,
): ScatterErrorBars | undefined {
  const nodes = serCursor.children(C_NS, "errBars");
  if (nodes.length === 0) return undefined;
  let x: ChartErrorBars | undefined;
  let y: ChartErrorBars | undefined;
  for (const node of nodes) {
    const parsed = readSingleErrorBars(node, pointCount);
    if (parsed.bars === undefined) continue;
    if (parsed.direction === "x" && x === undefined) x = parsed.bars;
    else if (parsed.direction === "y" && y === undefined) y = parsed.bars;
  }
  if (x === undefined && y === undefined) return undefined;
  return {
    ...(x === undefined ? {} : { x }),
    ...(y === undefined ? {} : { y }),
  };
}

type ImportedCategorySeriesDraft = {
  readonly order: number;
  readonly name: string;
  readonly categories: readonly string[];
  readonly categoryFormula?: string;
  readonly values: readonly (number | null)[];
  readonly color?: ColorValue;
  readonly line?: ChartLine;
  readonly explosion?: number;
  readonly marker?: boolean | ChartMarker;
  readonly smooth?: boolean;
  readonly dataLabels?: false | ChartDataLabelSettings;
  readonly dataLabelsStyle?: ChartTextStyle;
  readonly points?: Readonly<Record<number, ChartPoint>>;
  readonly trendlines?: readonly ChartTrendline[];
  readonly errorBars?: ChartErrorBars;
  readonly datalabelsRange?: readonly string[];
};

function readCategorySeries(
  serCursor: XmlElementCursor,
  kind: Exclude<ChartKind, "scatter">,
  fallbackIndex: number,
  stacking: "stacked" | "percent" | undefined,
  inheritedChartStyle: ChartTextStyle | undefined,
): ImportedCategorySeriesDraft {
  const idx = readIntVal(serCursor.child(C_NS, "idx")) ?? fallbackIndex;
  const order = readIntVal(serCursor.child(C_NS, "order")) ?? idx;
  const name =
    readSeriesName(serCursor.child(C_NS, "tx")) ??
    `Series ${fallbackIndex + 1}`;

  const spPr = serCursor.child(C_NS, "spPr");
  const { color, line } = spPr === undefined ? {} : readSeriesSpPr(spPr, false);

  const invertIfNegative = serCursor.child(C_NS, "invertIfNegative");
  if (invertIfNegative !== undefined) consumeSubtree(invertIfNegative);

  const markerCursor = serCursor.child(C_NS, "marker");
  const marker =
    markerCursor === undefined
      ? undefined
      : kind === "line"
        ? readMarkerElement(markerCursor)
        : (consumeSubtree(markerCursor), undefined);

  const explosionCursor = serCursor.child(C_NS, "explosion");
  const explosionVal = readIntVal(explosionCursor);
  const explosion =
    (kind === "pie" || kind === "doughnut") &&
    explosionVal !== undefined &&
    explosionVal >= 0
      ? explosionVal
      : undefined;

  const pointsRecord: Record<number, ChartPoint> = {};
  for (const dPt of serCursor.children(C_NS, "dPt")) {
    const pointIdx = readIntVal(dPt.child(C_NS, "idx"));
    const inv = dPt.child(C_NS, "invertIfNegative");
    if (inv !== undefined) consumeSubtree(inv);
    const bubble3D = dPt.child(C_NS, "bubble3D");
    if (bubble3D !== undefined) consumeSubtree(bubble3D);
    const pMarkerCursor = dPt.child(C_NS, "marker");
    const pMarker =
      pMarkerCursor === undefined
        ? undefined
        : kind === "line"
          ? readMarkerElement(pMarkerCursor)
          : (consumeSubtree(pMarkerCursor), undefined);
    const pExplosionVal = readIntVal(dPt.child(C_NS, "explosion"));
    const pExplosion =
      (kind === "pie" || kind === "doughnut") &&
      pExplosionVal !== undefined &&
      pExplosionVal >= 0
        ? pExplosionVal
        : undefined;
    const pSpPr = dPt.child(C_NS, "spPr");
    const pStyle = pSpPr === undefined ? {} : readSeriesSpPr(pSpPr, false);
    consumeSubtree(dPt);
    if (pointIdx !== undefined && pointIdx >= 0) {
      const pointColor = kind === "area" ? undefined : pStyle.color;
      if (
        pointColor !== undefined ||
        pStyle.line !== undefined ||
        pMarker !== undefined ||
        pExplosion !== undefined
      ) {
        pointsRecord[pointIdx] = {
          ...(pointColor === undefined ? {} : { color: pointColor }),
          ...(pStyle.line === undefined ? {} : { line: pStyle.line }),
          ...(pMarker === undefined ? {} : { marker: pMarker }),
          ...(pExplosion === undefined ? {} : { explosion: pExplosion }),
        };
      }
    }
  }

  const dLblsCursor = serCursor.child(C_NS, "dLbls");
  const dLblsResult =
    dLblsCursor === undefined
      ? undefined
      : readDataLabelsElement(dLblsCursor, kind, false, inheritedChartStyle);
  if (dLblsResult !== undefined) {
    for (const [ptIdx, labelColor] of dLblsResult.pointLabelColors.entries()) {
      const existing = pointsRecord[ptIdx] ?? {};
      pointsRecord[ptIdx] = {
        ...existing,
        labelColor,
      };
    }
    for (const [ptIdx, ptLabel] of dLblsResult.pointLabels.entries()) {
      const diffed = diffPointDataLabels(ptLabel, dLblsResult.dataLabels);
      if (diffed === undefined) continue;
      const existing = pointsRecord[ptIdx] ?? {};
      pointsRecord[ptIdx] = {
        ...existing,
        dataLabels: diffed,
      };
    }
  }

  const catCursor = serCursor.child(C_NS, "cat");
  const catData =
    catCursor === undefined ? { values: [] } : readStringDataSource(catCursor);

  const valCursor = serCursor.child(C_NS, "val");
  const valData =
    valCursor === undefined ? [] : readNumericDataSource(valCursor);

  const smoothCursor = serCursor.child(C_NS, "smooth");
  const smoothVal = readBoolVal(smoothCursor);
  const smooth = kind === "line" ? smoothVal : undefined;

  const fitPoints = valData.flatMap((y, valueIndex) =>
    y === null || !Number.isFinite(y) ? [] : [{ x: valueIndex + 1, y }],
  );
  const trendlines = readSeriesTrendlines(
    serCursor,
    kind,
    stacking,
    valData.length,
    fitPoints,
  );
  const errorBars = readCategoryErrorBars(serCursor, kind, valData.length);

  let datalabelsRange: readonly string[] | undefined;
  const extLst = serCursor.child(C_NS, "extLst");
  if (extLst !== undefined) {
    for (const ext of extLst.children(C_NS, "ext")) {
      const uri = ext.attr("uri");
      if (uri === C15_CAT_RANGE_EXT_URI) {
        const range = ext.child(C15_NS, "datalabelsRange");
        if (range !== undefined) {
          const f = range.child(C15_NS, "f");
          if (f !== undefined) consumeSubtree(f);
          const cache = range.child(C15_NS, "dlblRangeCache");
          if (cache !== undefined) {
            datalabelsRange = readStringCacheOrLit(cache);
          }
          consumeSubtree(range);
        }
        consumeSubtree(ext);
      } else if (uri === C16_UNIQUE_ID_EXT_URI) {
        consumeSubtree(ext);
      }
    }
  }

  const pointCount = Math.max(catData.values.length, valData.length);
  const normalizedPoints: Record<number, ChartPoint> = {};
  for (const [key, pt] of Object.entries(pointsRecord)) {
    const numKey = Number(key);
    if (numKey >= 0 && numKey < pointCount) {
      normalizedPoints[numKey] = pt;
    }
  }

  return {
    order,
    name: name.trim() === "" ? `Series ${fallbackIndex + 1}` : name,
    categories: catData.values,
    ...(catData.formula === undefined
      ? {}
      : { categoryFormula: catData.formula }),
    values: valData,
    ...(color === undefined ? {} : { color }),
    ...(line === undefined ? {} : { line }),
    ...(explosion === undefined ? {} : { explosion }),
    ...(marker === undefined ? {} : { marker }),
    ...(smooth === undefined ? {} : { smooth }),
    ...(dLblsResult?.dataLabels === undefined
      ? {}
      : { dataLabels: dLblsResult.dataLabels }),
    ...(dLblsResult?.dataLabelsStyle === undefined
      ? {}
      : { dataLabelsStyle: dLblsResult.dataLabelsStyle }),
    ...(Object.keys(normalizedPoints).length === 0
      ? {}
      : { points: normalizedPoints }),
    ...(trendlines === undefined ? {} : { trendlines }),
    ...(errorBars === undefined ? {} : { errorBars }),
    ...(datalabelsRange === undefined ? {} : { datalabelsRange }),
  };
}

function readScatterSeries(
  serCursor: XmlElementCursor,
  groupLine: ScatterLine | undefined,
  groupMarkers: boolean,
  fallbackIndex: number,
  inheritedChartStyle: ChartTextStyle | undefined,
): {
  readonly order: number;
  readonly series: ScatterChartSeries;
  readonly dataLabelsStyle?: ChartTextStyle;
} {
  const idx = readIntVal(serCursor.child(C_NS, "idx")) ?? fallbackIndex;
  const order = readIntVal(serCursor.child(C_NS, "order")) ?? idx;
  const name =
    readSeriesName(serCursor.child(C_NS, "tx")) ??
    `Series ${fallbackIndex + 1}`;

  const spPr = serCursor.child(C_NS, "spPr");
  const showLine = groupLine !== undefined && groupLine !== false;
  const spStyle = spPr === undefined ? {} : readSeriesSpPr(spPr, !showLine);

  const markerCursor = serCursor.child(C_NS, "marker");
  const rawMarker =
    markerCursor === undefined ? undefined : readMarkerElement(markerCursor);
  const marker =
    !groupMarkers && rawMarker !== undefined && rawMarker !== false
      ? undefined
      : rawMarker;

  const pointsRecord: Record<number, ScatterChartPoint> = {};
  for (const dPt of serCursor.children(C_NS, "dPt")) {
    const pointIdx = readIntVal(dPt.child(C_NS, "idx"));
    const pMarkerCursor = dPt.child(C_NS, "marker");
    const rawPMarker =
      pMarkerCursor === undefined
        ? undefined
        : readMarkerElement(pMarkerCursor);
    const pMarker =
      !groupMarkers && rawPMarker !== undefined && rawPMarker !== false
        ? undefined
        : rawPMarker;
    const pSpPr = dPt.child(C_NS, "spPr");
    const pStyle = pSpPr === undefined ? {} : readSeriesSpPr(pSpPr, false);
    consumeSubtree(dPt);
    if (pointIdx !== undefined && pointIdx >= 0) {
      if (
        pStyle.color !== undefined ||
        pStyle.line !== undefined ||
        pMarker !== undefined
      ) {
        pointsRecord[pointIdx] = {
          ...(pStyle.color === undefined ? {} : { color: pStyle.color }),
          ...(pStyle.line === undefined ? {} : { line: pStyle.line }),
          ...(pMarker === undefined ? {} : { marker: pMarker }),
        };
      }
    }
  }

  const dLblsCursor = serCursor.child(C_NS, "dLbls");
  const dLblsResult =
    dLblsCursor === undefined
      ? undefined
      : readDataLabelsElement(
          dLblsCursor,
          "scatter",
          false,
          inheritedChartStyle,
        );
  if (dLblsResult !== undefined) {
    for (const [ptIdx, labelColor] of dLblsResult.pointLabelColors.entries()) {
      const existing = pointsRecord[ptIdx] ?? {};
      pointsRecord[ptIdx] = {
        ...existing,
        labelColor,
      };
    }
    for (const [ptIdx, ptLabel] of dLblsResult.pointLabels.entries()) {
      const diffed = diffPointDataLabels(ptLabel, dLblsResult.dataLabels);
      if (diffed === undefined) continue;
      const existing = pointsRecord[ptIdx] ?? {};
      pointsRecord[ptIdx] = {
        ...existing,
        dataLabels: diffed as false | ScatterDataLabels,
      };
    }
  }

  const xCursor = serCursor.child(C_NS, "xVal");
  const yCursor = serCursor.child(C_NS, "yVal");
  let x = xCursor === undefined ? [] : readNumericDataSource(xCursor);
  let y = yCursor === undefined ? [] : readNumericDataSource(yCursor);
  if (x.length !== y.length) {
    const maxLen = Math.max(x.length, y.length);
    if (x.length < maxLen) {
      x = [...x, ...Array.from<null>({ length: maxLen - x.length }).fill(null)];
    }
    if (y.length < maxLen) {
      y = [...y, ...Array.from<null>({ length: maxLen - y.length }).fill(null)];
    }
  }

  const fitPoints = x.flatMap((xv, pointIdx) => {
    const yv = y[pointIdx];
    return xv === null ||
      yv === undefined ||
      yv === null ||
      !Number.isFinite(xv) ||
      !Number.isFinite(yv)
      ? []
      : [{ x: xv, y: yv }];
  });
  const trendlines = readSeriesTrendlines(
    serCursor,
    "scatter",
    undefined,
    x.length,
    fitPoints,
  );
  const errorBars = readScatterErrorBars(serCursor, x.length);

  const smoothCursor = serCursor.child(C_NS, "smooth");
  if (smoothCursor !== undefined) consumeSubtree(smoothCursor);

  const extLst = serCursor.child(C_NS, "extLst");
  if (extLst !== undefined) {
    for (const ext of extLst.children(C_NS, "ext")) {
      if (ext.attr("uri") === C16_UNIQUE_ID_EXT_URI) {
        consumeSubtree(ext);
      }
    }
  }

  const normalizedPoints: Record<number, ScatterChartPoint> = {};
  for (const [key, pt] of Object.entries(pointsRecord)) {
    const numKey = Number(key);
    if (numKey >= 0 && numKey < x.length) {
      normalizedPoints[numKey] = pt;
    }
  }

  return {
    order,
    ...(dLblsResult?.dataLabelsStyle === undefined
      ? {}
      : { dataLabelsStyle: dLblsResult.dataLabelsStyle }),
    series: {
      name: name.trim() === "" ? `Series ${fallbackIndex + 1}` : name,
      x,
      y,
      ...(spStyle.color === undefined ? {} : { color: spStyle.color }),
      ...(spStyle.line === undefined ? {} : { line: spStyle.line }),
      ...(dLblsResult?.dataLabels === undefined
        ? {}
        : {
            dataLabels: dLblsResult.dataLabels as false | ScatterDataLabels,
          }),
      ...(marker === undefined ? {} : { marker }),
      ...(Object.keys(normalizedPoints).length === 0
        ? {}
        : { points: normalizedPoints }),
      ...(trendlines === undefined ? {} : { trendlines }),
      ...(errorBars === undefined ? {} : { errorBars }),
    },
  };
}

type ImportedAxisRecord = {
  readonly id: number;
  readonly kind: "category" | "value";
  readonly crossAxisId?: number;
  readonly position?: ChartAxisPosition;
  readonly visible?: boolean;
  readonly reversed?: boolean;
  readonly gridlines?: boolean | ChartLine;
  readonly minorGridlines?: boolean | ChartLine;
  readonly line?: boolean | ChartLine;
  readonly style?: ChartTextStyle;
  readonly title?: ChartTitle;
  readonly majorTickMark?: ChartTickMark;
  readonly minorTickMark?: ChartTickMark;
  readonly tickLabelPosition?: ChartTickLabelPosition;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  readonly numberFormat?: string;
  readonly logBase?: number;
  readonly rawCrossing?:
    | { readonly kind: "crosses"; readonly value: ChartAxisCrossing }
    | { readonly kind: "crossesAt"; readonly value: number };
};

function readAxisGridline(
  el: XmlElementCursor | undefined,
): boolean | ChartLine | undefined {
  if (el === undefined) return undefined;
  const spPr = el.child(C_NS, "spPr");
  const result =
    spPr === undefined ? true : (readLineFromSpPr(spPr).line ?? true);
  consumeSubtree(el);
  return result;
}

function readAxisElement(
  axisCursor: XmlElementCursor,
  kind: "category" | "value",
  inheritedChartStyle: ChartTextStyle | undefined,
): ImportedAxisRecord | undefined {
  const id = readIntVal(axisCursor.child(C_NS, "axId"));
  if (id === undefined) return undefined;

  let reversed: boolean | undefined;
  let min: number | undefined;
  let max: number | undefined;
  let logBase: number | undefined;
  const scaling = axisCursor.child(C_NS, "scaling");
  if (scaling !== undefined) {
    const orientation = readValAttr(scaling.child(C_NS, "orientation"));
    if (orientation === "maxMin") reversed = true;
    if (kind === "value") {
      logBase = readFloatVal(scaling.child(C_NS, "logBase"));
      min = readFloatVal(scaling.child(C_NS, "min"));
      max = readFloatVal(scaling.child(C_NS, "max"));
    }
    consumeSubtree(scaling);
  }

  const deleted = readBoolVal(axisCursor.child(C_NS, "delete"));
  const visible = deleted === true ? false : undefined;

  let position: ChartAxisPosition | undefined;
  const axPos = axisCursor.child(C_NS, "axPos");
  if (axPos !== undefined) {
    const val = axPos.attrsObject(["val"])["val"];
    if (val !== undefined) {
      position = chartAxisPositionFromNative[val];
    }
  }

  const gridlines = readAxisGridline(axisCursor.child(C_NS, "majorGridlines"));
  const minorGridlines = readAxisGridline(
    axisCursor.child(C_NS, "minorGridlines"),
  );

  const txPr = axisCursor.child(C_NS, "txPr");
  const rawAxisStyle =
    txPr === undefined ? undefined : readChartTextProperties(txPr);
  const style = diffChartTextStyle(rawAxisStyle, inheritedChartStyle);

  const titleCursor = axisCursor.child(C_NS, "title");
  const title =
    titleCursor === undefined
      ? undefined
      : readChartTitle(
          titleCursor,
          rawAxisStyle === undefined && inheritedChartStyle === undefined
            ? undefined
            : { ...inheritedChartStyle, ...rawAxisStyle },
        );

  const numFmtCursor = axisCursor.child(C_NS, "numFmt");
  let numberFormat: string | undefined;
  if (numFmtCursor !== undefined) {
    const code = numFmtCursor.attr("formatCode");
    const sourceLinked = numFmtCursor.attr("sourceLinked");
    if (
      kind === "value" &&
      code !== undefined &&
      code.trim() !== "" &&
      (sourceLinked === "0" || sourceLinked === "false" || code !== "General")
    ) {
      numberFormat = code;
    }
    consumeSubtree(numFmtCursor);
  }

  const majorTickParsed = TickMarkSchema.safeParse(
    axisCursor.child(C_NS, "majorTickMark")?.attrsObject(["val"]) ?? {},
  );
  const majorTickMark = majorTickParsed.success
    ? majorTickParsed.data.val
    : undefined;

  const minorTickParsed = TickMarkSchema.safeParse(
    axisCursor.child(C_NS, "minorTickMark")?.attrsObject(["val"]) ?? {},
  );
  const minorTickMark = minorTickParsed.success
    ? minorTickParsed.data.val
    : undefined;

  const tickLblParsed = TickLblPosSchema.safeParse(
    axisCursor.child(C_NS, "tickLblPos")?.attrsObject(["val"]) ?? {},
  );
  const tickLabelPosition = tickLblParsed.success
    ? tickLblParsed.data.val
    : undefined;

  let line: boolean | ChartLine | undefined;
  const spPr = axisCursor.child(C_NS, "spPr");
  if (spPr !== undefined) {
    const lineRes = readLineFromSpPr(spPr);
    if (lineRes.noLine) line = false;
    else if (lineRes.line !== undefined) line = lineRes.line;
    consumeSubtree(spPr);
  }

  const crossAxisId = readIntVal(axisCursor.child(C_NS, "crossAx"));

  let rawCrossing: ImportedAxisRecord["rawCrossing"];
  const crossesAtCursor = axisCursor.child(C_NS, "crossesAt");
  const crossesCursor = axisCursor.child(C_NS, "crosses");
  if (crossesAtCursor !== undefined) {
    const val = readFloatVal(crossesAtCursor);
    if (val !== undefined) rawCrossing = { kind: "crossesAt", value: val };
  } else if (crossesCursor !== undefined) {
    const parsedCross = CrossesSchema.safeParse(
      crossesCursor.attrsObject(["val"]),
    );
    if (parsedCross.success) {
      rawCrossing = { kind: "crosses", value: parsedCross.data.val };
    }
  }

  const majorUnit =
    kind === "value"
      ? readFloatVal(axisCursor.child(C_NS, "majorUnit"))
      : undefined;
  const minorUnit =
    kind === "value"
      ? readFloatVal(axisCursor.child(C_NS, "minorUnit"))
      : undefined;

  for (const tag of [
    "crossBetween",
    "auto",
    "lblAlgn",
    "lblOffset",
    "noMultiLvlLbl",
    "tickLblSkip",
    "tickMarkSkip",
  ]) {
    const el = axisCursor.child(C_NS, tag);
    if (el !== undefined) consumeSubtree(el);
  }

  return {
    id,
    kind,
    ...(crossAxisId === undefined ? {} : { crossAxisId }),
    ...(position === undefined ? {} : { position }),
    ...(visible === undefined ? {} : { visible }),
    ...(reversed === undefined ? {} : { reversed }),
    ...(gridlines === undefined ? {} : { gridlines }),
    ...(minorGridlines === undefined ? {} : { minorGridlines }),
    ...(line === undefined ? {} : { line }),
    ...(style === undefined ? {} : { style }),
    ...(title === undefined ? {} : { title }),
    ...(majorTickMark === undefined ? {} : { majorTickMark }),
    ...(minorTickMark === undefined ? {} : { minorTickMark }),
    ...(tickLabelPosition === undefined ? {} : { tickLabelPosition }),
    ...(min === undefined ? {} : { min }),
    ...(max === undefined ? {} : { max }),
    ...(majorUnit !== undefined && majorUnit > 0 ? { majorUnit } : {}),
    ...(minorUnit !== undefined && minorUnit > 0 ? { minorUnit } : {}),
    ...(numberFormat === undefined ? {} : { numberFormat }),
    ...(logBase !== undefined && logBase >= 2 && logBase <= 1000
      ? { logBase }
      : {}),
    ...(rawCrossing === undefined ? {} : { rawCrossing }),
  };
}

type ImportedGroupDraft =
  | {
      readonly kind: Exclude<ChartKind, "scatter">;
      readonly axisIds: readonly number[];
      readonly series: readonly ImportedCategorySeriesDraft[];
      readonly dataLabels?: false | ChartDataLabelSettings;
      readonly seriesDataLabelsStyle?: ChartTextStyle;
      readonly stacking?: "stacked" | "percent";
      readonly gapWidth?: number;
      readonly overlap?: number;
      readonly explosion?: number;
      readonly firstSliceAngle?: number;
      readonly holeSize?: number;
    }
  | {
      readonly kind: "scatter";
      readonly axisIds: readonly number[];
      readonly series: readonly ScatterChartSeries[];
      readonly dataLabels?: false | ScatterDataLabels;
      readonly seriesDataLabelsStyle?: ChartTextStyle;
      readonly line?: ScatterLine;
      readonly markers: boolean;
    };

function readChartGroupElement(
  groupCursor: XmlElementCursor,
  inheritedChartStyle: ChartTextStyle | undefined,
): ImportedGroupDraft | undefined {
  const local = groupCursor.localName;
  const varyColors = groupCursor.child(C_NS, "varyColors");
  if (varyColors !== undefined) consumeSubtree(varyColors);

  if (local === "scatterChart") {
    const styleCursor = groupCursor.child(C_NS, "scatterStyle");
    const parsedStyle = ScatterStyleSchema.safeParse(
      styleCursor?.attrsObject(["val"]) ?? {},
    );
    const styleVal = parsedStyle.success ? parsedStyle.data.val : "lineMarker";
    let line: ScatterLine | undefined;
    let markers = true;
    if (styleVal === "marker") {
      line = false;
      markers = true;
    } else if (styleVal === "line") {
      line = "straight";
      markers = false;
    } else if (styleVal === "lineMarker") {
      line = "straight";
      markers = true;
    } else if (styleVal === "smooth") {
      line = "smooth";
      markers = false;
    } else if (styleVal === "smoothMarker") {
      line = "smooth";
      markers = true;
    }

    const rawEntries = groupCursor
      .children(C_NS, "ser")
      .map((ser, idx) =>
        readScatterSeries(ser, line, markers, idx, inheritedChartStyle),
      )
      .toSorted((a, b) => a.order - b.order);
    const rawSeries = rawEntries.map((entry) => entry.series);
    const seriesDataLabelsStyle = rawEntries.find(
      (entry) => entry.dataLabelsStyle !== undefined,
    )?.dataLabelsStyle;

    const dLblsCursor = groupCursor.child(C_NS, "dLbls");
    const groupLabelsResult =
      dLblsCursor === undefined
        ? undefined
        : readDataLabelsElement(
            dLblsCursor,
            "scatter",
            true,
            inheritedChartStyle,
          );
    const groupLabels = groupLabelsResult?.dataLabels;

    const axisIds = groupCursor
      .children(C_NS, "axId")
      .map((c) => readIntVal(c))
      .filter((v): v is number => v !== undefined);

    return {
      kind: "scatter",
      axisIds,
      series: rawSeries,
      markers,
      ...(line === undefined ? {} : { line }),
      ...(groupLabels === undefined
        ? {}
        : { dataLabels: groupLabels as false | ScatterDataLabels }),
      ...(seriesDataLabelsStyle === undefined ? {} : { seriesDataLabelsStyle }),
    };
  }

  let kind: Exclude<ChartKind, "scatter">;
  if (local === "barChart") {
    const barDirCursor = groupCursor.child(C_NS, "barDir");
    const parsedDir = BarDirSchema.safeParse(
      barDirCursor?.attrsObject(["val"]) ?? {},
    );
    kind = parsedDir.success && parsedDir.data.val === "bar" ? "bar" : "column";
  } else if (local === "lineChart") {
    kind = "line";
  } else if (local === "areaChart") {
    kind = "area";
  } else if (local === "pieChart") {
    kind = "pie";
  } else if (local === "doughnutChart") {
    kind = "doughnut";
  } else {
    return undefined;
  }

  let stacking: "stacked" | "percent" | undefined;
  if (
    kind === "column" ||
    kind === "bar" ||
    kind === "line" ||
    kind === "area"
  ) {
    const groupingCursor = groupCursor.child(C_NS, "grouping");
    const parsedGrouping = GroupingSchema.safeParse(
      groupingCursor?.attrsObject(["val"]) ?? {},
    );
    if (parsedGrouping.success) {
      if (parsedGrouping.data.val === "stacked") stacking = "stacked";
      else if (parsedGrouping.data.val === "percentStacked")
        stacking = "percent";
    }
  }

  const rawSeries = groupCursor
    .children(C_NS, "ser")
    .map((ser, idx) =>
      readCategorySeries(ser, kind, idx, stacking, inheritedChartStyle),
    )
    .toSorted((a, b) => a.order - b.order);
  const seriesDataLabelsStyle = rawSeries.find(
    (entry) => entry.dataLabelsStyle !== undefined,
  )?.dataLabelsStyle;

  const dLblsCursor = groupCursor.child(C_NS, "dLbls");
  const groupLabelsResult =
    dLblsCursor === undefined
      ? undefined
      : readDataLabelsElement(dLblsCursor, kind, true, inheritedChartStyle);
  const groupLabels = groupLabelsResult?.dataLabels;

  let gapWidth: number | undefined;
  let overlap: number | undefined;
  if (kind === "column" || kind === "bar") {
    const gw = readIntVal(groupCursor.child(C_NS, "gapWidth"));
    if (gw !== undefined && gw >= 0 && gw <= 500) gapWidth = gw;
    const ov = readIntVal(groupCursor.child(C_NS, "overlap"));
    if (ov !== undefined && ov >= -100 && ov <= 100) overlap = ov;
  }

  if (kind === "line") {
    const markerEl = groupCursor.child(C_NS, "marker");
    if (markerEl !== undefined) consumeSubtree(markerEl);
    const smoothEl = groupCursor.child(C_NS, "smooth");
    if (smoothEl !== undefined) consumeSubtree(smoothEl);
  }

  let firstSliceAngle: number | undefined;
  let holeSize: number | undefined;
  if (kind === "pie" || kind === "doughnut") {
    const ang = readIntVal(groupCursor.child(C_NS, "firstSliceAng"));
    if (ang !== undefined && ang >= 0 && ang <= 360) firstSliceAngle = ang;
    if (kind === "doughnut") {
      const hs = readIntVal(groupCursor.child(C_NS, "holeSize"));
      if (hs !== undefined && hs >= 10 && hs <= 90) holeSize = hs;
    }
  }

  const axisIds = groupCursor
    .children(C_NS, "axId")
    .map((c) => readIntVal(c))
    .filter((v): v is number => v !== undefined);

  return {
    kind,
    axisIds,
    series: rawSeries,
    ...(groupLabels === undefined ? {} : { dataLabels: groupLabels }),
    ...(seriesDataLabelsStyle === undefined ? {} : { seriesDataLabelsStyle }),
    ...(stacking === undefined ? {} : { stacking }),
    ...(gapWidth === undefined ? {} : { gapWidth }),
    ...(overlap === undefined ? {} : { overlap }),
    ...(firstSliceAngle === undefined ? {} : { firstSliceAngle }),
    ...(holeSize === undefined ? {} : { holeSize }),
  };
}

function narrowGroupDataLabels<const P extends ChartDataLabelPosition>(
  labels: false | ChartDataLabelSettings | undefined,
  allowedPositions: readonly P[],
): false | ChartDataLabels<P> | undefined {
  if (labels === undefined || labels === false) return labels;
  const { position, xValue: _xValue, ...rest } = labels;
  if (
    position !== undefined &&
    (allowedPositions as readonly string[]).includes(position)
  ) {
    return { ...rest, position: position as P };
  }
  return rest;
}

function sameStringArray(a: readonly string[], b: readonly string[]): boolean {
  return a.length === b.length && a.every((item, idx) => item === b[idx]);
}

type RawAxisCrossing =
  | { readonly kind: "crosses"; readonly value: ChartAxisCrossing }
  | { readonly kind: "crossesAt"; readonly value: number };

function sameRawCrossing(
  left: RawAxisCrossing,
  right: RawAxisCrossing,
): boolean {
  return left.kind === right.kind && left.value === right.value;
}

function positionForImportedRawCrossing(
  crossing: RawAxisCrossing,
  partner: ImportedAxisRecord,
  direction: "horizontal" | "vertical",
): ChartAxisPosition | undefined {
  if (crossing.kind !== "crosses" || crossing.value === "autoZero") {
    return undefined;
  }
  const maxAtFarEnd = partner.reversed !== true;
  if (direction === "vertical") {
    return (crossing.value === "max") === maxAtFarEnd ? "right" : "left";
  }
  return (crossing.value === "max") === maxAtFarEnd ? "top" : "bottom";
}

function importedCrossesKnownValueInterior(
  crossing: RawAxisCrossing,
  partner: ImportedAxisRecord,
): boolean {
  if (partner.kind !== "value") return false;
  if (partner.min === undefined || partner.max === undefined) return false;
  const low = Math.min(partner.min, partner.max);
  const high = Math.max(partner.min, partner.max);
  if (crossing.kind === "crossesAt") {
    return crossing.value > low && crossing.value < high;
  }
  return (
    crossing.kind === "crosses" &&
    crossing.value === "autoZero" &&
    low < 0 &&
    high > 0
  );
}

function importedKnownValueEndpoint(
  crossing: RawAxisCrossing,
  partner: ImportedAxisRecord,
): "min" | "max" | undefined {
  if (partner.kind !== "value") return undefined;
  if (
    partner.min === undefined ||
    partner.max === undefined ||
    partner.min === partner.max
  ) {
    return undefined;
  }
  if (crossing.kind === "crossesAt") {
    if (crossing.value === partner.min) return "min";
    if (crossing.value === partner.max) return "max";
    return undefined;
  }
  if (crossing.kind !== "crosses" || crossing.value !== "autoZero") {
    return undefined;
  }
  if (partner.min === 0) return "min";
  if (partner.max === 0) return "max";
  return undefined;
}

function withLiftedDataLabelsStyle(
  groupLabels: false | ChartDataLabelSettings | undefined,
  seriesStyle: ChartTextStyle | undefined,
): false | ChartDataLabelSettings | undefined {
  if (seriesStyle === undefined) return groupLabels;
  if (groupLabels === false || groupLabels === undefined) {
    return { style: seriesStyle };
  }
  return groupLabels.style === undefined
    ? { ...groupLabels, style: seriesStyle }
    : groupLabels;
}

function buildChartAxesAndGroups(
  groupDrafts: readonly ImportedGroupDraft[],
  importedAxes: readonly ImportedAxisRecord[],
): { readonly axes?: ChartAxes; readonly groups: readonly ChartGroup[] } {
  const axisById = new Map(importedAxes.map((ax) => [ax.id, ax]));
  const axisDirectionById = new Map<number, "horizontal" | "vertical">();
  for (const group of groupDrafts) {
    if (group.kind === "pie" || group.kind === "doughnut") continue;
    const [firstId, secondId] = group.axisIds;
    if (group.kind === "bar") {
      if (firstId !== undefined) axisDirectionById.set(firstId, "vertical");
      if (secondId !== undefined) axisDirectionById.set(secondId, "horizontal");
    } else {
      if (firstId !== undefined) axisDirectionById.set(firstId, "horizontal");
      if (secondId !== undefined) axisDirectionById.set(secondId, "vertical");
    }
  }

  // Map native crosses/crossesAt back to author-facing crossing partners.
  // An OOXML crossing on axis A positions A in crossAx partner B's domain.
  const incomingByTarget = new Map<number, ImportedAxisRecord[]>();
  for (const ax of importedAxes) {
    if (ax.crossAxisId === undefined) continue;
    const list = incomingByTarget.get(ax.crossAxisId) ?? [];
    list.push(ax);
    incomingByTarget.set(ax.crossAxisId, list);
  }

  const partnerCrossingFor = new Map<number, RawAxisCrossing>();
  const omitPositionIds = new Set<number>();
  for (const [targetId, incoming] of incomingByTarget.entries()) {
    const targetAxis = axisById.get(targetId);
    if (targetAxis === undefined || incoming.length === 0) continue;
    const firstCrossing = incoming[0]?.rawCrossing;
    if (firstCrossing === undefined) continue;
    if (
      !incoming.every(
        (ax) =>
          ax.rawCrossing !== undefined &&
          sameRawCrossing(ax.rawCrossing, firstCrossing),
      )
    ) {
      continue;
    }

    if (firstCrossing.kind === "crossesAt") {
      if (
        targetAxis.kind === "category" &&
        (!Number.isSafeInteger(firstCrossing.value) || firstCrossing.value < 1)
      ) {
        continue;
      }
      partnerCrossingFor.set(targetId, firstCrossing);
      for (const ax of incoming) {
        if (ax.position === undefined) continue;
        const axDir = axisDirectionById.get(ax.id);
        const endpoint = importedKnownValueEndpoint(firstCrossing, targetAxis);
        const endpointEdge =
          endpoint !== undefined && axDir !== undefined
            ? positionForImportedRawCrossing(
                { kind: "crosses", value: endpoint },
                targetAxis,
                axDir,
              )
            : undefined;
        if (
          importedCrossesKnownValueInterior(firstCrossing, targetAxis) ||
          (endpointEdge !== undefined && endpointEdge !== ax.position)
        ) {
          omitPositionIds.add(ax.id);
        }
      }
      continue;
    }

    if (firstCrossing.value === "autoZero") {
      if (importedCrossesKnownValueInterior(firstCrossing, targetAxis)) {
        partnerCrossingFor.set(targetId, firstCrossing);
        for (const ax of incoming) {
          omitPositionIds.add(ax.id);
        }
      }
      continue;
    }

    const agreesWithPositions = incoming.every((ax) => {
      if (ax.position === undefined) return true;
      const axDir = axisDirectionById.get(ax.id);
      if (axDir === undefined) return true;
      const edge = positionForImportedRawCrossing(
        firstCrossing,
        targetAxis,
        axDir,
      );
      return edge === undefined || edge === ax.position;
    });
    if (agreesWithPositions) {
      partnerCrossingFor.set(targetId, firstCrossing);
    }
  }

  // Assign deterministic semantic names for axes.
  const axisNameById = new Map<number, string>();
  const usedNames = new Set<string>();
  const assignName = (id: number, preferred: string): string => {
    const existing = axisNameById.get(id);
    if (existing !== undefined) return existing;
    let candidate = preferred;
    let suffix = 2;
    while (usedNames.has(candidate)) {
      candidate = `${preferred}${suffix}`;
      suffix += 1;
    }
    usedNames.add(candidate);
    axisNameById.set(id, candidate);
    return candidate;
  };

  for (const group of groupDrafts) {
    if (group.kind === "pie" || group.kind === "doughnut") continue;
    if (group.kind === "scatter") {
      const [xId, yId] = group.axisIds;
      if (xId !== undefined) assignName(xId, "x");
      if (yId !== undefined) assignName(yId, "y");
    } else {
      const [catId, valId] = group.axisIds;
      if (catId !== undefined) assignName(catId, "category");
      if (valId !== undefined) assignName(valId, "value");
    }
  }

  const axesRecord: Record<string, ChartAxis> = {};
  for (const ax of importedAxes) {
    const name = axisNameById.get(ax.id);
    if (name === undefined) continue; // unreferenced axis
    const partnerCrossing = partnerCrossingFor.get(ax.id);
    const keepPosition =
      ax.position !== undefined && !omitPositionIds.has(ax.id);
    const common = {
      ...(ax.visible === undefined ? {} : { visible: ax.visible }),
      ...(ax.gridlines === undefined ? {} : { gridlines: ax.gridlines }),
      ...(ax.minorGridlines === undefined
        ? {}
        : { minorGridlines: ax.minorGridlines }),
      ...(ax.line === undefined ? {} : { line: ax.line }),
      ...(ax.style === undefined ? {} : { style: ax.style }),
      ...(ax.title === undefined ? {} : { title: ax.title }),
      ...(ax.majorTickMark === undefined
        ? {}
        : { majorTickMark: ax.majorTickMark }),
      ...(ax.minorTickMark === undefined
        ? {}
        : { minorTickMark: ax.minorTickMark }),
      ...(ax.tickLabelPosition === undefined
        ? {}
        : { tickLabelPosition: ax.tickLabelPosition }),
      ...(ax.reversed === undefined ? {} : { reversed: ax.reversed }),
      ...(keepPosition ? { position: ax.position } : {}),
      ...(partnerCrossing?.kind === "crosses"
        ? { crosses: partnerCrossing.value }
        : {}),
      ...(partnerCrossing?.kind === "crossesAt" &&
      (ax.kind !== "category" ||
        (Number.isSafeInteger(partnerCrossing.value) &&
          partnerCrossing.value >= 1))
        ? { crossesAt: partnerCrossing.value }
        : {}),
    };
    if (ax.kind === "category") {
      axesRecord[name] = { kind: "category", ...common };
    } else {
      axesRecord[name] = {
        kind: "value",
        ...common,
        ...(ax.min === undefined ? {} : { min: ax.min }),
        ...(ax.max === undefined ? {} : { max: ax.max }),
        ...(ax.majorUnit === undefined ? {} : { majorUnit: ax.majorUnit }),
        ...(ax.minorUnit === undefined ? {} : { minorUnit: ax.minorUnit }),
        ...(ax.numberFormat === undefined
          ? {}
          : { numberFormat: ax.numberFormat }),
        ...(ax.logBase === undefined ? {} : { logBase: ax.logBase }),
      };
    }
  }

  // Ensure shared category axis domains agree across groups referencing the same category axis.
  const sharedDomainByCatAxis = new Map<number, readonly string[]>();
  for (const group of groupDrafts) {
    if (
      group.kind === "scatter" ||
      group.kind === "pie" ||
      group.kind === "doughnut"
    ) {
      continue;
    }
    const catId = group.axisIds[0];
    if (catId === undefined || sharedDomainByCatAxis.has(catId)) continue;
    const nonEmpty = group.series.find((s) => s.categories.length > 0);
    if (nonEmpty !== undefined) {
      sharedDomainByCatAxis.set(catId, nonEmpty.categories);
    }
  }

  const groups: ChartGroup[] = groupDrafts.map((group) => {
    const liftedDataLabels = withLiftedDataLabelsStyle(
      group.dataLabels,
      group.seriesDataLabelsStyle,
    );

    if (group.kind === "scatter") {
      const [xId, yId] = group.axisIds;
      const xName = xId === undefined ? undefined : axisNameById.get(xId);
      const yName = yId === undefined ? undefined : axisNameById.get(yId);
      return {
        kind: "scatter",
        data: { kind: "xy" },
        series: group.series,
        ...(xName !== undefined && yName !== undefined && xName !== yName
          ? { axes: { x: xName, y: yName } }
          : {}),
        ...(group.line === undefined ? {} : { line: group.line }),
        markers: group.markers,
        ...(liftedDataLabels === undefined
          ? {}
          : { dataLabels: liftedDataLabels as false | ScatterDataLabels }),
      };
    }

    if (group.kind === "doughnut") {
      const isIndependent =
        group.series.some((s) => s.datalabelsRange !== undefined) ||
        (group.series.length > 1 &&
          group.series.some(
            (s) =>
              (s.categoryFormula !== undefined &&
                group.series[0]?.categoryFormula !== undefined &&
                s.categoryFormula !== group.series[0].categoryFormula) ||
              !sameStringArray(s.categories, group.series[0]?.categories ?? []),
          ));

      if (isIndependent) {
        const series: IndependentCategorySeries[] = group.series.map((item) => {
          const cats = item.datalabelsRange ?? item.categories;
          const values =
            item.values.length === cats.length
              ? item.values
              : cats.map((_, idx) => item.values[idx] ?? null);
          return {
            name: item.name,
            categories: cats,
            values,
            ...(item.color === undefined ? {} : { color: item.color }),
            ...(item.line === undefined ? {} : { line: item.line }),
            ...(item.explosion === undefined
              ? {}
              : { explosion: item.explosion }),
            ...(item.dataLabels === undefined
              ? {}
              : { dataLabels: item.dataLabels }),
            ...(item.points === undefined ? {} : { points: item.points }),
            ...(item.trendlines === undefined
              ? {}
              : { trendlines: item.trendlines }),
            ...(item.errorBars === undefined
              ? {}
              : { errorBars: item.errorBars }),
          };
        });
        const indDoughnutLabels = narrowGroupDataLabels(liftedDataLabels, []);
        return {
          kind: "doughnut",
          data: { kind: "independent-categories" },
          series,
          ...(group.firstSliceAngle === undefined
            ? {}
            : { firstSliceAngle: group.firstSliceAngle }),
          ...(group.holeSize === undefined ? {} : { holeSize: group.holeSize }),
          ...(indDoughnutLabels === undefined
            ? {}
            : {
                dataLabels: indDoughnutLabels as
                  | false
                  | (Omit<ChartDataLabels, "position"> & {
                      readonly position?: never;
                    }),
              }),
        };
      }
    }

    const catAxisId =
      group.kind === "pie" || group.kind === "doughnut"
        ? undefined
        : group.axisIds[0];
    const domain =
      (catAxisId !== undefined
        ? sharedDomainByCatAxis.get(catAxisId)
        : undefined) ??
      group.series.find((s) => s.categories.length > 0)?.categories ??
      [];

    const normalizedSeries: ChartSeries[] = (
      group.kind === "pie" ? group.series.slice(0, 1) : group.series
    ).map((item) => {
      const values =
        domain.length === 0
          ? []
          : item.values.length === domain.length
            ? item.values
            : domain.map((_, idx) => item.values[idx] ?? null);
      const safeValues =
        group.kind === "pie" || group.kind === "doughnut"
          ? values.map((v) => (v !== null && v < 0 ? null : v))
          : values;
      return {
        name: item.name,
        values: safeValues,
        ...(item.color === undefined ? {} : { color: item.color }),
        ...(item.line === undefined ? {} : { line: item.line }),
        ...(item.explosion === undefined ? {} : { explosion: item.explosion }),
        ...(item.dataLabels === undefined
          ? {}
          : { dataLabels: item.dataLabels }),
        ...(item.marker === undefined ? {} : { marker: item.marker }),
        ...(item.smooth === undefined ? {} : { smooth: item.smooth }),
        ...(item.points === undefined ? {} : { points: item.points }),
        ...(item.trendlines === undefined
          ? {}
          : { trendlines: item.trendlines }),
        ...(item.errorBars === undefined ? {} : { errorBars: item.errorBars }),
      };
    });

    if (group.kind === "pie") {
      const pieLabels = narrowGroupDataLabels(liftedDataLabels, [
        "center",
        "insideEnd",
        "outsideEnd",
        "bestFit",
      ]);
      return {
        kind: "pie",
        data: { kind: "shared-categories", categories: domain },
        series: normalizedSeries,
        ...(group.firstSliceAngle === undefined
          ? {}
          : { firstSliceAngle: group.firstSliceAngle }),
        ...(pieLabels === undefined ? {} : { dataLabels: pieLabels }),
      };
    }

    if (group.kind === "doughnut") {
      const doughnutLabels = narrowGroupDataLabels(liftedDataLabels, []);
      return {
        kind: "doughnut",
        data: { kind: "shared-categories", categories: domain },
        series: normalizedSeries,
        ...(group.firstSliceAngle === undefined
          ? {}
          : { firstSliceAngle: group.firstSliceAngle }),
        ...(group.holeSize === undefined ? {} : { holeSize: group.holeSize }),
        ...(doughnutLabels === undefined
          ? {}
          : {
              dataLabels: doughnutLabels as
                | false
                | (Omit<ChartDataLabels, "position"> & {
                    readonly position?: never;
                  }),
            }),
      };
    }

    const [catId, valId] = group.axisIds;
    const catName =
      catId !== undefined && axisById.has(catId)
        ? axisNameById.get(catId)
        : undefined;
    const valName =
      valId !== undefined && axisById.has(valId)
        ? axisNameById.get(valId)
        : undefined;
    const axesRef =
      catName !== undefined && valName !== undefined && catName !== valName
        ? { category: catName, value: valName }
        : undefined;

    if (group.kind === "column" || group.kind === "bar") {
      const commonColBar = {
        kind: group.kind,
        data: { kind: "shared-categories" as const, categories: domain },
        series: normalizedSeries,
        ...(axesRef === undefined ? {} : { axes: axesRef }),
        ...(group.gapWidth === undefined ? {} : { gapWidth: group.gapWidth }),
        ...(group.overlap === undefined ? {} : { overlap: group.overlap }),
      };
      if (group.stacking !== undefined) {
        const stackedLabels = narrowGroupDataLabels(liftedDataLabels, [
          "center",
          "insideEnd",
          "insideBase",
        ]);
        return {
          ...commonColBar,
          stacking: group.stacking,
          ...(stackedLabels === undefined ? {} : { dataLabels: stackedLabels }),
        };
      }
      const unstackedLabels = narrowGroupDataLabels(liftedDataLabels, [
        "center",
        "insideEnd",
        "insideBase",
        "outsideEnd",
      ]);
      return {
        ...commonColBar,
        ...(unstackedLabels === undefined
          ? {}
          : { dataLabels: unstackedLabels }),
      };
    }

    if (group.kind === "line") {
      const lineLabels = narrowGroupDataLabels(liftedDataLabels, [
        "left",
        "right",
        "bottom",
        "top",
        "center",
      ]);
      return {
        kind: "line",
        data: { kind: "shared-categories", categories: domain },
        series: normalizedSeries,
        ...(axesRef === undefined ? {} : { axes: axesRef }),
        ...(group.stacking === undefined ? {} : { stacking: group.stacking }),
        ...(lineLabels === undefined ? {} : { dataLabels: lineLabels }),
      };
    }

    const areaLabels = narrowGroupDataLabels(liftedDataLabels, []);
    return {
      kind: "area",
      data: { kind: "shared-categories", categories: domain },
      series: normalizedSeries,
      ...(axesRef === undefined ? {} : { axes: axesRef }),
      ...(group.stacking === undefined ? {} : { stacking: group.stacking }),
      ...(areaLabels === undefined
        ? {}
        : {
            dataLabels: areaLabels as
              | false
              | (Omit<ChartDataLabels, "position"> & {
                  readonly position?: never;
                }),
          }),
    };
  });

  return {
    ...(Object.keys(axesRecord).length > 0 ? { axes: axesRecord } : {}),
    groups,
  };
}

function consumeChartAlternateContent(ac: XmlElementCursor): boolean {
  const choice = ac.child(MC_NS, "Choice");
  const fallback = ac.child(MC_NS, "Fallback");
  if (choice !== undefined) {
    choice.consumeAttr("Requires");
    const c14Style = choice.child(C14_NS, "style");
    if (c14Style !== undefined) {
      consumeSubtree(c14Style);
      if (fallback !== undefined) consumeSubtree(fallback);
      return true;
    }
  }
  return false;
}

function readLinkFromCNvPr(
  cNvPr: XmlElementCursor,
  part: string,
  slideIndex: number,
  pkg: PptxPackage,
  diagnostics: ImportDiagnostic[],
): LinkTarget | undefined {
  const hlinkClick = cNvPr.child(A_NS, "hlinkClick");
  if (hlinkClick === undefined) return undefined;
  const rId = hlinkClick.attr("id", R_NS);
  const action = hlinkClick.attr("action");
  if (action === "ppaction://hlinkshowjump?jump=firstslide") {
    return { action: "home" };
  }
  if (rId !== undefined) {
    const rels = pkg.readRelationships(part);
    diagnostics.push(...rels.diagnostics);
    const rel = rels.entries.find((entry) => entry.Id === rId);
    if (rel === undefined) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part, xpath: hlinkClick.xpath, slideIndex },
        relationshipId: rId,
      });
      return undefined;
    }
    if (rel.Type === REL_HYPERLINK && rel.TargetMode === "External") {
      return { url: rel.Target };
    }
  }
  return undefined;
}

/**
 * Reads a Transitional `<p:graphicFrame>` chart and its `ppt/charts/chart{n}.xml` part into an editable `ChartNode`.
 */
export function readChartGraphicFrame(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, package: pkg } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const nvGraphicFramePr = cursor.child(P_NS, "nvGraphicFramePr");
  if (nvGraphicFramePr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part, xpath: cursor.xpath, slideIndex },
    });
    return { diagnostics };
  }

  let objectId: number | undefined;
  let objectName: string | undefined;
  let link: LinkTarget | undefined;
  const cNvPr = nvGraphicFramePr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    const rawCNvPr = cNvPr.attrsObject(["id", "name"]);
    const parsedCNvPr = NonVisualIdSchema.safeParse(rawCNvPr);
    if (!parsedCNvPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: cNvPr.xpath, slideIndex },
        attributeOrElement: "p:cNvPr",
        rawValue: JSON.stringify(rawCNvPr),
      });
    } else {
      objectId = parsedCNvPr.data.id;
      if (
        parsedCNvPr.data.name !== undefined &&
        parsedCNvPr.data.name.trim() !== ""
      ) {
        objectName = parsedCNvPr.data.name;
      }
    }
    link = readLinkFromCNvPr(cNvPr, part, slideIndex, pkg, diagnostics);
  }

  const cNvGraphicFramePr = nvGraphicFramePr.child(P_NS, "cNvGraphicFramePr");
  if (cNvGraphicFramePr !== undefined) {
    const locks = cNvGraphicFramePr.child(A_NS, "graphicFrameLocks");
    if (locks !== undefined) consumeSubtree(locks);
  }

  const nvPr = nvGraphicFramePr.child(P_NS, "nvPr");
  if (nvPr !== undefined) {
    const extLst = nvPr.child(P_NS, "extLst");
    if (extLst !== undefined) {
      for (const ext of extLst.children(P_NS, "ext")) {
        if (ext.attr("uri") === P14_MOD_ID_EXT_URI) {
          consumeSubtree(ext);
        }
      }
    }
  }

  const xfrm = cursor.child(P_NS, "xfrm");
  const off = xfrm?.child(A_NS, "off");
  const ext = xfrm?.child(A_NS, "ext");
  if (xfrm === undefined || off === undefined || ext === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part,
        xpath: cursor.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
    });
    return { diagnostics };
  }

  const rawXfrm = xfrm.attrsObject(["rot", "flipH", "flipV"]);
  TransformAttributesFragmentSchema.safeParse(rawXfrm);
  const parsedOff = OffsetEmuFragmentSchema.safeParse(
    off.attrsObject(["x", "y"]),
  );
  const parsedExt = ExtentEmuFragmentSchema.safeParse(
    ext.attrsObject(["cx", "cy"]),
  );
  if (!parsedOff.success || !parsedExt.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part,
        xpath: xfrm.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
      attributeOrElement: "a:xfrm",
    });
    return { diagnostics };
  }

  const boxX = emuToPx(parsedOff.data.x);
  const boxY = emuToPx(parsedOff.data.y);
  const boxWidth = emuToPx(parsedExt.data.cx);
  const boxHeight = emuToPx(parsedExt.data.cy);

  const graphic = cursor.child(A_NS, "graphic");
  const graphicData = graphic?.child(A_NS, "graphicData");
  if (graphicData !== undefined) {
    GraphicDataAttrSchema.safeParse(graphicData.attrsObject(["uri"]));
  }
  const cChart = graphicData?.child(C_NS, "chart");
  const rId = cChart?.attr("id", R_NS);
  if (cChart === undefined || rId === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part,
        xpath: cursor.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
    });
    return { diagnostics };
  }

  const relsResult = pkg.readRelationships(part);
  diagnostics.push(...relsResult.diagnostics);
  const rel = relsResult.entries.find(
    (entry) =>
      entry.Id === rId &&
      entry.Type === REL_CHART &&
      entry.TargetMode === "Internal",
  );
  if (rel === undefined) {
    diagnostics.push({
      code: "broken-relationship",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part,
        xpath: cChart.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
      relationshipId: rId,
    });
    return { diagnostics };
  }

  const chartPartPath = resolvePartPath(part, rel.Target);
  const chartXml =
    chartPartPath === undefined ? undefined : pkg.readTextPart(chartPartPath);
  if (chartPartPath === undefined || chartXml === undefined) {
    diagnostics.push({
      code: "missing-target-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part,
        xpath: cChart.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
      relationshipId: rId,
      targetPart: chartPartPath ?? rel.Target,
    });
    return { diagnostics };
  }

  let chartRootEl: ReturnType<typeof parseXmlPart>;
  try {
    chartRootEl = parseXmlPart(chartPartPath, chartXml);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part: chartPartPath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
    });
    return { diagnostics };
  }

  const chartSpace = XmlElementCursor.createRoot(
    chartRootEl,
    chartPartPath,
    "/c:chartSpace",
  );
  chartSpace.consumeAttr("lang", "http://www.w3.org/XML/1998/namespace");

  for (const tag of [
    "date1904",
    "lang",
    "roundedCorners",
    "style",
    "clrMapOvr",
  ]) {
    const el = chartSpace.child(C_NS, tag);
    if (el !== undefined) consumeSubtree(el);
  }
  for (const ac of chartSpace.children(MC_NS, "AlternateContent")) {
    consumeChartAlternateContent(ac);
  }
  const chartTxPr = chartSpace.child(C_NS, "txPr");
  const rawChartStyle =
    chartTxPr === undefined ? undefined : readChartTextProperties(chartTxPr);
  const chartStyle = normalizeRootChartStyle(rawChartStyle);
  for (const tag of ["spPr", "externalData"]) {
    const el = chartSpace.child(C_NS, tag);
    if (el !== undefined) consumeSubtree(el);
  }

  const chartEl = chartSpace.child(C_NS, "chart");
  if (chartEl === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part: chartPartPath,
        xpath: chartSpace.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
    });
    return { diagnostics };
  }

  const titleEl = chartEl.child(C_NS, "title");
  const title =
    titleEl === undefined ? undefined : readChartTitle(titleEl, rawChartStyle);
  const autoTitleDeleted = chartEl.child(C_NS, "autoTitleDeleted");
  if (autoTitleDeleted !== undefined) consumeSubtree(autoTitleDeleted);
  const pivotFmts = chartEl.child(C_NS, "pivotFmts");
  if (pivotFmts !== undefined) consumeSubtree(pivotFmts);

  let legend: ChartProps["legend"];
  const legendEl = chartEl.child(C_NS, "legend");
  if (legendEl !== undefined) {
    const posVal = legendEl.child(C_NS, "legendPos")?.attrsObject(["val"])[
      "val"
    ];
    legend =
      (posVal === undefined
        ? undefined
        : chartLegendPositionFromNative[posVal]) ?? "right";
    consumeSubtree(legendEl);
  }

  const plotVisOnly = chartEl.child(C_NS, "plotVisOnly");
  if (plotVisOnly !== undefined) consumeSubtree(plotVisOnly);

  const dispParsed = DispBlanksAsSchema.safeParse(
    chartEl.child(C_NS, "dispBlanksAs")?.attrsObject(["val"]) ?? {},
  );
  let displayBlanksAs: ChartDisplayBlanksAs | undefined = dispParsed.success
    ? dispParsed.data.val
    : undefined;

  const showDLblsOverMax = chartEl.child(C_NS, "showDLblsOverMax");
  if (showDLblsOverMax !== undefined) consumeSubtree(showDLblsOverMax);

  const chartExtLst = chartEl.child(C_NS, "extLst");
  if (chartExtLst !== undefined) {
    for (const extEntry of chartExtLst.children(C_NS, "ext")) {
      const uri = extEntry.attr("uri");
      if (
        uri === WPS_CHART_PROPS_EXT_URI ||
        directElementChild(extEntry.element, WPS_ET_NS, "chartProps") !==
          undefined
      ) {
        consumeSubtree(extEntry);
      }
    }
  }

  const plotAreaEl = chartEl.child(C_NS, "plotArea");
  if (plotAreaEl === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part: chartPartPath,
        xpath: chartEl.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
    });
    return { diagnostics };
  }

  let plotArea: ChartProps["plotArea"];
  const layoutEl = plotAreaEl.child(C_NS, "layout");
  if (layoutEl !== undefined) {
    const manual = layoutEl.child(C_NS, "manualLayout");
    if (manual !== undefined) {
      const target = readValAttr(manual.child(C_NS, "layoutTarget"));
      const xMode = readValAttr(manual.child(C_NS, "xMode"));
      const yMode = readValAttr(manual.child(C_NS, "yMode"));
      const wMode = readValAttr(manual.child(C_NS, "wMode"));
      const hMode = readValAttr(manual.child(C_NS, "hMode"));
      const xFrac = readFloatVal(manual.child(C_NS, "x"));
      const yFrac = readFloatVal(manual.child(C_NS, "y"));
      const wFrac = readFloatVal(manual.child(C_NS, "w"));
      const hFrac = readFloatVal(manual.child(C_NS, "h"));
      if (
        (target === undefined || target === "inner") &&
        (xMode === undefined || xMode === "edge") &&
        (yMode === undefined || yMode === "edge") &&
        (wMode === undefined || wMode === "factor") &&
        (hMode === undefined || hMode === "factor") &&
        xFrac !== undefined &&
        yFrac !== undefined &&
        wFrac !== undefined &&
        hFrac !== undefined &&
        wFrac > 0 &&
        hFrac > 0
      ) {
        const left = Math.max(0, xFrac * boxWidth);
        const top = Math.max(0, yFrac * boxHeight);
        const width = Math.min(boxWidth - left, wFrac * boxWidth);
        const height = Math.min(boxHeight - top, hFrac * boxHeight);
        if (width > 0 && height > 0) {
          plotArea = { left, top, width, height };
        }
      }
      consumeSubtree(manual);
    }
    consumeSubtree(layoutEl);
  }

  const plotSpPr = plotAreaEl.child(C_NS, "spPr");
  if (plotSpPr !== undefined) consumeSubtree(plotSpPr);

  const groupDrafts: ImportedGroupDraft[] = [];
  const importedAxes: ImportedAxisRecord[] = [];
  for (const child of plotAreaEl.remainingElementChildren()) {
    if (child.namespaceUri !== C_NS) continue;
    if (child.localName === "catAx") {
      const ax = readAxisElement(child, "category", rawChartStyle);
      if (ax !== undefined) importedAxes.push(ax);
      continue;
    }
    if (child.localName === "valAx") {
      const ax = readAxisElement(child, "value", rawChartStyle);
      if (ax !== undefined) importedAxes.push(ax);
      continue;
    }
    const group = readChartGroupElement(child, rawChartStyle);
    if (group !== undefined) {
      groupDrafts.push(group);
      continue;
    }
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: {
        part: chartPartPath,
        xpath: child.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
      namespaceUri: child.namespaceUri,
      localName: child.localName,
    });
  }

  for (const unconsumed of chartSpace.collectUnconsumed()) {
    diagnostics.push({
      code: "unconsumed-xml",
      category: "unsupported",
      disposition: "omitted",
      kind: unconsumed.kind,
      location: {
        ...unconsumed.location,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
      namespaceUri: unconsumed.namespaceUri,
      localName: unconsumed.localName,
    });
  }

  if (groupDrafts.length === 0) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part: chartPartPath,
        xpath: plotAreaEl.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
    });
    return { diagnostics };
  }

  const { axes, groups } = buildChartAxesAndGroups(groupDrafts, importedAxes);

  // Guard displayBlanksAs against chart kind / stacking constraints.
  if (displayBlanksAs !== undefined) {
    const supported = groups.every((group) => {
      if (group.kind === "scatter") {
        return !(
          displayBlanksAs === "span" &&
          (group.line === undefined || group.line === false)
        );
      }
      if (group.kind === "line") {
        return group.stacking === undefined ? true : displayBlanksAs === "zero";
      }
      if (group.kind === "area") {
        return group.stacking === undefined
          ? displayBlanksAs === "gap" || displayBlanksAs === "zero"
          : displayBlanksAs === "zero";
      }
      return displayBlanksAs === "gap" || displayBlanksAs === "zero";
    });
    if (!supported) {
      displayBlanksAs = undefined;
    }
  }

  const props: ChartProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    x: boxX,
    y: boxY,
    width: boxWidth,
    height: boxHeight,
    ...(axes === undefined ? {} : { axes }),
    groups,
    ...(title === undefined ? {} : { title }),
    ...(chartStyle === undefined ? {} : { style: chartStyle }),
    ...(displayBlanksAs === undefined ? {} : { displayBlanksAs }),
    ...(legend === undefined ? {} : { legend }),
    ...(plotArea === undefined ? {} : { plotArea }),
  };

  return {
    node: createChartNode(props),
    diagnostics,
  };
}

export const chartImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "graphicFrame",
  matches: isChartGraphicFrame,
  read: readChartGraphicFrame,
};
