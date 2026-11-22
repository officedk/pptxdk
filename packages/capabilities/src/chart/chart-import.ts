import type {
  ColorValue,
  LinkTarget,
  NativeColorSlot,
  ThemeShade,
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
import type {
  ChartAxes,
  ChartAxis,
  ChartAxisCrossing,
  ChartAxisPosition,
  ChartDataLabelPosition,
  ChartDataLabelSettings,
  ChartDataLabels,
  ChartDisplayBlanksAs,
  ChartGroup,
  ChartKind,
  ChartLine,
  ChartLineDash,
  ChartMarker,
  ChartPoint,
  ChartProps,
  ChartSeries,
  ChartTickLabelPosition,
  ChartTickMark,
  IndependentCategorySeries,
  ScatterChartPoint,
  ScatterChartSeries,
  ScatterDataLabels,
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
const AxPosSchema = z.object({ val: z.enum(["b", "t", "l", "r"]) });
const CrossesSchema = z.object({ val: z.enum(["autoZero", "max", "min"]) });
const TickMarkSchema = z.object({
  val: z.enum(["none", "in", "out", "cross"]),
});
const TickLblPosSchema = z.object({
  val: z.enum(["nextTo", "high", "low", "none"]),
});
const LegendPosSchema = z.object({ val: z.enum(["t", "b", "l", "r", "tr"]) });
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
const PresetDashSchema = z.object({
  val: z.enum([
    "solid",
    "dash",
    "dot",
    "dashDot",
    "lgDash",
    "lgDashDot",
    "lgDashDotDot",
  ]),
});

const axPosMap: Readonly<Record<"b" | "t" | "l" | "r", ChartAxisPosition>> = {
  b: "bottom",
  t: "top",
  l: "left",
  r: "right",
};

const legendPosMap: Readonly<
  Record<"t" | "b" | "l" | "r" | "tr", NonNullable<ChartProps["legend"]>>
> = {
  t: "top",
  b: "bottom",
  l: "left",
  r: "right",
  tr: "topRight",
};

const dLblPosMap: Readonly<
  Record<z.infer<typeof DLblPosSchema>["val"], ChartDataLabelPosition>
> = {
  ctr: "center",
  inEnd: "insideEnd",
  inBase: "insideBase",
  outEnd: "outsideEnd",
  bestFit: "bestFit",
  l: "left",
  r: "right",
  b: "bottom",
  t: "top",
};

const dashReverseMap: Readonly<
  Record<z.infer<typeof PresetDashSchema>["val"], ChartLineDash>
> = {
  solid: "solid",
  dash: "dash",
  dot: "dot",
  dashDot: "dashDot",
  lgDash: "longDash",
  lgDashDot: "longDashDot",
  lgDashDotDot: "longDashDotDot",
};

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

  let baseColor: string | ThemeShade | undefined;
  if (srgb !== undefined) {
    const raw = srgb.attrsObject(["val"]);
    const parsed = HexRgb6Schema.safeParse(raw["val"]);
    if (parsed.success) {
      baseColor = `#${parsed.data.toUpperCase()}`;
    }
  } else if (scheme !== undefined) {
    const raw = scheme.attrsObject(["val"]);
    const parsed = SchemeColorSchema.safeParse(raw);
    if (parsed.success) {
      baseColor = schemeToNativeSlot[parsed.data.val];
    }
  }
  const alphaCursor = colorCursor.child(A_NS, "alpha");
  let alpha: number | undefined;
  if (alphaCursor !== undefined) {
    const val = readIntVal(alphaCursor);
    if (val !== undefined && val >= 0 && val <= 100_000) {
      alpha = val / 100_000;
    }
  }
  consumeSubtree(colorCursor);
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
    const parsedDash = PresetDashSchema.safeParse(
      prstDash.attrsObject(["val"]),
    );
    if (parsedDash.success) {
      dash = dashReverseMap[parsedDash.data.val];
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

function extractRichTextString(cursor: XmlElementCursor): string {
  const parts: string[] = [];
  for (const p of cursor.children(A_NS, "p")) {
    const runs: string[] = [];
    for (const child of p.remainingElementChildren()) {
      if (child.namespaceUri === A_NS && child.localName === "r") {
        const t = child.child(A_NS, "t");
        if (t !== undefined) {
          runs.push(t.textContent);
        }
        consumeSubtree(child);
      } else if (child.namespaceUri === A_NS && child.localName === "br") {
        runs.push("\n");
        consumeSubtree(child);
      } else {
        consumeSubtree(child);
      }
    }
    parts.push(runs.join(""));
  }
  consumeSubtree(cursor);
  return parts.join("\n");
}

function readTitleText(titleCursor: XmlElementCursor): string | undefined {
  let titleText: string | undefined;
  const tx = titleCursor.child(C_NS, "tx");
  if (tx !== undefined) {
    const rich = tx.child(C_NS, "rich");
    if (rich !== undefined) {
      titleText = extractRichTextString(rich);
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
  consumeSubtree(titleCursor);
  if (titleText === undefined || titleText.trim() === "") return undefined;
  return titleText;
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
  readonly pointLabels: ReadonlyMap<number, false | ChartDataLabelSettings>;
  readonly showDataLabelsRange?: boolean;
};

function readDataLabelSettingsBody(
  cursor: XmlElementCursor,
  kind: ChartKind,
): {
  readonly settings: ChartDataLabelSettings;
  readonly showDataLabelsRange?: boolean;
} {
  const numFmtCursor = cursor.child(C_NS, "numFmt");
  let numberFormat: string | undefined;
  if (numFmtCursor !== undefined) {
    const code = numFmtCursor.attr("formatCode");
    const sourceLinked = readBoolVal(undefined);
    void sourceLinked;
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
  if (txPr !== undefined) consumeSubtree(txPr);

  const dLblPos = cursor.child(C_NS, "dLblPos");
  let position: ChartDataLabelPosition | undefined;
  if (dLblPos !== undefined) {
    const parsed = DLblPosSchema.safeParse(dLblPos.attrsObject(["val"]));
    if (parsed.success && kind !== "area" && kind !== "doughnut") {
      position = dLblPosMap[parsed.data.val];
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
        uri === "{0b15fc19-7d7d-44ad-8c2d-2c3a37ce22c3}"
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
    // Consume any remaining benign c15 extensions if extLst is now empty of unconsumed
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
  if ((kind === "pie" || kind === "doughnut") && hasVisibleContent) {
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
    ...(showDataLabelsRange === undefined ? {} : { showDataLabelsRange }),
  };
}

function readDataLabelsElement(
  dLblsCursor: XmlElementCursor,
  kind: ChartKind,
  isGroupLevel: boolean,
): ImportedDataLabelsResult {
  const pointLabels = new Map<number, false | ChartDataLabelSettings>();
  for (const dLbl of dLblsCursor.children(C_NS, "dLbl")) {
    const idx = readIntVal(dLbl.child(C_NS, "idx"));
    const del = readBoolVal(dLbl.child(C_NS, "delete"));
    if (idx !== undefined) {
      if (del === true) {
        pointLabels.set(idx, false);
      } else {
        const { settings } = readDataLabelSettingsBody(dLbl, kind);
        pointLabels.set(idx, settings);
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
    };
  }

  const { settings, showDataLabelsRange } = readDataLabelSettingsBody(
    dLblsCursor,
    kind,
  );

  // At group level, an all-false boilerplate <c:dLbls> (e.g. PowerPoint/WPS default)
  // represents no authored group dataLabels unless a visible flag or explicit format/position is set.
  const hasAnyEnabledOrCustomized =
    settings.value === true ||
    settings.percent === true ||
    settings.category === true ||
    settings.xValue === true ||
    settings.seriesName === true ||
    settings.separator !== undefined ||
    settings.numberFormat !== undefined ||
    settings.position !== undefined ||
    settings.leaderLines !== undefined;

  if (isGroupLevel && !hasAnyEnabledOrCustomized) {
    return { pointLabels };
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
    pointLabels,
    ...(showDataLabelsRange === undefined ? {} : { showDataLabelsRange }),
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
  readonly points?: Readonly<Record<number, ChartPoint>>;
  readonly datalabelsRange?: readonly string[];
};

function readCategorySeries(
  serCursor: XmlElementCursor,
  kind: Exclude<ChartKind, "scatter">,
  fallbackIndex: number,
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
      : readDataLabelsElement(dLblsCursor, kind, false);
  if (dLblsResult !== undefined) {
    for (const [ptIdx, ptLabel] of dLblsResult.pointLabels.entries()) {
      const existing = pointsRecord[ptIdx] ?? {};
      pointsRecord[ptIdx] = {
        ...existing,
        dataLabels: ptLabel,
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
    ...(Object.keys(normalizedPoints).length === 0
      ? {}
      : { points: normalizedPoints }),
    ...(datalabelsRange === undefined ? {} : { datalabelsRange }),
  };
}

function readScatterSeries(
  serCursor: XmlElementCursor,
  groupLine: ScatterLine | undefined,
  groupMarkers: boolean,
  fallbackIndex: number,
): { readonly order: number; readonly series: ScatterChartSeries } {
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
      : readDataLabelsElement(dLblsCursor, "scatter", false);
  if (dLblsResult !== undefined) {
    for (const [ptIdx, ptLabel] of dLblsResult.pointLabels.entries()) {
      const existing = pointsRecord[ptIdx] ?? {};
      pointsRecord[ptIdx] = {
        ...existing,
        dataLabels: ptLabel as false | ScatterDataLabels,
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
  readonly title?: string;
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

function readAxisElement(
  axisCursor: XmlElementCursor,
  kind: "category" | "value",
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
    const parsedPos = AxPosSchema.safeParse(axPos.attrsObject(["val"]));
    if (parsedPos.success) {
      position = axPosMap[parsedPos.data.val];
    }
  }

  const readGridline = (
    el: XmlElementCursor | undefined,
  ): boolean | ChartLine | undefined => {
    if (el === undefined) return undefined;
    const spPr = el.child(C_NS, "spPr");
    if (spPr === undefined) return true;
    const lineRes = readLineFromSpPr(spPr);
    consumeSubtree(spPr);
    return lineRes.line ?? true;
  };
  const gridlines = readGridline(axisCursor.child(C_NS, "majorGridlines"));
  const minorGridlines = readGridline(axisCursor.child(C_NS, "minorGridlines"));

  const titleCursor = axisCursor.child(C_NS, "title");
  const title =
    titleCursor === undefined ? undefined : readTitleText(titleCursor);

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

  const txPr = axisCursor.child(C_NS, "txPr");
  if (txPr !== undefined) consumeSubtree(txPr);

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
      readonly line?: ScatterLine;
      readonly markers: boolean;
    };

function readChartGroupElement(
  groupCursor: XmlElementCursor,
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

    const rawSeries = groupCursor
      .children(C_NS, "ser")
      .map((ser, idx) => readScatterSeries(ser, line, markers, idx))
      .toSorted((a, b) => a.order - b.order)
      .map((entry) => entry.series);

    const dLblsCursor = groupCursor.child(C_NS, "dLbls");
    const groupLabels =
      dLblsCursor === undefined
        ? undefined
        : readDataLabelsElement(dLblsCursor, "scatter", true).dataLabels;

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
    .map((ser, idx) => readCategorySeries(ser, kind, idx))
    .toSorted((a, b) => a.order - b.order);

  const dLblsCursor = groupCursor.child(C_NS, "dLbls");
  const groupLabels =
    dLblsCursor === undefined
      ? undefined
      : readDataLabelsElement(dLblsCursor, kind, true).dataLabels;

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

function buildChartAxesAndGroups(
  groupDrafts: readonly ImportedGroupDraft[],
  importedAxes: readonly ImportedAxisRecord[],
): { readonly axes?: ChartAxes; readonly groups: readonly ChartGroup[] } {
  const axisById = new Map(importedAxes.map((ax) => [ax.id, ax]));

  // Map native crosses/crossesAt back to author-facing crossing partners.
  // An OOXML crossing on axis A positions A in crossAx partner B's domain.
  const partnerCrossingFor = new Map<
    number,
    | { readonly kind: "crosses"; readonly value: ChartAxisCrossing }
    | { readonly kind: "crossesAt"; readonly value: number }
  >();
  for (const ax of importedAxes) {
    if (ax.rawCrossing === undefined || ax.crossAxisId === undefined) continue;
    if (ax.rawCrossing.kind === "crossesAt") {
      partnerCrossingFor.set(ax.crossAxisId, ax.rawCrossing);
    } else if (
      ax.rawCrossing.kind === "crosses" &&
      ax.rawCrossing.value !== "autoZero" &&
      ax.position === undefined
    ) {
      partnerCrossingFor.set(ax.crossAxisId, ax.rawCrossing);
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
    const common = {
      ...(ax.visible === undefined ? {} : { visible: ax.visible }),
      ...(ax.gridlines === undefined ? {} : { gridlines: ax.gridlines }),
      ...(ax.minorGridlines === undefined
        ? {}
        : { minorGridlines: ax.minorGridlines }),
      ...(ax.line === undefined ? {} : { line: ax.line }),
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
      ...(ax.position === undefined ? {} : { position: ax.position }),
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
        ...(group.dataLabels === undefined
          ? {}
          : { dataLabels: group.dataLabels }),
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
          };
        });
        const indDoughnutLabels = narrowGroupDataLabels(group.dataLabels, []);
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
      };
    });

    if (group.kind === "pie") {
      const pieLabels = narrowGroupDataLabels(group.dataLabels, [
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
      const doughnutLabels = narrowGroupDataLabels(group.dataLabels, []);
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
        const stackedLabels = narrowGroupDataLabels(group.dataLabels, [
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
      const unstackedLabels = narrowGroupDataLabels(group.dataLabels, [
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
      const lineLabels = narrowGroupDataLabels(group.dataLabels, [
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

    const areaLabels = narrowGroupDataLabels(group.dataLabels, []);
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
  for (const tag of ["spPr", "txPr", "externalData"]) {
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
  const title = titleEl === undefined ? undefined : readTitleText(titleEl);
  const autoTitleDeleted = chartEl.child(C_NS, "autoTitleDeleted");
  if (autoTitleDeleted !== undefined) consumeSubtree(autoTitleDeleted);
  const pivotFmts = chartEl.child(C_NS, "pivotFmts");
  if (pivotFmts !== undefined) consumeSubtree(pivotFmts);

  let legend: ChartProps["legend"];
  const legendEl = chartEl.child(C_NS, "legend");
  if (legendEl !== undefined) {
    const posParsed = LegendPosSchema.safeParse(
      legendEl.child(C_NS, "legendPos")?.attrsObject(["val"]) ?? {},
    );
    legend = posParsed.success ? legendPosMap[posParsed.data.val] : "right";
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
      const xFrac = readFloatVal(manual.child(C_NS, "x"));
      const yFrac = readFloatVal(manual.child(C_NS, "y"));
      const wFrac = readFloatVal(manual.child(C_NS, "w"));
      const hFrac = readFloatVal(manual.child(C_NS, "h"));
      if (
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
      const ax = readAxisElement(child, "category");
      if (ax !== undefined) importedAxes.push(ax);
      continue;
    }
    if (child.localName === "valAx") {
      const ax = readAxisElement(child, "value");
      if (ax !== undefined) importedAxes.push(ax);
      continue;
    }
    const group = readChartGroupElement(child);
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
