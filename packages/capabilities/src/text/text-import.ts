import {
  assertValidLang,
  schemeColorNameSchema,
  type ColorValue,
  type SchemeColorName,
  type Shadow,
  type TextAlignment,
  type TextStyle,
} from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";
import {
  A_NS,
  ExtentEmuFragmentSchema,
  HexRgb6Schema,
  OffsetEmuFragmentSchema,
  OoxmlBoolSchema,
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  OoxmlPositiveIntSchema,
  P_NS,
  TransformAttributesFragmentSchema,
  XML_NS,
  emuToPx,
  type ImportDiagnostic,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type TextInlineImportContext,
  type TextInlineImportHandler,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  consumeNonPlaceholderNvPr,
  hasPlaceholderElement,
} from "../shared/identity.ts";
import type { ParagraphStyle, TextList } from "../shared/text-schemas.ts";
import {
  createParagraphNode,
  createSlideNumberNode,
  createSpanNode,
  createTextNode,
  type SpanDeclarationProps,
  type TextDeclarationProps,
} from "./text-declaration.ts";

const NonVisualSpIdSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const NonVisualSpPrFlagsSchema = z.object({
  txBox: OoxmlBoolSchema.optional(),
});

const PresetGeomFragmentSchema = z.object({
  prst: z.string().min(1),
});

const BodyPrFragmentSchema = z.object({
  wrap: z.enum(["square", "none"]).optional(),
  rtlCol: OoxmlBoolSchema.optional(),
  lIns: OoxmlNonNegativeIntSchema.optional(),
  tIns: OoxmlNonNegativeIntSchema.optional(),
  rIns: OoxmlNonNegativeIntSchema.optional(),
  bIns: OoxmlNonNegativeIntSchema.optional(),
  anchor: z.enum(["t", "ctr", "b"]).optional(),
});

const ParagraphPrFragmentSchema = z.object({
  algn: z.enum(["l", "ctr", "r", "just"]).optional(),
  rtl: OoxmlBoolSchema.optional(),
  lvl: OoxmlNonNegativeIntSchema.refine((value) => value <= 8).optional(),
});

const SpacingPointsXmlSchema = z.object({
  val: OoxmlNonNegativeIntSchema,
});

const SpacingPercentXmlSchema = z.object({
  val: OoxmlPositiveIntSchema,
});

const BulletCharXmlSchema = z.object({
  char: z.string().min(1),
});

const AutoNumXmlSchema = z.object({
  type: z.string().min(1),
  startAt: OoxmlPositiveIntSchema.optional(),
});

const RunPrFragmentSchema = z.object({
  lang: z.string().min(1).optional(),
  sz: OoxmlPositiveIntSchema.optional(),
  b: OoxmlBoolSchema.optional(),
  i: OoxmlBoolSchema.optional(),
  u: z.enum(["none", "sng"]).optional(),
  strike: z.enum(["noStrike", "sngStrike"]).optional(),
  cap: z.enum(["none", "all"]).optional(),
  spc: OoxmlIntSchema.optional(),
  dirty: OoxmlBoolSchema.optional(),
  smtClean: OoxmlBoolSchema.optional(),
});

const TextFieldXmlSchema = z.object({
  id: z.string().min(1),
  type: z.string().min(1),
});

const OuterShadowXmlSchema = z.object({
  blurRad: OoxmlNonNegativeIntSchema.optional(),
  dist: OoxmlNonNegativeIntSchema.optional(),
  dir: OoxmlIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: OoxmlBoolSchema.optional(),
});

const TypefaceFragmentSchema = z.object({
  typeface: z.string().min(1),
  pitchFamily: z.string().optional(),
  charset: z.string().optional(),
  panose: z.string().optional(),
});

const PercentModifierSchema = z.object({
  val: OoxmlNonNegativeIntSchema,
});

const SrgbColorSchema = z.object({
  val: HexRgb6Schema,
});

const SchemeColorSchema = z.object({
  val: schemeColorNameSchema,
});

const anchorToVerticalAlign: Readonly<
  Record<"t" | "ctr" | "b", NonNullable<TextStyle["verticalAlign"]>>
> = {
  t: "top",
  ctr: "middle",
  b: "bottom",
};

const algnToTextAlign: Readonly<
  Record<"l" | "ctr" | "r" | "just", TextAlignment>
> = {
  l: "left",
  ctr: "center",
  r: "right",
  just: "justify",
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

/** Side-effect-free predicate distinguishing `<p:sp>` text boxes (`txBox="1"`) from shapes and placeholders. */
function isTextBoxSpElement(cursor: XmlElementCursor): boolean {
  if (hasPlaceholderElement(cursor)) return false;
  const nvSpPr = directElementChild(cursor.element, P_NS, "nvSpPr");
  if (nvSpPr === undefined) return false;
  const cNvSpPr = directElementChild(nvSpPr, P_NS, "cNvSpPr");
  if (cNvSpPr === undefined) return false;
  const txBox = cNvSpPr.getAttribute("txBox");
  return txBox === "1" || txBox === "true";
}

function readColorModifier(
  colorCursor: XmlElementCursor,
  localName: "alpha" | "lumMod" | "lumOff",
): number | undefined {
  const mod = colorCursor.child(A_NS, localName);
  if (mod === undefined) return undefined;
  const parsed = PercentModifierSchema.safeParse(mod.attrsObject(["val"]));
  return parsed.success ? parsed.data.val / 100_000 : undefined;
}

function readColorElement(
  containerCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): ColorValue | undefined {
  const srgb = containerCursor.child(A_NS, "srgbClr");
  const scheme =
    srgb === undefined ? containerCursor.child(A_NS, "schemeClr") : undefined;
  const colorCursor = srgb ?? scheme;
  if (colorCursor === undefined) return undefined;

  let baseString: string | undefined;
  let schemeSlot: SchemeColorName | undefined;

  if (srgb !== undefined) {
    const raw = srgb.attrsObject(["val"]);
    const parsed = SrgbColorSchema.safeParse(raw);
    if (!parsed.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part, xpath: srgb.xpath, slideIndex },
        attributeOrElement: "a:srgbClr/@val",
        rawValue: raw["val"] ?? "",
      });
      return undefined;
    }
    baseString = `#${parsed.data.val.toUpperCase()}`;
  } else if (scheme !== undefined) {
    const raw = scheme.attrsObject(["val"]);
    const parsed = SchemeColorSchema.safeParse(raw);
    if (!parsed.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part, xpath: scheme.xpath, slideIndex },
        attributeOrElement: "a:schemeClr/@val",
        rawValue: raw["val"] ?? "",
      });
      return undefined;
    }
    schemeSlot = parsed.data.val;
  }

  const lumMod = readColorModifier(colorCursor, "lumMod");
  const lumOff = readColorModifier(colorCursor, "lumOff");
  const alpha = readColorModifier(colorCursor, "alpha");

  const baseColor: ColorValue | undefined =
    schemeSlot !== undefined && lumMod !== undefined
      ? {
          slot: schemeSlot,
          lumMod: Math.min(1, lumMod),
          ...(lumOff === undefined ? {} : { lumOff: Math.min(1, lumOff) }),
        }
      : (schemeSlot ?? baseString);

  if (baseColor === undefined) return undefined;
  if (alpha !== undefined) {
    return { color: baseColor, alpha: Math.min(1, alpha) };
  }
  return baseColor;
}

function readOuterShadow(
  rPrCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): Shadow | undefined {
  const effectLst = rPrCursor.child(A_NS, "effectLst");
  if (effectLst === undefined) return undefined;
  const outerShdw = effectLst.child(A_NS, "outerShdw");
  if (outerShdw === undefined) return undefined;

  const raw = outerShdw.attrsObject([
    "blurRad",
    "dist",
    "dir",
    "algn",
    "rotWithShape",
  ]);
  const parsed = OuterShadowXmlSchema.safeParse(raw);
  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part, xpath: outerShdw.xpath, slideIndex },
      attributeOrElement: "a:outerShdw",
      rawValue: JSON.stringify(raw),
    });
    return undefined;
  }

  const color = readColorElement(outerShdw, part, slideIndex, diagnostics);
  const blur =
    parsed.data.blurRad === undefined
      ? undefined
      : emuToPx(parsed.data.blurRad);
  const distance =
    parsed.data.dist === undefined ? undefined : emuToPx(parsed.data.dist);
  const angle =
    parsed.data.dir === undefined ? undefined : parsed.data.dir / 60_000;

  return {
    ...(color === undefined ? {} : { color }),
    ...(blur === undefined ? {} : { blur }),
    ...(distance === undefined ? {} : { distance }),
    ...(angle === undefined ? {} : { angle }),
  };
}

function resolveTypefaceToken(
  typeface: string,
): NonNullable<TextStyle["fontFamily"]> {
  if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
    return "minor";
  }
  if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
    return "major";
  }
  return typeface;
}

export type ParsedRunStyle = NonNullable<SpanDeclarationProps["style"]>;

function readRunProperties(
  rPrCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): { readonly style: ParsedRunStyle; readonly sizePt?: number } {
  const raw = rPrCursor.attrsObject([
    "lang",
    "sz",
    "b",
    "i",
    "u",
    "strike",
    "cap",
    "spc",
    "dirty",
    "smtClean",
  ]);
  const parsed = RunPrFragmentSchema.safeParse(raw);
  let lang: string | undefined;
  let fontSize: NonNullable<TextStyle["fontSize"]> | undefined;
  let sizePt: number | undefined;
  let fontWeight: NonNullable<TextStyle["fontWeight"]> | undefined;
  let fontStyle: NonNullable<TextStyle["fontStyle"]> | undefined;
  let textDecoration: NonNullable<TextStyle["textDecoration"]> | undefined;
  let textTransform: NonNullable<TextStyle["textTransform"]> | undefined;
  let letterSpacing: number | undefined;

  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { part, xpath: rPrCursor.xpath, slideIndex },
      attributeOrElement: "a:rPr",
      rawValue: JSON.stringify(raw),
    });
  } else {
    const data = parsed.data;
    if (data.lang !== undefined) {
      try {
        lang = assertValidLang(data.lang);
      } catch {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part, xpath: `${rPrCursor.xpath}/@lang`, slideIndex },
          attributeOrElement: "@lang",
          rawValue: data.lang,
        });
      }
    }
    if (data.sz !== undefined) {
      sizePt = data.sz / 100;
      fontSize = `${sizePt}pt`;
    }
    if (data.b !== undefined) {
      fontWeight = data.b ? "bold" : "normal";
    }
    if (data.i !== undefined) {
      fontStyle = data.i ? "italic" : "normal";
    }
    if (data.u === "sng") {
      textDecoration = "underline";
    } else if (data.strike === "sngStrike") {
      textDecoration = "line-through";
    } else if (data.u === "none" && data.strike === "noStrike") {
      textDecoration = "none";
    }
    if (data.cap === "all") {
      textTransform = "uppercase";
    } else if (data.cap === "none") {
      textTransform = "none";
    }
    if (data.spc !== undefined) {
      letterSpacing = data.spc / 75;
    }
  }

  let color: NonNullable<TextStyle["color"]> | undefined;
  const solidFill = rPrCursor.child(A_NS, "solidFill");
  if (solidFill !== undefined) {
    color = readColorElement(solidFill, part, slideIndex, diagnostics);
  }

  const textShadow = readOuterShadow(rPrCursor, part, slideIndex, diagnostics);

  let fontFamily: NonNullable<TextStyle["fontFamily"]> | undefined;
  const latin = rPrCursor.child(A_NS, "latin");
  const ea = rPrCursor.child(A_NS, "ea");
  const cs = rPrCursor.child(A_NS, "cs");
  const fontCursor = latin ?? ea ?? cs;
  if (fontCursor !== undefined) {
    const rawFont = fontCursor.attrsObject([
      "typeface",
      "pitchFamily",
      "charset",
      "panose",
    ]);
    const parsedFont = TypefaceFragmentSchema.safeParse(rawFont);
    if (parsedFont.success) {
      fontFamily = resolveTypefaceToken(parsedFont.data.typeface);
    }
  }
  for (const extraSlot of [ea, cs]) {
    if (extraSlot !== undefined && extraSlot !== fontCursor) {
      extraSlot.attrsObject(["typeface", "pitchFamily", "charset", "panose"]);
    }
  }

  return {
    style: {
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
    },
    ...(sizePt === undefined ? {} : { sizePt }),
  };
}

/**
 * Import handler for `<a:fld type="slidenum">` inline text fields (`text-field-slidenum`).
 */
export const textFieldSlideNumImportHandler: TextInlineImportHandler = {
  target: "text-inline",
  localName: "fld",
  matches(cursor: XmlElementCursor): boolean {
    return cursor.element.getAttribute("type") === "slidenum";
  },
  read(context: TextInlineImportContext) {
    const { cursor, part, slideIndex } = context;
    const diagnostics: ImportDiagnostic[] = [];
    const rawField = cursor.attrsObject(["id", "type"]);
    const parsedField = TextFieldXmlSchema.safeParse(rawField);
    const rPrCursor = cursor.child(A_NS, "rPr");
    const tCursor = cursor.child(A_NS, "t");
    if (tCursor !== undefined) {
      tCursor.consumeAttr("space", XML_NS);
    }
    const fallbackText = tCursor?.textContent ?? "";

    if (!parsedField.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part, xpath: cursor.xpath, slideIndex },
        attributeOrElement: "a:fld",
        rawValue: JSON.stringify(rawField),
      });
      return {
        kind: "unsupported",
        fieldType: rawField["type"] ?? "",
        ...(rPrCursor === undefined ? {} : { rPrCursor }),
        fallbackText,
        diagnostics,
      };
    }

    if (parsedField.data.type !== "slidenum") {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { part, xpath: cursor.xpath, slideIndex },
        namespaceUri: cursor.namespaceUri,
        localName: cursor.localName,
        mechanismHint: `text-field:${parsedField.data.type}`,
      });
      return {
        kind: "unsupported",
        fieldType: parsedField.data.type,
        ...(rPrCursor === undefined ? {} : { rPrCursor }),
        fallbackText,
        diagnostics,
      };
    }

    return {
      kind: "slide-number",
      fieldType: "slidenum",
      ...(rPrCursor === undefined ? {} : { rPrCursor }),
      fallbackText,
      diagnostics,
    };
  },
};

type ImportedRun =
  | {
      readonly kind: "text";
      readonly text: string;
      readonly style: ParsedRunStyle;
      readonly sizePt?: number;
    }
  | {
      readonly kind: "slide-number";
      readonly style: ParsedRunStyle;
      readonly sizePt?: number;
    };

type ImportedParagraph = {
  readonly runs: readonly ImportedRun[];
  readonly style: ParagraphStyle;
};

function readParagraph(
  pCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): ImportedParagraph {
  const pPr = pCursor.child(A_NS, "pPr");

  let align: TextAlignment | undefined;
  let lvl: number | undefined;
  let list: TextList | undefined;
  let lnSpcNode: XmlElementCursor | undefined;
  let spcAftNode: XmlElementCursor | undefined;

  if (pPr !== undefined) {
    const rawPPr = pPr.attrsObject(["algn", "rtl", "lvl"]);
    const parsedPPr = ParagraphPrFragmentSchema.safeParse(rawPPr);
    if (!parsedPPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: pPr.xpath, slideIndex },
        attributeOrElement: "a:pPr",
        rawValue: JSON.stringify(rawPPr),
      });
    } else {
      if (parsedPPr.data.algn !== undefined) {
        align = algnToTextAlign[parsedPPr.data.algn];
      }
      lvl = parsedPPr.data.lvl;
    }

    lnSpcNode = pPr.child(A_NS, "lnSpc");
    spcAftNode = pPr.child(A_NS, "spcAft");

    pPr.ignoreEmptyChild(A_NS, "buNone");
    const buChar = pPr.child(A_NS, "buChar");
    const buAutoNum = pPr.child(A_NS, "buAutoNum");
    if (buChar !== undefined) {
      const rawBu = buChar.attrsObject(["char"]);
      const parsedBu = BulletCharXmlSchema.safeParse(rawBu);
      if (!parsedBu.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part, xpath: buChar.xpath, slideIndex },
          attributeOrElement: "a:buChar/@char",
          rawValue: rawBu["char"] ?? "",
        });
      } else {
        list = {
          type: "bullet",
          character: parsedBu.data.char,
          ...(lvl === undefined ? {} : { level: lvl }),
        };
      }
    } else if (buAutoNum !== undefined) {
      const rawNum = buAutoNum.attrsObject(["type", "startAt"]);
      const parsedNum = AutoNumXmlSchema.safeParse(rawNum);
      if (!parsedNum.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part, xpath: buAutoNum.xpath, slideIndex },
          attributeOrElement: "a:buAutoNum",
          rawValue: JSON.stringify(rawNum),
        });
      } else {
        if (parsedNum.data.type !== "arabicPeriod") {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "recovered",
            location: {
              part,
              xpath: `${buAutoNum.xpath}/@type`,
              slideIndex,
            },
            attributeOrElement: "a:buAutoNum/@type",
            rawValue: parsedNum.data.type,
          });
        }
        list = {
          type: "number",
          style: "arabicPeriod",
          ...(parsedNum.data.startAt === undefined
            ? {}
            : { startAt: parsedNum.data.startAt }),
          ...(lvl === undefined ? {} : { level: lvl }),
        };
      }
    }
  }

  const runs: ImportedRun[] = [];
  let endParaSizePt: number | undefined;

  for (const child of pCursor.remainingElementChildren()) {
    if (child.namespaceUri === A_NS && child.localName === "r") {
      const rPr = child.child(A_NS, "rPr");
      const parsedRPr =
        rPr === undefined
          ? { style: {} }
          : readRunProperties(rPr, part, slideIndex, diagnostics);
      const t = child.child(A_NS, "t");
      if (t !== undefined) {
        t.consumeAttr("space", XML_NS);
        runs.push({
          kind: "text",
          text: t.textContent,
          style: parsedRPr.style,
          ...(parsedRPr.sizePt === undefined
            ? {}
            : { sizePt: parsedRPr.sizePt }),
        });
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "br") {
      const rPr = child.child(A_NS, "rPr");
      const parsedRPr =
        rPr === undefined
          ? { style: {} }
          : readRunProperties(rPr, part, slideIndex, diagnostics);
      runs.push({
        kind: "text",
        text: "\n",
        style: parsedRPr.style,
        ...(parsedRPr.sizePt === undefined ? {} : { sizePt: parsedRPr.sizePt }),
      });
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "fld") {
      const outcome = textFieldSlideNumImportHandler.read({
        cursor: child,
        part,
        slideIndex,
      });
      diagnostics.push(...outcome.diagnostics);
      if (outcome.kind === "slide-number") {
        const parsedRPr =
          outcome.rPrCursor === undefined
            ? { style: {} }
            : readRunProperties(
                outcome.rPrCursor,
                part,
                slideIndex,
                diagnostics,
              );
        runs.push({
          kind: "slide-number",
          style: parsedRPr.style,
          ...(parsedRPr.sizePt === undefined
            ? {}
            : { sizePt: parsedRPr.sizePt }),
        });
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "endParaRPr") {
      const parsedEnd = readRunProperties(child, part, slideIndex, diagnostics);
      endParaSizePt = parsedEnd.sizePt;
      continue;
    }
    // Any other child under `<a:p>` becomes an explicit unconsumed-xml item.
    diagnostics.push({
      code: "unconsumed-xml",
      category: "unsupported",
      disposition: "omitted",
      kind: "element",
      location: { part, xpath: child.xpath, slideIndex },
      namespaceUri: child.namespaceUri,
      localName: child.localName,
    });
  }

  const baseFontPt =
    runs.find((run) => run.sizePt !== undefined)?.sizePt ?? endParaSizePt ?? 18;

  let lineHeight: ParagraphStyle["lineHeight"] | undefined;
  if (lnSpcNode !== undefined) {
    const spcPts = lnSpcNode.child(A_NS, "spcPts");
    const spcPct =
      spcPts === undefined ? lnSpcNode.child(A_NS, "spcPct") : undefined;
    if (spcPts !== undefined) {
      const raw = spcPts.attrsObject(["val"]);
      const parsed = SpacingPointsXmlSchema.safeParse(raw);
      if (!parsed.success || parsed.data.val === 0) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part, xpath: spcPts.xpath, slideIndex },
          attributeOrElement: "a:spcPts/@val",
          rawValue: raw["val"] ?? "",
        });
      } else {
        lineHeight = `${parsed.data.val / 100}pt`;
      }
    } else if (spcPct !== undefined) {
      const raw = spcPct.attrsObject(["val"]);
      const parsed = SpacingPercentXmlSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part, xpath: spcPct.xpath, slideIndex },
          attributeOrElement: "a:spcPct/@val",
          rawValue: raw["val"] ?? "",
        });
      } else {
        const pt = Number(
          ((baseFontPt * parsed.data.val) / 100_000).toFixed(2),
        );
        if (pt > 0) {
          lineHeight = `${pt}pt`;
        }
      }
    }
  }

  let spacingAfter: number | undefined;
  if (spcAftNode !== undefined) {
    const spcPts = spcAftNode.child(A_NS, "spcPts");
    const spcPct =
      spcPts === undefined ? spcAftNode.child(A_NS, "spcPct") : undefined;
    if (spcPts !== undefined) {
      const raw = spcPts.attrsObject(["val"]);
      const parsed = SpacingPointsXmlSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part, xpath: spcPts.xpath, slideIndex },
          attributeOrElement: "a:spcPts/@val",
          rawValue: raw["val"] ?? "",
        });
      } else {
        spacingAfter = parsed.data.val / 75;
      }
    } else if (spcPct !== undefined) {
      const raw = spcPct.attrsObject(["val"]);
      const parsed = SpacingPercentXmlSchema.safeParse(raw);
      if (parsed.success) {
        spacingAfter = ((baseFontPt * parsed.data.val) / 100_000) * (4 / 3);
      }
    }
  }

  const style: ParagraphStyle = {
    ...(align === undefined ? {} : { align }),
    ...(lineHeight === undefined ? {} : { lineHeight }),
    ...(spacingAfter === undefined ? {} : { spacingAfter }),
    ...(list === undefined ? {} : { list }),
  };

  return { runs, style };
}

export type ImportedRichTextBody = {
  readonly boxStyle: NonNullable<TextDeclarationProps["style"]>;
  readonly importedParagraphs: readonly ImportedParagraph[];
};

/**
 * Shared rich-text body reader for text boxes, shapes, table cells, and speaker notes.
 */
export function readRichTextBody(
  txBody: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): ImportedRichTextBody {
  let paddingTop: number | undefined;
  let paddingRight: number | undefined;
  let paddingBottom: number | undefined;
  let paddingLeft: number | undefined;
  let verticalAlign: NonNullable<TextStyle["verticalAlign"]> | undefined;
  let textWrap: NonNullable<TextStyle["textWrap"]> | undefined;

  const bodyPr = txBody.child(A_NS, "bodyPr");
  if (bodyPr !== undefined) {
    const rawBodyPr = bodyPr.attrsObject([
      "wrap",
      "rtlCol",
      "lIns",
      "tIns",
      "rIns",
      "bIns",
      "anchor",
    ]);
    const parsedBodyPr = BodyPrFragmentSchema.safeParse(rawBodyPr);
    if (!parsedBodyPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: bodyPr.xpath, slideIndex },
        attributeOrElement: "a:bodyPr",
        rawValue: JSON.stringify(rawBodyPr),
      });
    } else {
      const b = parsedBodyPr.data;
      if (b.wrap === "none") textWrap = "nowrap";
      else if (b.wrap === "square") textWrap = "normal";
      if (b.anchor !== undefined) {
        verticalAlign = anchorToVerticalAlign[b.anchor];
      }
      if (b.lIns !== undefined && b.lIns !== 0) paddingLeft = emuToPx(b.lIns);
      if (b.tIns !== undefined && b.tIns !== 0) paddingTop = emuToPx(b.tIns);
      if (b.rIns !== undefined && b.rIns !== 0) paddingRight = emuToPx(b.rIns);
      if (b.bIns !== undefined && b.bIns !== 0) paddingBottom = emuToPx(b.bIns);
    }
    bodyPr.ignoreEmptyChild(A_NS, "noAutofit");
    bodyPr.ignoreEmptyChild(A_NS, "spAutoFit");
  }

  txBody.ignoreEmptyChild(A_NS, "lstStyle");

  const importedParagraphs = txBody
    .children(A_NS, "p")
    .map((pCursor) => readParagraph(pCursor, part, slideIndex, diagnostics));

  const boxStyle: NonNullable<TextDeclarationProps["style"]> = {
    ...(paddingTop === undefined ? {} : { paddingTop }),
    ...(paddingRight === undefined ? {} : { paddingRight }),
    ...(paddingBottom === undefined ? {} : { paddingBottom }),
    ...(paddingLeft === undefined ? {} : { paddingLeft }),
    ...(verticalAlign === undefined ? {} : { verticalAlign }),
    ...(textWrap === undefined ? {} : { textWrap }),
  };

  return { boxStyle, importedParagraphs };
}

/** Populates a `RichTextContainer` (`TextNode`, `ShapeNode`, `CellNode`, `NotesNode`) from imported paragraphs. */
export function populateRichTextNode(
  target: {
    content(...items: (string | CommonNode)[]): unknown;
  },
  importedParagraphs: readonly ImportedParagraph[],
): void {
  const paragraphNodes = importedParagraphs.map((para) => {
    const pNode = createParagraphNode(
      Object.keys(para.style).length === 0 ? {} : { style: para.style },
    );
    if (
      para.runs.length === 1 &&
      para.runs[0]!.kind === "text" &&
      Object.keys(para.runs[0]!.style).length === 0
    ) {
      pNode.content(para.runs[0]!.text);
    } else if (para.runs.length > 0) {
      pNode.content(
        ...para.runs.map((run) => {
          if (run.kind === "slide-number") {
            return createSlideNumberNode(
              Object.keys(run.style).length === 0 ? {} : { style: run.style },
            );
          }
          return Object.keys(run.style).length === 0
            ? run.text
            : createSpanNode({ style: run.style }).content(run.text);
        }),
      );
    }
    return pNode;
  });
  target.content(...paragraphNodes);
}

/**
 * Reads a Transitional `<p:sp>` text box (`<p:cNvSpPr txBox="1"/>`) into an editable `TextNode`.
 */
export function readTextBoxSp(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const nvSpPr = cursor.child(P_NS, "nvSpPr");
  if (nvSpPr === undefined) {
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
  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    const rawCNvPr = cNvPr.attrsObject(["id", "name"]);
    const parsedCNvPr = NonVisualSpIdSchema.safeParse(rawCNvPr);
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
  }

  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    const rawFlags = cNvSpPr.attrsObject(["txBox"]);
    NonVisualSpPrFlagsSchema.safeParse(rawFlags);
  }
  consumeNonPlaceholderNvPr(nvSpPr);

  const spPr = cursor.child(P_NS, "spPr");
  const txBody = cursor.child(P_NS, "txBody");
  if (spPr === undefined || txBody === undefined) {
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

  const xfrm = spPr.child(A_NS, "xfrm");
  const off = xfrm?.child(A_NS, "off");
  const ext = xfrm?.child(A_NS, "ext");
  if (xfrm === undefined || off === undefined || ext === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part,
        xpath: spPr.xpath,
        slideIndex,
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      },
    });
    return { diagnostics };
  }

  const rawXfrm = xfrm.attrsObject(["rot", "flipH", "flipV"]);
  const parsedXfrm = TransformAttributesFragmentSchema.safeParse(rawXfrm);
  const rawOff = off.attrsObject(["x", "y"]);
  const parsedOff = OffsetEmuFragmentSchema.safeParse(rawOff);
  const rawExt = ext.attrsObject(["cx", "cy"]);
  const parsedExt = ExtentEmuFragmentSchema.safeParse(rawExt);

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

  const prstGeom = spPr.child(A_NS, "prstGeom");
  if (prstGeom !== undefined) {
    const rawGeom = prstGeom.attrsObject(["prst"]);
    const parsedGeom = PresetGeomFragmentSchema.safeParse(rawGeom);
    if (!parsedGeom.success || parsedGeom.data.prst !== "rect") {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: prstGeom.xpath, slideIndex },
        attributeOrElement: "a:prstGeom/@prst",
        rawValue: rawGeom["prst"] ?? "",
      });
    }
    prstGeom.ignoreEmptyChild(A_NS, "avLst");
  }

  spPr.ignoreEmptyChild(A_NS, "noFill");
  const ln = spPr.child(A_NS, "ln");
  if (ln !== undefined) {
    ln.ignoreEmptyChild(A_NS, "noFill");
  }
  let backgroundColor: ColorValue | undefined;
  const spSolidFill = spPr.child(A_NS, "solidFill");
  if (spSolidFill !== undefined) {
    backgroundColor = readColorElement(
      spSolidFill,
      part,
      slideIndex,
      diagnostics,
    );
  }

  const rotate =
    parsedXfrm.success &&
    parsedXfrm.data.rot !== undefined &&
    parsedXfrm.data.rot !== 0
      ? parsedXfrm.data.rot / 60_000
      : undefined;

  const { boxStyle: bodyStyle, importedParagraphs } = readRichTextBody(
    txBody,
    part,
    slideIndex,
    diagnostics,
  );

  const singleParagraph =
    importedParagraphs.length === 1 ? importedParagraphs[0] : undefined;
  const singleTextRun =
    singleParagraph !== undefined &&
    singleParagraph.style.list === undefined &&
    singleParagraph.runs.length === 1 &&
    singleParagraph.runs[0]!.kind === "text"
      ? singleParagraph.runs[0]
      : undefined;

  const flatParagraphStyle: Omit<ParagraphStyle, "list"> | undefined =
    singleTextRun === undefined || singleParagraph === undefined
      ? undefined
      : (() => {
          const { list: _list, ...rest } = singleParagraph.style;
          return rest;
        })();

  const boxStyle: NonNullable<TextDeclarationProps["style"]> = {
    ...bodyStyle,
    ...(rotate === undefined ? {} : { rotate }),
    ...(backgroundColor === undefined ? {} : { backgroundColor }),
    ...flatParagraphStyle,
    ...(singleTextRun === undefined ? {} : singleTextRun.style),
  };

  const props: TextDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    x: emuToPx(parsedOff.data.x),
    y: emuToPx(parsedOff.data.y),
    width: emuToPx(parsedExt.data.cx),
    height: emuToPx(parsedExt.data.cy),
    ...(Object.keys(boxStyle).length === 0 ? {} : { style: boxStyle }),
  };

  const node = createTextNode(props);
  if (singleTextRun !== undefined) {
    node.content(singleTextRun.text);
  } else if (importedParagraphs.length > 0) {
    populateRichTextNode(node, importedParagraphs);
  }

  return { node, diagnostics };
}

export const textBoxImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: isTextBoxSpElement,
  read: readTextBoxSp,
};
