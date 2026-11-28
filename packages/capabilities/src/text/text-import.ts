import {
  assertValidLang,
  matchLayoutPlaceholderKey,
  matchMasterPlaceholderSpec,
  placeholderStyleSchema,
  schemeColorNameSchema,
  type ColorValue,
  type HexColor,
  type NativeColorSlot,
  type NativePlaceholderRef,
  type SchemeColorName,
  type Shadow,
  type TextAlignment,
  type TextStyle,
  type VisualStyle,
} from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";
import {
  A_NS,
  EMU_PER_PX,
  extentEmuFragmentSchema,
  hexRgb6Schema,
  offsetEmuFragmentSchema,
  ooxmlBoolSchema,
  ooxmlIntSchema,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
  P_NS,
  readRunFontSlots,
  transformAttributesFragmentSchema,
  XML_NS,
  emuToPx,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type TextInlineImportContext,
  type TextInlineImportHandler,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import { readNonVisualCommonProperties } from "../shape/shape-import.ts";
import type { Geometry } from "../shared/geometry.ts";
import {
  attachDeferredSlideLink,
  consumeNonPlaceholderNvPr,
  hasPlaceholderElement,
} from "../shared/identity.ts";
import {
  directElementChild,
  directPlaceholderElement,
  isPlaceholderSpElement,
  readNvPrPlaceholder,
} from "../shared/placeholder-import.ts";
import type { ParagraphStyle, TextList } from "../shared/text-schemas.ts";
import {
  readSpPrAppearance,
  readSpPrVisuals,
  type LineEndStyle,
} from "../shared/visual.ts";
import {
  createParagraphNode,
  createSlideNumberNode,
  createSpanNode,
  createSubNode,
  createSupNode,
  createTextNode,
  type SpanDeclarationProps,
  type TextDeclarationProps,
} from "./text-declaration.ts";

const NonVisualSpPrFlagsSchema = z.object({
  txBox: ooxmlBoolSchema.optional(),
});

const BodyPrFragmentSchema = z.object({
  wrap: z.enum(["square", "none"]).optional(),
  rtlCol: ooxmlBoolSchema.optional(),
  lIns: ooxmlNonNegativeIntSchema.optional(),
  tIns: ooxmlNonNegativeIntSchema.optional(),
  rIns: ooxmlNonNegativeIntSchema.optional(),
  bIns: ooxmlNonNegativeIntSchema.optional(),
  anchor: z.enum(["t", "ctr", "b"]).optional(),
  anchorCtr: ooxmlBoolSchema.optional(),
  spcFirstLastPara: ooxmlBoolSchema.optional(),
  numCol: z.string().optional(),
  spcCol: z.string().optional(),
});

const ParagraphPrFragmentSchema = z.object({
  algn: z.enum(["l", "ctr", "r", "just"]).optional(),
  rtl: ooxmlBoolSchema.optional(),
  lvl: ooxmlNonNegativeIntSchema.refine((value) => value <= 8).optional(),
});

const SpacingPointsXmlSchema = z.object({
  val: ooxmlNonNegativeIntSchema,
});

const SpacingPercentXmlSchema = z.object({
  val: ooxmlPositiveIntSchema,
});

const BulletCharXmlSchema = z.object({
  char: z.string().min(1),
});

const AutoNumXmlSchema = z.object({
  type: z.string().min(1),
  startAt: ooxmlPositiveIntSchema.optional(),
});

const RunPrFragmentSchema = z.object({
  lang: z.string().min(1).optional(),
  sz: ooxmlPositiveIntSchema.optional(),
  b: ooxmlBoolSchema.optional(),
  i: ooxmlBoolSchema.optional(),
  u: z.enum(["none", "sng"]).optional(),
  strike: z.enum(["noStrike", "sngStrike"]).optional(),
  cap: z.enum(["none", "all"]).optional(),
  spc: ooxmlIntSchema.optional(),
  baseline: ooxmlIntSchema.optional(),
  dirty: ooxmlBoolSchema.optional(),
  smtClean: ooxmlBoolSchema.optional(),
});

const TextFieldXmlSchema = z.object({
  id: z.string().min(1),
  type: z.string().min(1),
});

const OuterShadowXmlSchema = z.object({
  blurRad: ooxmlNonNegativeIntSchema.optional(),
  dist: ooxmlNonNegativeIntSchema.optional(),
  dir: ooxmlIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: ooxmlBoolSchema.optional(),
});

const PercentModifierSchema = z.object({
  val: ooxmlNonNegativeIntSchema,
});

const SrgbColorSchema = z.object({
  val: hexRgb6Schema,
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

// Metadata placeholders (`dt`, `ftr`, `sldNum`) are owned by `footerPlaceholderShapeImportHandler`.
const textCompatiblePlaceholderRoles = new Set<string>([
  "title",
  "ctrTitle",
  "subTitle",
  "body",
  "hdr",
  "obj",
]);

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

export { isPlaceholderSpElement };

/** Side-effect-free predicate matching populated text-compatible `<p:sp>` placeholder elements. */
export function isTextPlaceholderSpElement(cursor: XmlElementCursor): boolean {
  const ph = directPlaceholderElement(cursor, "nvSpPr");
  if (ph === undefined) return false;
  if (directElementChild(cursor.element, P_NS, "txBody") === undefined) {
    return false;
  }
  const rawType = ph.getAttribute("type");
  const role = rawType === null || rawType === "" ? "obj" : rawType;
  return textCompatiblePlaceholderRoles.has(role);
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

  let baseString: HexColor | undefined;
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
  allowShadowNone = false,
): Shadow | "none" | undefined {
  const effectLst = rPrCursor.child(A_NS, "effectLst");
  if (effectLst === undefined) return undefined;
  const outerShdw = effectLst.child(A_NS, "outerShdw");
  if (outerShdw === undefined) {
    return allowShadowNone && effectLst.element.children.length === 0
      ? "none"
      : undefined;
  }

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

export type ParsedRunStyle = NonNullable<SpanDeclarationProps["style"]>;

function readRunProperties(
  rPrCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
  allowShadowNone = false,
): {
  readonly style: ParsedRunStyle;
  readonly sizePt?: number;
  readonly inlineKind?: "sup" | "sub";
} {
  const raw = rPrCursor.attrsObject([
    "lang",
    "sz",
    "b",
    "i",
    "u",
    "strike",
    "cap",
    "spc",
    "baseline",
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
  let baselineShift: ParsedRunStyle["baselineShift"] | undefined;
  let inlineKind: "sup" | "sub" | undefined;

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
    if (data.baseline !== undefined) {
      if (data.baseline === 30_000) {
        inlineKind = "sup";
      } else if (data.baseline === -25_000) {
        inlineKind = "sub";
      } else if (data.baseline !== 0) {
        if (data.baseline < -2_147_483_648 || data.baseline > 2_147_483_647) {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "omitted",
            location: {
              part,
              xpath: `${rPrCursor.xpath}/@baseline`,
              slideIndex,
            },
            attributeOrElement: "@baseline",
            rawValue: String(data.baseline),
          });
        } else {
          baselineShift = `${data.baseline / 1000}%`;
        }
      }
    }
  }

  let color: NonNullable<TextStyle["color"]> | undefined;
  const solidFill = rPrCursor.child(A_NS, "solidFill");
  if (solidFill !== undefined) {
    color = readColorElement(solidFill, part, slideIndex, diagnostics);
  }

  const textShadow = readOuterShadow(
    rPrCursor,
    part,
    slideIndex,
    diagnostics,
    allowShadowNone,
  );

  const { fontFamily, eastAsianFontFamily } = readRunFontSlots(rPrCursor);

  return {
    style: {
      ...(fontFamily === undefined ? {} : { fontFamily }),
      ...(eastAsianFontFamily === undefined ? {} : { eastAsianFontFamily }),
      ...(fontSize === undefined ? {} : { fontSize }),
      ...(fontWeight === undefined ? {} : { fontWeight }),
      ...(fontStyle === undefined ? {} : { fontStyle }),
      ...(color === undefined ? {} : { color }),
      ...(letterSpacing === undefined ? {} : { letterSpacing }),
      ...(textTransform === undefined ? {} : { textTransform }),
      ...(textDecoration === undefined ? {} : { textDecoration }),
      ...(textShadow === undefined ? {} : { textShadow }),
      ...(lang === undefined ? {} : { lang }),
      ...(baselineShift === undefined ? {} : { baselineShift }),
    },
    ...(sizePt === undefined ? {} : { sizePt }),
    ...(inlineKind === undefined ? {} : { inlineKind }),
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
    cursor.ignoreEmptyChild(A_NS, "pPr");
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
      readonly kind: "text" | "sup" | "sub";
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
  allowShadowNone = false,
): ImportedParagraph {
  const pPr = pCursor.child(A_NS, "pPr");

  let align: TextAlignment | undefined;
  let lvl: number | undefined;
  let list: TextList | undefined;
  let lnSpcNode: XmlElementCursor | undefined;
  let spcAftNode: XmlElementCursor | undefined;
  let defParaSizePt: number | undefined;

  if (pPr !== undefined) {
    const rawPPr = pPr.attrsObject([
      "algn",
      "rtl",
      "lvl",
      "marL",
      "indent",
      "fontAlgn",
      "eaLnBrk",
      "latinLnBrk",
      "hangingPunct",
      "defTabSz",
    ]);
    const parsedPPr = ParagraphPrFragmentSchema.safeParse({
      algn: rawPPr["algn"],
      rtl: rawPPr["rtl"],
      lvl: rawPPr["lvl"],
    });
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

    const spcBefNode = pPr.child(A_NS, "spcBef");
    if (spcBefNode !== undefined) {
      spcBefNode.child(A_NS, "spcPts")?.consumeAttr("val");
      spcBefNode.child(A_NS, "spcPct")?.consumeAttr("val");
    }
    lnSpcNode = pPr.child(A_NS, "lnSpc");
    spcAftNode = pPr.child(A_NS, "spcAft");

    pPr.ignoreEmptyChild(A_NS, "buNone");
    const buFont = pPr.child(A_NS, "buFont");
    buFont?.attrsObject(["typeface", "pitchFamily", "charset", "panose"]);
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

    const defRPr = pPr.child(A_NS, "defRPr");
    if (defRPr !== undefined) {
      const parsedDef = readRunProperties(
        defRPr,
        part,
        slideIndex,
        diagnostics,
        allowShadowNone,
      );
      defParaSizePt = parsedDef.sizePt;
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
          : readRunProperties(
              rPr,
              part,
              slideIndex,
              diagnostics,
              allowShadowNone,
            );
      const t = child.child(A_NS, "t");
      if (t !== undefined) {
        t.consumeAttr("space", XML_NS);
        runs.push({
          kind: parsedRPr.inlineKind ?? "text",
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
          : readRunProperties(
              rPr,
              part,
              slideIndex,
              diagnostics,
              allowShadowNone,
            );
      runs.push({
        kind: parsedRPr.inlineKind ?? "text",
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
                allowShadowNone,
              );
        const fieldStyle: ParsedRunStyle =
          parsedRPr.inlineKind === undefined
            ? parsedRPr.style
            : {
                ...parsedRPr.style,
                baselineShift: parsedRPr.inlineKind === "sup" ? "30%" : "-25%",
              };
        runs.push({
          kind: "slide-number",
          style: fieldStyle,
          ...(parsedRPr.sizePt === undefined
            ? {}
            : { sizePt: parsedRPr.sizePt }),
        });
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "endParaRPr") {
      const parsedEnd = readRunProperties(
        child,
        part,
        slideIndex,
        diagnostics,
        allowShadowNone,
      );
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
    runs.find((run) => run.sizePt !== undefined)?.sizePt ??
    endParaSizePt ??
    defParaSizePt ??
    18;

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
      if (!parsed.success || parsed.data.val > 13_200_000) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part, xpath: spcPct.xpath, slideIndex },
          attributeOrElement: "a:spcPct/@val",
          rawValue: raw["val"] ?? "",
        });
      } else {
        lineHeight = `${parsed.data.val / 1000}%`;
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
  _clrMap?: Readonly<Record<string, NativeColorSlot>>,
  options?: { readonly allowShadowNone?: boolean },
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
      "anchorCtr",
      "spcFirstLastPara",
      "numCol",
      "spcCol",
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
    const normAutofit = bodyPr.child(A_NS, "normAutofit");
    normAutofit?.attrsObject(["fontScale", "lnSpcReduction"]);
  }

  txBody.ignoreEmptyChild(A_NS, "lstStyle");

  const importedParagraphs = txBody
    .children(A_NS, "p")
    .map((pCursor) =>
      readParagraph(
        pCursor,
        part,
        slideIndex,
        diagnostics,
        options?.allowShadowNone === true,
      ),
    );

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
          const runProps =
            Object.keys(run.style).length === 0 ? {} : { style: run.style };
          if (run.kind === "sup") {
            return createSupNode(runProps).content(run.text);
          }
          if (run.kind === "sub") {
            return createSubNode(runProps).content(run.text);
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

function textBoxVisualStyle(
  appearanceStyle: VisualStyle,
  borderArrows: LineEndStyle | undefined,
  geometry: Geometry | undefined,
  boxSize: { readonly width: number; readonly height: number } | undefined,
  rotate: number | undefined,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): NonNullable<TextDeclarationProps["style"]> {
  const { border: rawBorder, ...restVisualStyle } = appearanceStyle;
  const border: NonNullable<TextDeclarationProps["style"]>["border"] =
    rawBorder;
  const hasHeadArrow = borderArrows?.startArrow !== undefined;
  const hasTailArrow = borderArrows?.endArrow !== undefined;
  if (hasHeadArrow || hasTailArrow) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location,
      namespaceUri: A_NS,
      localName: hasHeadArrow ? "headEnd" : "tailEnd",
      mechanismHint: "textbox-line-arrow",
    });
  }

  let borderRadius: number | undefined;
  if (geometry === "roundRect") {
    if (boxSize !== undefined) {
      borderRadius =
        (16667 / 100_000) * Math.min(boxSize.width, boxSize.height);
    }
  } else if (typeof geometry === "object" && geometry.type === "preset") {
    if (geometry.name === "roundRect") {
      if (boxSize !== undefined) {
        const adj = geometry.adjust?.["adj"] ?? 16667;
        if (adj > 0) {
          borderRadius =
            (adj / 100_000) * Math.min(boxSize.width, boxSize.height);
        }
      }
    } else if (geometry.name !== "rect") {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location,
        namespaceUri: A_NS,
        localName: "prstGeom",
        mechanismHint: `textbox-preset:${geometry.name}`,
      });
    }
  } else if (typeof geometry === "object" && geometry.type === "custom") {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location,
      namespaceUri: A_NS,
      localName: "custGeom",
      mechanismHint: "textbox-custGeom",
    });
  }

  return {
    ...restVisualStyle,
    ...(border === undefined ? {} : { border }),
    ...(borderRadius === undefined ? {} : { borderRadius }),
    ...(rotate === undefined ? {} : { rotate }),
  };
}

function flattenImportedTextBody(
  imported: ImportedRichTextBody | undefined,
  visualStyle: NonNullable<TextDeclarationProps["style"]>,
): {
  readonly boxStyle: NonNullable<TextDeclarationProps["style"]>;
  readonly singleText?: string;
  readonly paragraphsToPopulate?: readonly ImportedParagraph[];
} {
  const importedParagraphs = imported?.importedParagraphs ?? [];
  const singleParagraph =
    importedParagraphs.length === 1 ? importedParagraphs[0] : undefined;
  const singleTextRun =
    singleParagraph !== undefined &&
    singleParagraph.style.list === undefined &&
    singleParagraph.runs.length === 1 &&
    singleParagraph.runs[0]!.kind === "text" &&
    singleParagraph.runs[0]!.style.baselineShift === undefined
      ? singleParagraph.runs[0]
      : undefined;
  const singleEmptyParagraph =
    singleParagraph !== undefined &&
    singleParagraph.style.list === undefined &&
    singleParagraph.runs.length === 0;

  const flatParagraphStyle: Omit<ParagraphStyle, "list"> | undefined =
    (singleTextRun === undefined && !singleEmptyParagraph) ||
    singleParagraph === undefined
      ? undefined
      : (() => {
          const { list: _list, ...rest } = singleParagraph.style;
          return rest;
        })();

  const { baselineShift: _baselineShift, ...hoistableRunStyle } =
    singleTextRun === undefined ? {} : singleTextRun.style;

  const boxStyle: NonNullable<TextDeclarationProps["style"]> = {
    ...imported?.boxStyle,
    ...visualStyle,
    ...flatParagraphStyle,
    ...hoistableRunStyle,
  };

  if (singleTextRun !== undefined) {
    return { boxStyle, singleText: singleTextRun.text };
  }
  if (!singleEmptyParagraph && importedParagraphs.length > 0) {
    return { boxStyle, paragraphsToPopulate: importedParagraphs };
  }
  return { boxStyle };
}

/**
 * Reads a Transitional `<p:sp>` text box (`<p:cNvSpPr txBox="1"/>` without `<p:ph>`) into an editable `TextNode`.
 */
export function readTextBoxSp(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, package: pkg } = context;
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

  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  const { objectId, objectName, link, deferredSlideLink } =
    readNonVisualCommonProperties(cNvPr, context, diagnostics);
  cNvPr?.ignoreEmptyChild(A_NS, "extLst");

  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    const rawFlags = cNvSpPr.attrsObject(["txBox"]);
    NonVisualSpPrFlagsSchema.safeParse(rawFlags);
    const spLocks = cNvSpPr.child(A_NS, "spLocks");
    spLocks?.attrsObject([
      "noGrp",
      "noRot",
      "noChangeAspect",
      "noMove",
      "noResize",
      "noSelect",
    ]);
  }
  consumeNonPlaceholderNvPr(nvSpPr);

  const location: ImportLocation = {
    part,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };

  const spPr = cursor.child(P_NS, "spPr");
  const txBody = cursor.child(P_NS, "txBody");
  if (spPr === undefined || txBody === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: cursor.xpath },
    });
    return { diagnostics };
  }

  const visuals = readSpPrVisuals(spPr, pkg, location, diagnostics);
  if (visuals === undefined) {
    return { diagnostics };
  }

  const rawVisualStyle = textBoxVisualStyle(
    visuals.style,
    visuals.borderArrows,
    visuals.geometry,
    { width: visuals.transform.width, height: visuals.transform.height },
    visuals.transform.rotate,
    location,
    diagnostics,
  );
  // Free text boxes are transparent and borderless by default; strip default <a:noFill/> / <a:ln><a:noFill/></a:ln>.
  const {
    background: rawBg,
    border: rawBorder,
    ...restFreeVisuals
  } = rawVisualStyle;
  const visualStyle: NonNullable<TextDeclarationProps["style"]> = {
    ...restFreeVisuals,
    ...(rawBg === undefined || rawBg === "none" ? {} : { background: rawBg }),
    ...(rawBorder === undefined || rawBorder === "none"
      ? {}
      : { border: rawBorder }),
  };

  const imported = readRichTextBody(txBody, part, slideIndex, diagnostics);
  const { boxStyle, singleText, paragraphsToPopulate } =
    flattenImportedTextBody(imported, visualStyle);

  const props: TextDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    style: {
      position: "absolute",
      left: visuals.transform.x,
      top: visuals.transform.y,
      width: visuals.transform.width,
      height: visuals.transform.height,
      ...boxStyle,
    },
  };

  const node = createTextNode(props);
  if (singleText !== undefined) {
    node.content(singleText);
  } else if (paragraphsToPopulate !== undefined) {
    populateRichTextNode(node, paragraphsToPopulate);
  }

  attachDeferredSlideLink(node, deferredSlideLink, context);
  return { node, diagnostics };
}

/**
 * Reads a Transitional `<p:sp>` bound placeholder (`<p:nvPr><p:ph .../></p:nvPr>`)
 * into a layout-bound `TextNode` preserving differential transforms and styles.
 */
export function readTextPlaceholderSp(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const {
    cursor,
    part,
    slideIndex,
    layoutPlaceholders,
    package: pkg,
  } = context;
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

  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  const { objectId, objectName, link, deferredSlideLink } =
    readNonVisualCommonProperties(cNvPr, context, diagnostics);
  cNvPr?.ignoreEmptyChild(A_NS, "extLst");

  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    cNvSpPr.consumeAttr("txBox");
    const spLocks = cNvSpPr.child(A_NS, "spLocks");
    spLocks?.attrsObject([
      "noGrp",
      "noRot",
      "noChangeAspect",
      "noMove",
      "noResize",
      "noSelect",
    ]);
  }

  const location: ImportLocation = {
    part,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };

  const nvPr = nvSpPr.child(P_NS, "nvPr");
  const phResult = readNvPrPlaceholder(
    nvPr,
    { ...location, xpath: cursor.xpath },
    diagnostics,
  );
  if (phResult.kind === "none") {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: cursor.xpath },
    });
    return { diagnostics };
  }
  if (phResult.kind === "error") {
    return { diagnostics };
  }
  const nativeRef: NativePlaceholderRef = phResult.nativeRef;

  const spPr = cursor.child(P_NS, "spPr");
  if (spPr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: cursor.xpath },
    });
    return { diagnostics };
  }

  let localBox:
    | {
        readonly x: number;
        readonly y: number;
        readonly width: number;
        readonly height: number;
      }
    | undefined;
  let rotate: number | undefined;
  let flipH: boolean | undefined;
  let flipV: boolean | undefined;
  let fallbackUnits = { w: 0, h: 0 };
  const xfrm = spPr.child(A_NS, "xfrm");
  if (xfrm !== undefined) {
    const hasExplicitFlipH = xfrm.element.hasAttribute("flipH");
    const hasExplicitFlipV = xfrm.element.hasAttribute("flipV");
    const rawXfrm = xfrm.attrsObject(["rot", "flipH", "flipV"]);
    const parsedXfrm = transformAttributesFragmentSchema.safeParse(rawXfrm);
    if (parsedXfrm.success) {
      if (parsedXfrm.data.rot !== undefined && parsedXfrm.data.rot !== 0) {
        rotate = parsedXfrm.data.rot / 60_000;
      }
      if (hasExplicitFlipH && parsedXfrm.data.flipH !== undefined) {
        flipH = parsedXfrm.data.flipH;
      }
      if (hasExplicitFlipV && parsedXfrm.data.flipV !== undefined) {
        flipV = parsedXfrm.data.flipV;
      }
    }
    const off = xfrm.child(A_NS, "off");
    const ext = xfrm.child(A_NS, "ext");
    if (off !== undefined && ext !== undefined) {
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
        return { diagnostics };
      }
      fallbackUnits = { w: parsedExt.data.cx, h: parsedExt.data.cy };
      localBox = {
        x: emuToPx(parsedOff.data.x),
        y: emuToPx(parsedOff.data.y),
        width: emuToPx(parsedExt.data.cx),
        height: emuToPx(parsedExt.data.cy),
      };
    }
  }

  const matchedKey = matchLayoutPlaceholderKey(layoutPlaceholders, {
    type: nativeRef.type,
    idx: nativeRef.idx,
  });
  const layoutPlaceholder =
    matchedKey === undefined ? undefined : layoutPlaceholders?.[matchedKey];
  const masterPlaceholder = matchMasterPlaceholderSpec(
    context.masterPlaceholders,
    nativeRef.type,
  );
  const layoutPhWidth =
    typeof layoutPlaceholder?.style?.width === "number"
      ? layoutPlaceholder.style.width
      : undefined;
  const layoutPhHeight =
    typeof layoutPlaceholder?.style?.height === "number"
      ? layoutPlaceholder.style.height
      : undefined;
  const masterPhWidth =
    typeof masterPlaceholder?.style?.width === "number"
      ? masterPlaceholder.style.width
      : undefined;
  const masterPhHeight =
    typeof masterPlaceholder?.style?.height === "number"
      ? masterPlaceholder.style.height
      : undefined;
  const inheritedWidth = layoutPhWidth ?? masterPhWidth;
  const inheritedHeight = layoutPhHeight ?? masterPhHeight;
  const inheritedPlaceholderBox =
    inheritedWidth !== undefined && inheritedHeight !== undefined
      ? { width: inheritedWidth, height: inheritedHeight }
      : undefined;
  const effectiveFallbackUnits =
    localBox === undefined && inheritedPlaceholderBox !== undefined
      ? {
          w: Math.round(inheritedPlaceholderBox.width * EMU_PER_PX),
          h: Math.round(inheritedPlaceholderBox.height * EMU_PER_PX),
        }
      : fallbackUnits;

  const appearance = readSpPrAppearance(
    spPr,
    pkg,
    effectiveFallbackUnits,
    location,
    diagnostics,
    { preservePlaceholderClears: true },
  );
  const visualStyle = textBoxVisualStyle(
    appearance.style,
    appearance.borderArrows,
    appearance.geometry,
    localBox ?? inheritedPlaceholderBox,
    rotate,
    location,
    diagnostics,
  );

  const txBody = cursor.child(P_NS, "txBody");
  const imported =
    txBody === undefined
      ? undefined
      : readRichTextBody(txBody, part, slideIndex, diagnostics, undefined, {
          allowShadowNone: true,
        });

  const { boxStyle, singleText, paragraphsToPopulate } =
    flattenImportedTextBody(imported, visualStyle);

  const combinedStyle: NonNullable<TextDeclarationProps["style"]> = {
    ...(localBox === undefined
      ? {}
      : {
          position: "absolute",
          left: localBox.x,
          top: localBox.y,
          width: localBox.width,
          height: localBox.height,
        }),
    ...(flipH === undefined ? {} : { flipH }),
    ...(flipV === undefined ? {} : { flipV }),
    ...boxStyle,
  };

  let validatedStyle: TextDeclarationProps["style"] | undefined;
  if (Object.keys(combinedStyle).length > 0) {
    const parsedStyle = placeholderStyleSchema.safeParse(combinedStyle);
    if (parsedStyle.success) {
      validatedStyle = parsedStyle.data;
    } else {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { ...location, xpath: cursor.xpath },
        attributeOrElement: "p:sp",
        rawValue: parsedStyle.error.issues[0]?.message ?? "invalid style",
      });
    }
  }

  const props: TextDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    ...(matchedKey === undefined ? {} : { placeholder: matchedKey }),
    nativePlaceholder: nativeRef,
    ...(validatedStyle === undefined ? {} : { style: validatedStyle }),
  };

  const node = createTextNode(props);
  if (singleText !== undefined) {
    node.content(singleText);
  } else if (paragraphsToPopulate !== undefined) {
    populateRichTextNode(node, paragraphsToPopulate);
  }

  attachDeferredSlideLink(node, deferredSlideLink, context);
  return { node, diagnostics };
}

export const textBoxImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: (cursor) =>
    isTextBoxSpElement(cursor) && !isPlaceholderSpElement(cursor),
  read: readTextBoxSp,
};

export const textPlaceholderImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: isTextPlaceholderSpElement,
  read: readTextPlaceholderSp,
};
