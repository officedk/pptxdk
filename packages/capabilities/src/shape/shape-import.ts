import {
  assertValidLang,
  type LinkTarget,
  type TextAlignment,
  type TextStyle,
} from "@pptxdk/core/capability";
import {
  A_NS,
  ooxmlBoolSchema,
  ooxmlIntSchema,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
  P_NS,
  readRunFontSlots,
  XML_NS,
  emuToPx,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import { isLineNormalizedSpElement } from "../line/line-import.ts";
import {
  attachDeferredSlideLink,
  consumeNonPlaceholderNvPr,
  readClickLink,
  type ClickLinkReadResult,
} from "../shared/identity.ts";
import { isPlaceholderSpElement } from "../shared/placeholder-import.ts";
import type { ParagraphStyle, TextList } from "../shared/text-service.ts";
import {
  readDrawingColor,
  readEffectList,
  readSpPrVisuals,
} from "../shared/visual.ts";
import {
  createParagraphNode,
  createSlideNumberNode,
  createSpanNode,
  type ParagraphDeclarationProps,
  type SpanDeclarationProps,
} from "../text/text-declaration.ts";
import {
  createShapeNode,
  type ShapeDeclarationProps,
} from "./shape-declaration.ts";

const NonVisualIdAttrSchema = z.object({
  id: ooxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const BodyPrAttrSchema = z.object({
  wrap: z.enum(["square", "none"]).optional(),
  rtlCol: ooxmlBoolSchema.optional(),
  lIns: ooxmlNonNegativeIntSchema.optional(),
  tIns: ooxmlNonNegativeIntSchema.optional(),
  rIns: ooxmlNonNegativeIntSchema.optional(),
  bIns: ooxmlNonNegativeIntSchema.optional(),
  anchor: z.enum(["t", "ctr", "b"]).optional(),
});

const ParagraphPrAttrSchema = z.object({
  algn: z.enum(["l", "ctr", "r", "just"]).optional(),
  rtl: ooxmlBoolSchema.optional(),
  lvl: ooxmlNonNegativeIntSchema.optional(),
});

const RunPrAttrSchema = z.object({
  lang: z.string().min(1).optional(),
  sz: ooxmlPositiveIntSchema.optional(),
  b: ooxmlBoolSchema.optional(),
  i: ooxmlBoolSchema.optional(),
  u: z.enum(["none", "sng"]).optional(),
  strike: z.enum(["noStrike", "sngStrike"]).optional(),
  cap: z.enum(["none", "all"]).optional(),
  spc: ooxmlIntSchema.optional(),
  dirty: ooxmlBoolSchema.optional(),
  smtClean: ooxmlBoolSchema.optional(),
});

const BulletCharAttrSchema = z.object({
  char: z.string().min(1),
});

const BulletAutoNumAttrSchema = z.object({
  type: z.string().min(1),
  startAt: ooxmlIntSchema.optional(),
});

const SpacingValAttrSchema = z.object({
  val: ooxmlNonNegativeIntSchema,
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

function isTextBoxSpElement(cursor: XmlElementCursor): boolean {
  const nvSpPr = directElementChild(cursor.element, P_NS, "nvSpPr");
  if (nvSpPr === undefined) return false;
  const cNvSpPr = directElementChild(nvSpPr, P_NS, "cNvSpPr");
  if (cNvSpPr === undefined) return false;
  const txBox = cNvSpPr.getAttribute("txBox");
  return txBox === "1" || txBox === "true";
}

function isCustomGeomSpElement(cursor: XmlElementCursor): boolean {
  const spPr = directElementChild(cursor.element, P_NS, "spPr");
  if (spPr === undefined) return false;
  return directElementChild(spPr, A_NS, "custGeom") !== undefined;
}

/** Side-effect-free predicate matching non-textbox preset shapes (`<p:sp>` without `txBox="1"`, `<p:ph>`, `<a:custGeom>`, or normalized 1D lines). */
export function isPresetShapeSpElement(cursor: XmlElementCursor): boolean {
  return (
    !isPlaceholderSpElement(cursor) &&
    !isTextBoxSpElement(cursor) &&
    !isCustomGeomSpElement(cursor) &&
    !isLineNormalizedSpElement(cursor)
  );
}

/** Side-effect-free predicate matching non-textbox custom-geometry shapes (`<p:sp>` without `<p:ph>` and with `<a:custGeom>`). */
export function isCustomGeometryShapeSpElement(
  cursor: XmlElementCursor,
): boolean {
  return (
    !isPlaceholderSpElement(cursor) &&
    !isTextBoxSpElement(cursor) &&
    isCustomGeomSpElement(cursor)
  );
}

/** Reads `<p:cNvPr>` id, name, and hyperlink without assuming shape vs group wrapper tag. */
export function readNonVisualCommonProperties(
  cNvPrCursor: XmlElementCursor | undefined,
  context: ShapeTreeImportContext,
  diagnostics: ImportDiagnostic[],
): {
  readonly objectId?: number;
  readonly objectName?: string;
  readonly link?: LinkTarget;
  readonly deferredSlideLink?: ClickLinkReadResult["deferredSlideLink"];
} {
  if (cNvPrCursor === undefined) return {};
  const raw = cNvPrCursor.attrsObject(["id", "name"]);
  const parsed = NonVisualIdAttrSchema.safeParse(raw);
  let objectId: number | undefined;
  let objectName: string | undefined;
  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: {
        part: context.part,
        xpath: cNvPrCursor.xpath,
        slideIndex: context.slideIndex,
      },
      attributeOrElement: "p:cNvPr",
      rawValue: JSON.stringify(raw),
    });
  } else {
    objectId = parsed.data.id;
    if (parsed.data.name !== undefined && parsed.data.name.trim() !== "") {
      objectName = parsed.data.name;
    }
  }
  const hlinkClick = cNvPrCursor.child(A_NS, "hlinkClick");
  let clickResult: ClickLinkReadResult | undefined;
  if (hlinkClick !== undefined) {
    clickResult = readClickLink(hlinkClick, context, {
      ...(objectId === undefined ? {} : { objectId }),
      ...(objectName === undefined ? {} : { objectName }),
    });
    diagnostics.push(...clickResult.diagnostics);
  }
  return {
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
    ...(clickResult?.link === undefined ? {} : { link: clickResult.link }),
    ...(clickResult?.deferredSlideLink === undefined
      ? {}
      : { deferredSlideLink: clickResult.deferredSlideLink }),
  };
}

type ParsedRunStyle = NonNullable<SpanDeclarationProps["style"]>;

function readShapeRunProperties(
  rPrCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ParsedRunStyle {
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
  const parsed = RunPrAttrSchema.safeParse(raw);
  let lang: string | undefined;
  let fontSize: NonNullable<TextStyle["fontSize"]> | undefined;
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
      location: { ...location, xpath: rPrCursor.xpath },
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
          location: { ...location, xpath: `${rPrCursor.xpath}/@lang` },
          attributeOrElement: "@lang",
          rawValue: data.lang,
        });
      }
    }
    if (data.sz !== undefined) {
      fontSize = `${data.sz / 100}pt`;
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

  const solidFill = rPrCursor.child(A_NS, "solidFill");
  const color =
    solidFill === undefined
      ? undefined
      : readDrawingColor(solidFill, location, diagnostics);

  const effectLst = rPrCursor.child(A_NS, "effectLst");
  const effects =
    effectLst === undefined
      ? {}
      : readEffectList(effectLst, location, diagnostics);

  const { fontFamily, eastAsianFontFamily } = readRunFontSlots(rPrCursor);

  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(eastAsianFontFamily === undefined ? {} : { eastAsianFontFamily }),
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(fontWeight === undefined ? {} : { fontWeight }),
    ...(fontStyle === undefined ? {} : { fontStyle }),
    ...(color === undefined ? {} : { color }),
    ...(letterSpacing === undefined ? {} : { letterSpacing }),
    ...(textTransform === undefined ? {} : { textTransform }),
    ...(textDecoration === undefined ? {} : { textDecoration }),
    ...(effects.boxShadow === undefined
      ? {}
      : { textShadow: effects.boxShadow }),
    ...(lang === undefined ? {} : { lang }),
  };
}

type ImportedShapeRun =
  | {
      readonly kind: "text";
      readonly text: string;
      readonly style: ParsedRunStyle;
    }
  | { readonly kind: "slidenum"; readonly style: ParsedRunStyle };

type ImportedShapeParagraph = {
  readonly runs: readonly ImportedShapeRun[];
  readonly style: ParagraphStyle;
  readonly defRunStyle?: ParsedRunStyle;
};

function readLineSpacing(
  lnSpcCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): number | undefined {
  const spcPts = lnSpcCursor.child(A_NS, "spcPts");
  if (spcPts !== undefined) {
    const parsed = SpacingValAttrSchema.safeParse(spcPts.attrsObject(["val"]));
    if (parsed.success && parsed.data.val > 0) {
      return parsed.data.val / 75;
    }
  }
  const spcPct = lnSpcCursor.child(A_NS, "spcPct");
  if (spcPct !== undefined) {
    spcPct.consumeAttr("val");
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: spcPct.xpath },
      namespaceUri: A_NS,
      localName: "spcPct",
    });
  }
  return undefined;
}

function readShapeParagraph(
  pCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ImportedShapeParagraph {
  let align: TextAlignment | undefined;
  let lineHeight: number | undefined;
  let spacingAfter: number | undefined;
  let list: TextList | undefined;
  let defRunStyle: ParsedRunStyle | undefined;
  let level = 0;

  const pPr = pCursor.child(A_NS, "pPr");
  if (pPr !== undefined) {
    const rawPPr = pPr.attrsObject(["algn", "rtl", "lvl"]);
    const parsedPPr = ParagraphPrAttrSchema.safeParse(rawPPr);
    if (!parsedPPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { ...location, xpath: pPr.xpath },
        attributeOrElement: "a:pPr",
        rawValue: JSON.stringify(rawPPr),
      });
    } else {
      if (parsedPPr.data.algn !== undefined) {
        align = algnToTextAlign[parsedPPr.data.algn];
      }
      if (parsedPPr.data.lvl !== undefined) {
        level = parsedPPr.data.lvl;
      }
    }

    const lnSpc = pPr.child(A_NS, "lnSpc");
    if (lnSpc !== undefined) {
      lineHeight = readLineSpacing(lnSpc, location, diagnostics);
    }

    const spcAft = pPr.child(A_NS, "spcAft");
    if (spcAft !== undefined) {
      const spcPts = spcAft.child(A_NS, "spcPts");
      if (spcPts !== undefined) {
        const parsed = SpacingValAttrSchema.safeParse(
          spcPts.attrsObject(["val"]),
        );
        if (parsed.success) {
          spacingAfter = parsed.data.val / 75;
        }
      }
    }

    pPr.ignoreEmptyChild(A_NS, "buNone");
    const buChar = pPr.child(A_NS, "buChar");
    if (buChar !== undefined) {
      const parsedBu = BulletCharAttrSchema.safeParse(
        buChar.attrsObject(["char"]),
      );
      if (parsedBu.success) {
        list = {
          type: "bullet",
          character: parsedBu.data.char,
          ...(level > 0 ? { level } : {}),
        };
      }
    } else {
      const buAutoNum = pPr.child(A_NS, "buAutoNum");
      if (buAutoNum !== undefined) {
        const parsedAuto = BulletAutoNumAttrSchema.safeParse(
          buAutoNum.attrsObject(["type", "startAt"]),
        );
        if (parsedAuto.success) {
          list = {
            type: "number",
            style: "arabicPeriod",
            ...(parsedAuto.data.startAt === undefined
              ? {}
              : { startAt: parsedAuto.data.startAt }),
            ...(level > 0 ? { level } : {}),
          };
        }
      }
    }

    const defRPr = pPr.child(A_NS, "defRPr");
    if (defRPr !== undefined) {
      defRunStyle = readShapeRunProperties(defRPr, location, diagnostics);
    }
  }

  const runs: ImportedShapeRun[] = [];
  for (const child of pCursor.remainingElementChildren()) {
    if (child.namespaceUri === A_NS && child.localName === "r") {
      const rPr = child.child(A_NS, "rPr");
      const runStyle =
        rPr === undefined
          ? {}
          : readShapeRunProperties(rPr, location, diagnostics);
      const t = child.child(A_NS, "t");
      if (t !== undefined) {
        t.consumeAttr("space", XML_NS);
        runs.push({ kind: "text", text: t.textContent, style: runStyle });
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "fld") {
      child.consumeAttr("id");
      const fieldType = child.attr("type");
      const rPr = child.child(A_NS, "rPr");
      const runStyle =
        rPr === undefined
          ? {}
          : readShapeRunProperties(rPr, location, diagnostics);
      const t = child.child(A_NS, "t");
      if (t !== undefined) {
        t.consumeAttr("space", XML_NS);
      }
      if (fieldType === "slidenum") {
        runs.push({ kind: "slidenum", style: runStyle });
      } else {
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          namespaceUri: A_NS,
          localName: "fld",
          mechanismHint: `field:${fieldType ?? "unknown"}`,
        });
        if (t !== undefined) {
          runs.push({ kind: "text", text: t.textContent, style: runStyle });
        }
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "br") {
      const rPr = child.child(A_NS, "rPr");
      const runStyle =
        rPr === undefined
          ? {}
          : readShapeRunProperties(rPr, location, diagnostics);
      runs.push({ kind: "text", text: "\n", style: runStyle });
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "endParaRPr") {
      readShapeRunProperties(child, location, diagnostics);
      continue;
    }
    diagnostics.push({
      code: "unconsumed-xml",
      category: "unsupported",
      disposition: "omitted",
      kind: "element",
      location: { ...location, xpath: child.xpath },
      namespaceUri: child.namespaceUri,
      localName: child.localName,
    });
  }

  const pStyle: ParagraphStyle = {
    ...(align === undefined ? {} : { align }),
    ...(lineHeight === undefined ? {} : { lineHeight }),
    ...(spacingAfter === undefined ? {} : { spacingAfter }),
    ...(list === undefined ? {} : { list }),
  };

  return {
    runs,
    style: pStyle,
    ...(defRunStyle === undefined ? {} : { defRunStyle }),
  };
}

function readShapeTxBody(
  txBodyCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): {
  readonly paddingTop?: number;
  readonly paddingRight?: number;
  readonly paddingBottom?: number;
  readonly paddingLeft?: number;
  readonly verticalAlign?: NonNullable<TextStyle["verticalAlign"]>;
  readonly textWrap?: NonNullable<TextStyle["textWrap"]>;
  readonly paragraphs: readonly ImportedShapeParagraph[];
} {
  let paddingTop: number | undefined;
  let paddingRight: number | undefined;
  let paddingBottom: number | undefined;
  let paddingLeft: number | undefined;
  let verticalAlign: NonNullable<TextStyle["verticalAlign"]> | undefined;
  let textWrap: NonNullable<TextStyle["textWrap"]> | undefined;

  const bodyPr = txBodyCursor.child(A_NS, "bodyPr");
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
    const parsedBodyPr = BodyPrAttrSchema.safeParse(rawBodyPr);
    if (!parsedBodyPr.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { ...location, xpath: bodyPr.xpath },
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

  txBodyCursor.ignoreEmptyChild(A_NS, "lstStyle");

  const paragraphs = txBodyCursor
    .children(A_NS, "p")
    .map((pCursor) => readShapeParagraph(pCursor, location, diagnostics));

  return {
    ...(paddingTop === undefined ? {} : { paddingTop }),
    ...(paddingRight === undefined ? {} : { paddingRight }),
    ...(paddingBottom === undefined ? {} : { paddingBottom }),
    ...(paddingLeft === undefined ? {} : { paddingLeft }),
    ...(verticalAlign === undefined ? {} : { verticalAlign }),
    ...(textWrap === undefined ? {} : { textWrap }),
    paragraphs,
  };
}

export type BoxCoordinateMapper = (box: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}) => {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

/**
 * Reads a Transitional `<p:sp>` native shape (preset or custom geometry, fills, borders, effects,
 * transforms, hyperlink, and optional rich text body) into an editable `ShapeNode`.
 */
export function readShapeSp(
  context: ShapeTreeImportContext,
  mapBox?: BoxCoordinateMapper,
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

  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    cNvSpPr.consumeAttr("txBox");
  }
  consumeNonPlaceholderNvPr(nvSpPr);

  const location: ImportLocation = {
    part,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };

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

  const visuals = readSpPrVisuals(spPr, pkg, location, diagnostics);
  if (visuals === undefined) {
    return { diagnostics };
  }

  const mappedBox = mapBox
    ? mapBox({
        x: visuals.transform.x,
        y: visuals.transform.y,
        width: visuals.transform.width,
        height: visuals.transform.height,
      })
    : {
        x: visuals.transform.x,
        y: visuals.transform.y,
        width: visuals.transform.width,
        height: visuals.transform.height,
      };

  const txBody = cursor.child(P_NS, "txBody");
  const importedText =
    txBody === undefined
      ? undefined
      : readShapeTxBody(txBody, location, diagnostics);

  const singleParagraph =
    importedText !== undefined && importedText.paragraphs.length === 1
      ? importedText.paragraphs[0]
      : undefined;
  const isSingleEmptyParagraph =
    singleParagraph !== undefined &&
    singleParagraph.runs.length === 0 &&
    singleParagraph.style.list === undefined;
  const singlePlainRun =
    singleParagraph !== undefined &&
    singleParagraph.style.list === undefined &&
    singleParagraph.runs.length === 1 &&
    singleParagraph.runs[0]!.kind === "text"
      ? singleParagraph.runs[0]!
      : undefined;
  const hasStyledEmptyParagraph =
    isSingleEmptyParagraph && singleParagraph.defRunStyle !== undefined;
  const hasEffectiveText =
    (singlePlainRun !== undefined && singlePlainRun.text !== "") ||
    (importedText !== undefined &&
      singlePlainRun === undefined &&
      !isSingleEmptyParagraph &&
      importedText.paragraphs.length > 0);
  const hasEffectiveTextBody = hasEffectiveText || hasStyledEmptyParagraph;
  const emptyParagraphTextStyle =
    hasStyledEmptyParagraph && singleParagraph !== undefined
      ? (() => {
          const { list: _list, ...restPStyle } = singleParagraph.style;
          return {
            ...restPStyle,
            ...singleParagraph.defRunStyle,
          };
        })()
      : undefined;

  const { border: rawBorder, ...restVisualStyle } = visuals.style;
  if (visuals.borderArrows !== undefined) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location,
      namespaceUri: A_NS,
      localName:
        visuals.borderArrows.startArrow !== undefined ? "headEnd" : "tailEnd",
      mechanismHint: "shape-line-arrow",
    });
  }
  const effectiveVisualStyle =
    rawBorder === "none" ? restVisualStyle : visuals.style;

  const singlePlainTextStyle: Omit<ParagraphStyle, "list"> | undefined =
    hasEffectiveText &&
    singlePlainRun !== undefined &&
    singleParagraph !== undefined
      ? (() => {
          const { list: _list, ...restPStyle } = singleParagraph.style;
          return {
            ...restPStyle,
            ...singlePlainRun.style,
          };
        })()
      : undefined;

  const shapeStyle: NonNullable<ShapeDeclarationProps["style"]> = {
    position: "absolute",
    left: mappedBox.x,
    top: mappedBox.y,
    width: mappedBox.width,
    height: mappedBox.height,
    ...effectiveVisualStyle,
    ...(hasEffectiveTextBody && importedText?.paddingTop !== undefined
      ? { paddingTop: importedText.paddingTop }
      : {}),
    ...(hasEffectiveTextBody && importedText?.paddingRight !== undefined
      ? { paddingRight: importedText.paddingRight }
      : {}),
    ...(hasEffectiveTextBody && importedText?.paddingBottom !== undefined
      ? { paddingBottom: importedText.paddingBottom }
      : {}),
    ...(hasEffectiveTextBody && importedText?.paddingLeft !== undefined
      ? { paddingLeft: importedText.paddingLeft }
      : {}),
    ...(hasEffectiveTextBody && importedText?.verticalAlign !== undefined
      ? { verticalAlign: importedText.verticalAlign }
      : {}),
    ...(hasEffectiveTextBody && importedText?.textWrap !== undefined
      ? { textWrap: importedText.textWrap }
      : {}),
    ...singlePlainTextStyle,
    ...emptyParagraphTextStyle,
  };

  const props: ShapeDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    ...(visuals.geometry === undefined ? {} : { geometry: visuals.geometry }),
    style: shapeStyle,
  };

  const node = createShapeNode(props);
  if (hasStyledEmptyParagraph) {
    node.content("");
  } else if (hasEffectiveText && singlePlainRun !== undefined) {
    node.content(singlePlainRun.text);
  } else if (
    hasEffectiveText &&
    importedText !== undefined &&
    singlePlainRun === undefined &&
    importedText.paragraphs.length > 0
  ) {
    const paragraphNodes = importedText.paragraphs.map((para) => {
      const pProps: ParagraphDeclarationProps =
        Object.keys(para.style).length === 0 ? {} : { style: para.style };
      const pNode = createParagraphNode(pProps);
      if (
        para.runs.length === 1 &&
        para.runs[0]!.kind === "text" &&
        Object.keys(para.runs[0]!.style).length === 0
      ) {
        pNode.content(para.runs[0]!.text);
      } else {
        pNode.content(
          ...para.runs.map((run) => {
            if (run.kind === "slidenum") {
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
    node.content(...paragraphNodes);
  }

  attachDeferredSlideLink(node, deferredSlideLink, context);
  return { node, diagnostics };
}

export const presetShapeImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: isPresetShapeSpElement,
  read: (context) => readShapeSp(context),
};

export const customGeometryShapeImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: isCustomGeometryShapeSpElement,
  read: (context) => readShapeSp(context),
};
