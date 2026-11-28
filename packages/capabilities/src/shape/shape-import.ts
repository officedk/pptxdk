import {
  type LinkTarget,
  type TextAlignment,
  type TextStyle,
} from "@pptxdk/core/capability";
import {
  A_NS,
  normalizeBulletCharacter,
  ooxmlAutoNumStartAtSchema,
  ooxmlBoolSchema,
  ooxmlLineSpacingPercentSchema,
  ooxmlLineSpacingPointsSchema,
  ooxmlListLevelSchema,
  ooxmlNonNegativeIntSchema,
  ooxmlParagraphSpacingPointsSchema,
  P_NS,
  readOoxmlBoxInsets,
  readOoxmlRunAttrStyles,
  readRunFontSlots,
  XML_NS,
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
  createSubNode,
  createSupNode,
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
  anchor: z.enum(["t", "ctr", "b"]).optional(),
});

const ParagraphAlignAttrSchema = z.enum(["l", "ctr", "r", "just"]);

const SpacingValAttrSchema = z.object({
  val: ooxmlParagraphSpacingPointsSchema,
});

const LineSpacingPtsAttrSchema = z.object({
  val: ooxmlLineSpacingPointsSchema,
});

const LineSpacingPctAttrSchema = z.object({
  val: ooxmlLineSpacingPercentSchema,
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
): { readonly style: ParsedRunStyle; readonly inlineKind?: "sup" | "sub" } {
  const attrStyles = readOoxmlRunAttrStyles({
    rPrCursor,
    location,
    diagnostics,
    allowBaseline: true,
    allowNoneTextDecoration: false,
    diagnoseUnsupportedEnums: true,
  });

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
    style: {
      ...(fontFamily === undefined ? {} : { fontFamily }),
      ...(eastAsianFontFamily === undefined ? {} : { eastAsianFontFamily }),
      ...(attrStyles.fontSize === undefined
        ? {}
        : { fontSize: attrStyles.fontSize }),
      ...(attrStyles.fontWeight === undefined
        ? {}
        : { fontWeight: attrStyles.fontWeight }),
      ...(attrStyles.fontStyle === undefined
        ? {}
        : { fontStyle: attrStyles.fontStyle }),
      ...(color === undefined ? {} : { color }),
      ...(attrStyles.letterSpacing === undefined
        ? {}
        : { letterSpacing: attrStyles.letterSpacing }),
      ...(attrStyles.textTransform === undefined
        ? {}
        : { textTransform: attrStyles.textTransform }),
      ...(attrStyles.textDecoration === undefined
        ? {}
        : { textDecoration: attrStyles.textDecoration }),
      ...(effects.boxShadow === undefined
        ? {}
        : { textShadow: effects.boxShadow }),
      ...(attrStyles.lang === undefined ? {} : { lang: attrStyles.lang }),
      ...(attrStyles.baselineShift === undefined
        ? {}
        : { baselineShift: attrStyles.baselineShift }),
    },
    ...(attrStyles.inlineKind === undefined
      ? {}
      : { inlineKind: attrStyles.inlineKind }),
  };
}

type ImportedShapeRun =
  | {
      readonly kind: "text" | "sup" | "sub";
      readonly text: string;
      readonly style: ParsedRunStyle;
    }
  | { readonly kind: "slidenum"; readonly style: ParsedRunStyle };

type ImportedShapeParagraph = {
  readonly runs: readonly ImportedShapeRun[];
  readonly style: ParagraphStyle;
  readonly defRunStyle?: Omit<ParsedRunStyle, "baselineShift">;
};

function readLineSpacing(
  lnSpcCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ParagraphStyle["lineHeight"] | undefined {
  const spcPts = lnSpcCursor.child(A_NS, "spcPts");
  if (spcPts !== undefined) {
    const raw = spcPts.attrsObject(["val"]);
    const parsed = LineSpacingPtsAttrSchema.safeParse(raw);
    if (!parsed.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: spcPts.xpath },
        attributeOrElement: "a:spcPts/@val",
        rawValue: raw["val"] ?? "",
      });
    } else {
      return parsed.data.val / 75;
    }
  }
  const spcPct = lnSpcCursor.child(A_NS, "spcPct");
  if (spcPct !== undefined) {
    const raw = spcPct.attrsObject(["val"]);
    const parsed = LineSpacingPctAttrSchema.safeParse(raw);
    if (!parsed.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: spcPct.xpath },
        attributeOrElement: "a:spcPct/@val",
        rawValue: raw["val"] ?? "",
      });
    } else {
      return `${parsed.data.val / 1000}%`;
    }
  }
  return undefined;
}

function readShapeParagraph(
  pCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ImportedShapeParagraph {
  let align: TextAlignment | undefined;
  let lineHeight: ParagraphStyle["lineHeight"] | undefined;
  let spacingAfter: number | undefined;
  let list: TextList | undefined;
  let defRunStyle: Omit<ParsedRunStyle, "baselineShift"> | undefined;
  let level = 0;

  const pPr = pCursor.child(A_NS, "pPr");
  if (pPr !== undefined) {
    const rawPPr = pPr.attrsObject(["algn", "rtl", "lvl", "marL", "indent"]);
    if (rawPPr["algn"] !== undefined) {
      const parsedAlgn = ParagraphAlignAttrSchema.safeParse(rawPPr["algn"]);
      if (parsedAlgn.success) {
        align = algnToTextAlign[parsedAlgn.data];
      } else {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: { ...location, xpath: `${pPr.xpath}/@algn` },
          attributeOrElement: "@algn",
          rawValue: rawPPr["algn"],
        });
      }
    }
    if (rawPPr["lvl"] !== undefined) {
      const parsedLvl = ooxmlListLevelSchema.safeParse(rawPPr["lvl"]);
      if (parsedLvl.success) {
        level = parsedLvl.data;
      } else {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: { ...location, xpath: `${pPr.xpath}/@lvl` },
          attributeOrElement: "@lvl",
          rawValue: rawPPr["lvl"],
        });
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
        const raw = spcPts.attrsObject(["val"]);
        const parsed = SpacingValAttrSchema.safeParse(raw);
        if (parsed.success) {
          spacingAfter = parsed.data.val / 75;
        } else {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "omitted",
            location: { ...location, xpath: spcPts.xpath },
            attributeOrElement: "a:spcPts/@val",
            rawValue: raw["val"] ?? "",
          });
        }
      }
    }

    pPr.ignoreEmptyChild(A_NS, "buNone");
    const buChar = pPr.child(A_NS, "buChar");
    if (buChar !== undefined) {
      const rawBu = buChar.attrsObject(["char"]);
      const normalized = normalizeBulletCharacter(rawBu["char"]);
      if (normalized.recovered) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: { ...location, xpath: buChar.xpath },
          attributeOrElement: "a:buChar/@char",
          rawValue: rawBu["char"] ?? "",
        });
      }
      list = {
        type: "bullet",
        character: normalized.character,
        ...(level > 0 ? { level } : {}),
      };
    } else {
      const buAutoNum = pPr.child(A_NS, "buAutoNum");
      if (buAutoNum !== undefined) {
        const rawAuto = buAutoNum.attrsObject(["type", "startAt"]);
        const rawType = rawAuto["type"];
        if (rawType === undefined || rawType.length === 0) {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "omitted",
            location: { ...location, xpath: buAutoNum.xpath },
            attributeOrElement: "a:buAutoNum",
            rawValue: JSON.stringify(rawAuto),
          });
        } else {
          let startAt: number | undefined;
          if (rawAuto["startAt"] !== undefined) {
            const parsedStart = ooxmlAutoNumStartAtSchema.safeParse(
              rawAuto["startAt"],
            );
            if (parsedStart.success) {
              startAt = parsedStart.data;
            } else {
              diagnostics.push({
                code: "invalid-xml-value",
                category: "corrupt-input",
                disposition: "omitted",
                location: { ...location, xpath: `${buAutoNum.xpath}/@startAt` },
                attributeOrElement: "a:buAutoNum/@startAt",
                rawValue: rawAuto["startAt"],
              });
            }
          }
          list = {
            type: "number",
            style: "arabicPeriod",
            ...(startAt === undefined ? {} : { startAt }),
            ...(level > 0 ? { level } : {}),
          };
        }
      }
    }

    const defRPr = pPr.child(A_NS, "defRPr");
    if (defRPr !== undefined) {
      const { baselineShift: _baselineShift, ...restDefRunStyle } =
        readShapeRunProperties(defRPr, location, diagnostics).style;
      defRunStyle = restDefRunStyle;
    }
  }

  const runs: ImportedShapeRun[] = [];
  for (const child of pCursor.remainingElementChildren()) {
    if (child.namespaceUri === A_NS && child.localName === "r") {
      const rPr = child.child(A_NS, "rPr");
      const parsedRPr =
        rPr === undefined
          ? { style: {} }
          : readShapeRunProperties(rPr, location, diagnostics);
      const t = child.child(A_NS, "t");
      if (t !== undefined) {
        t.consumeAttr("space", XML_NS);
        runs.push({
          kind: parsedRPr.inlineKind ?? "text",
          text: t.textContent,
          style: parsedRPr.style,
        });
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "fld") {
      child.consumeAttr("id");
      const fieldType = child.attr("type");
      const rPr = child.child(A_NS, "rPr");
      const parsedRPr =
        rPr === undefined
          ? { style: {} }
          : readShapeRunProperties(rPr, location, diagnostics);
      const t = child.child(A_NS, "t");
      if (t !== undefined) {
        t.consumeAttr("space", XML_NS);
      }
      if (fieldType === "slidenum") {
        const fieldStyle: ParsedRunStyle =
          parsedRPr.inlineKind === undefined
            ? parsedRPr.style
            : {
                ...parsedRPr.style,
                baselineShift: parsedRPr.inlineKind === "sup" ? "30%" : "-25%",
              };
        runs.push({ kind: "slidenum", style: fieldStyle });
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
          runs.push({
            kind: parsedRPr.inlineKind ?? "text",
            text: t.textContent,
            style: parsedRPr.style,
          });
        }
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "br") {
      const rPr = child.child(A_NS, "rPr");
      const parsedRPr =
        rPr === undefined
          ? { style: {} }
          : readShapeRunProperties(rPr, location, diagnostics);
      runs.push({
        kind: parsedRPr.inlineKind ?? "text",
        text: "\n",
        style: parsedRPr.style,
      });
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
    }
    const insets = readOoxmlBoxInsets({
      rawAttrs: rawBodyPr,
      attrNames: {
        left: "lIns",
        top: "tIns",
        right: "rIns",
        bottom: "bIns",
      },
      elementXpath: bodyPr.xpath,
      location,
      diagnostics,
    });
    paddingLeft = insets.paddingLeft;
    paddingTop = insets.paddingTop;
    paddingRight = insets.paddingRight;
    paddingBottom = insets.paddingBottom;
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
    singleParagraph.runs[0]!.kind === "text" &&
    singleParagraph.runs[0]!.style.baselineShift === undefined
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
          const { baselineShift: _baselineShift, ...restRunStyle } =
            singlePlainRun.style;
          return {
            ...restPStyle,
            ...restRunStyle,
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
