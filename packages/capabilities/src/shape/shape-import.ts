import {
  assertValidLang,
  type LinkTarget,
  type TextAlignment,
  type TextStyle,
} from "@pptxdk/core/capability";
import {
  A_NS,
  OoxmlBoolSchema,
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  OoxmlPositiveIntSchema,
  P_NS,
  REL_SLIDE,
  R_NS,
  XML_NS,
  emuToPx,
  parseXmlPart,
  resolvePartPath,
  type ImportDiagnostic,
  type ImportLocation,
  type PptxPackage,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

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
import { isPlaceholderSpElement } from "../text/text-import.ts";
import {
  createShapeNode,
  type ShapeDeclarationProps,
} from "./shape-declaration.ts";

const NonVisualIdAttrSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

const BodyPrAttrSchema = z.object({
  wrap: z.enum(["square", "none"]).optional(),
  rtlCol: OoxmlBoolSchema.optional(),
  lIns: OoxmlNonNegativeIntSchema.optional(),
  tIns: OoxmlNonNegativeIntSchema.optional(),
  rIns: OoxmlNonNegativeIntSchema.optional(),
  bIns: OoxmlNonNegativeIntSchema.optional(),
  anchor: z.enum(["t", "ctr", "b"]).optional(),
});

const ParagraphPrAttrSchema = z.object({
  algn: z.enum(["l", "ctr", "r", "just"]).optional(),
  rtl: OoxmlBoolSchema.optional(),
  lvl: OoxmlNonNegativeIntSchema.optional(),
});

const RunPrAttrSchema = z.object({
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

const TypefaceAttrSchema = z.object({
  typeface: z.string().min(1),
});

const BulletCharAttrSchema = z.object({
  char: z.string().min(1),
});

const BulletAutoNumAttrSchema = z.object({
  type: z.string().min(1),
  startAt: OoxmlIntSchema.optional(),
});

const SpacingValAttrSchema = z.object({
  val: OoxmlNonNegativeIntSchema,
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

/** Side-effect-free predicate matching non-textbox preset shapes (`<p:sp>` without `txBox="1"`, `<p:ph>`, or `<a:custGeom>`). */
export function isPresetShapeSpElement(cursor: XmlElementCursor): boolean {
  return (
    !isPlaceholderSpElement(cursor) &&
    !isTextBoxSpElement(cursor) &&
    !isCustomGeomSpElement(cursor)
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

function resolveTargetSlideAuthorId(
  pkg: PptxPackage,
  targetSlidePart: string,
): string | undefined {
  const xmlText = pkg.readTextPart(targetSlidePart);
  if (xmlText === undefined) return undefined;
  let slideEl: XmlElement;
  try {
    slideEl = parseXmlPart(targetSlidePart, xmlText);
  } catch {
    return undefined;
  }
  const cSld = directElementChild(slideEl, P_NS, "cSld");
  const explicitName = cSld?.getAttribute("name")?.trim();
  if (explicitName !== undefined && explicitName.length > 0) {
    return explicitName;
  }
  const relsById = new Map(
    pkg.presentationRels.map((entry) => [entry.Id, entry]),
  );
  const sldIdLst = directElementChild(
    pkg.presentationRoot.element,
    P_NS,
    "sldIdLst",
  );
  if (sldIdLst === undefined) return undefined;
  let index = 0;
  for (
    let child = sldIdLst.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if (
      (el.namespaceURI ?? "") !== P_NS ||
      (el.localName ?? el.tagName) !== "sldId"
    ) {
      continue;
    }
    const rId = el.getAttributeNS(R_NS, "id");
    const rel = rId === null ? undefined : relsById.get(rId);
    const part =
      rel === undefined
        ? undefined
        : resolvePartPath(pkg.presentationPath, rel.Target);
    if (part === targetSlidePart) {
      return `slide-${index + 1}`;
    }
    index += 1;
  }
  return undefined;
}

/** Reads `<a:hlinkClick>` from `<p:cNvPr>` into a `LinkTarget` while isolating broken links. */
export function readHyperlinkClick(
  cNvPrCursor: XmlElementCursor,
  pkg: PptxPackage,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): LinkTarget | undefined {
  const hlink = cNvPrCursor.child(A_NS, "hlinkClick");
  if (hlink === undefined) return undefined;
  const rId = hlink.attr("id", R_NS);
  const action = hlink.attr("action");
  if (action === "ppaction://hlinkshowjump?jump=firstslide") {
    return { action: "home" };
  }
  if (rId === undefined) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: hlink.xpath },
      namespaceUri: A_NS,
      localName: "hlinkClick",
      mechanismHint: action ?? "hlinkClick",
    });
    return undefined;
  }
  const rels = pkg.readRelationships(location.part);
  diagnostics.push(...rels.diagnostics);
  const rel = rels.entries.find((entry) => entry.Id === rId);
  if (rel === undefined) {
    diagnostics.push({
      code: "broken-relationship",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: hlink.xpath },
      relationshipId: rId,
    });
    return undefined;
  }
  if (rel.TargetMode === "External") {
    return { url: rel.Target };
  }
  if (rel.Type === REL_SLIDE) {
    const targetPart = resolvePartPath(location.part, rel.Target);
    if (targetPart === undefined || !pkg.hasPart(targetPart)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: hlink.xpath },
        relationshipId: rId,
        targetPart: targetPart ?? rel.Target,
      });
      return undefined;
    }
    const targetId = resolveTargetSlideAuthorId(pkg, targetPart);
    if (targetId !== undefined) {
      return { id: targetId };
    }
  }
  diagnostics.push({
    code: "unsupported-mechanism",
    category: "unsupported",
    disposition: "omitted",
    location: { ...location, xpath: hlink.xpath },
    namespaceUri: A_NS,
    localName: "hlinkClick",
    mechanismHint: rel.Type,
  });
  return undefined;
}

/** Reads `<p:cNvPr>` id, name, and hyperlink without assuming shape vs group wrapper tag. */
export function readNonVisualCommonProperties(
  cNvPrCursor: XmlElementCursor | undefined,
  pkg: PptxPackage,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): {
  readonly objectId?: number;
  readonly objectName?: string;
  readonly link?: LinkTarget;
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
      location: { part, xpath: cNvPrCursor.xpath, slideIndex },
      attributeOrElement: "p:cNvPr",
      rawValue: JSON.stringify(raw),
    });
  } else {
    objectId = parsed.data.id;
    if (parsed.data.name !== undefined && parsed.data.name.trim() !== "") {
      objectName = parsed.data.name;
    }
  }
  const location: ImportLocation = {
    part,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };
  const link = readHyperlinkClick(cNvPrCursor, pkg, location, diagnostics);
  return {
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
    ...(link === undefined ? {} : { link }),
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

  let fontFamily: NonNullable<TextStyle["fontFamily"]> | undefined;
  const latin = rPrCursor.child(A_NS, "latin");
  const ea = rPrCursor.child(A_NS, "ea");
  const cs = rPrCursor.child(A_NS, "cs");
  const fontCursor = latin ?? ea ?? cs;
  if (fontCursor !== undefined) {
    const parsedFont = TypefaceAttrSchema.safeParse(
      fontCursor.attrsObject(["typeface"]),
    );
    if (parsedFont.success) {
      const tf = parsedFont.data.typeface;
      if (tf.startsWith("+mj-")) fontFamily = "major";
      else if (tf.startsWith("+mn-")) fontFamily = "minor";
      else fontFamily = tf;
    }
  }
  if (ea !== undefined && ea !== fontCursor) ea.consumeAttr("typeface");
  if (cs !== undefined && cs !== fontCursor) cs.consumeAttr("typeface");

  return {
    ...(fontFamily === undefined ? {} : { fontFamily }),
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
      child.consumeAttr("lang");
      child.consumeAttr("sz");
      child.consumeAttr("b");
      child.consumeAttr("i");
      child.consumeAttr("dirty");
      child.consumeAttr("smtClean");
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

  return { runs, style: pStyle };
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
  const { objectId, objectName, link } = readNonVisualCommonProperties(
    cNvPr,
    pkg,
    part,
    slideIndex,
    diagnostics,
  );

  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    cNvSpPr.consumeAttr("txBox");
  }
  nvSpPr.ignoreEmptyChild(P_NS, "nvPr");

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
  const hasEffectiveText =
    (singlePlainRun !== undefined && singlePlainRun.text !== "") ||
    (importedText !== undefined &&
      singlePlainRun === undefined &&
      !isSingleEmptyParagraph &&
      importedText.paragraphs.length > 0);

  const shapeStyle: NonNullable<ShapeDeclarationProps["style"]> = {
    ...visuals.style,
    ...(hasEffectiveText && importedText?.paddingTop !== undefined
      ? { paddingTop: importedText.paddingTop }
      : {}),
    ...(hasEffectiveText && importedText?.paddingRight !== undefined
      ? { paddingRight: importedText.paddingRight }
      : {}),
    ...(hasEffectiveText && importedText?.paddingBottom !== undefined
      ? { paddingBottom: importedText.paddingBottom }
      : {}),
    ...(hasEffectiveText && importedText?.paddingLeft !== undefined
      ? { paddingLeft: importedText.paddingLeft }
      : {}),
    ...(hasEffectiveText && importedText?.verticalAlign !== undefined
      ? { verticalAlign: importedText.verticalAlign }
      : {}),
    ...(hasEffectiveText && importedText?.textWrap !== undefined
      ? { textWrap: importedText.textWrap }
      : {}),
  };

  const singleTextStyle: TextStyle | undefined =
    hasEffectiveText &&
    singlePlainRun !== undefined &&
    singleParagraph !== undefined
      ? {
          ...singleParagraph.style,
          ...singlePlainRun.style,
        }
      : undefined;

  const props: ShapeDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    ...(visuals.geometry === undefined ? {} : { geometry: visuals.geometry }),
    ...(visuals.transform.flipH ? { flipH: true } : {}),
    ...(visuals.transform.flipV ? { flipV: true } : {}),
    x: mappedBox.x,
    y: mappedBox.y,
    width: mappedBox.width,
    height: mappedBox.height,
    ...(Object.keys(shapeStyle).length === 0 ? {} : { style: shapeStyle }),
    ...(hasEffectiveText && singlePlainRun !== undefined
      ? { text: singlePlainRun.text }
      : {}),
    ...(singleTextStyle !== undefined && Object.keys(singleTextStyle).length > 0
      ? { textStyle: singleTextStyle }
      : {}),
  };

  const node = createShapeNode(props);
  if (
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
