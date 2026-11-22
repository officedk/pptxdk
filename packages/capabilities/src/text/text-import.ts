import {
  assertValidLang,
  type ColorValue,
  type NativeColorSlot,
  type TextAlignment,
  type TextStyle,
} from "@pptxdk/core/capability";
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
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  createParagraphNode,
  createSpanNode,
  createTextNode,
  type ParagraphDeclarationProps,
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
});

const TypefaceFragmentSchema = z.object({
  typeface: z.string().min(1),
});

const AlphaModifierSchema = z.object({
  val: OoxmlNonNegativeIntSchema,
});

const SrgbColorSchema = z.object({
  val: HexRgb6Schema,
});

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

/** Side-effect-free predicate distinguishing `<p:sp>` text boxes (`txBox="1"`) from shapes. */
function isTextBoxSpElement(cursor: XmlElementCursor): boolean {
  const nvSpPr = directElementChild(cursor.element, P_NS, "nvSpPr");
  if (nvSpPr === undefined) return false;
  const cNvSpPr = directElementChild(nvSpPr, P_NS, "cNvSpPr");
  if (cNvSpPr === undefined) return false;
  const txBox = cNvSpPr.getAttribute("txBox");
  return txBox === "1" || txBox === "true";
}

function readSolidFillColor(
  fillCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): ColorValue | undefined {
  const srgb = fillCursor.child(A_NS, "srgbClr");
  const scheme =
    srgb === undefined ? fillCursor.child(A_NS, "schemeClr") : undefined;
  const colorCursor = srgb ?? scheme;
  if (colorCursor === undefined) return undefined;

  let baseColor: string | undefined;
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
    baseColor = `#${parsed.data.val.toUpperCase()}`;
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
    baseColor = schemeToNativeSlot[parsed.data.val];
  }

  if (baseColor === undefined) return undefined;

  const alphaCursor = colorCursor.child(A_NS, "alpha");
  if (alphaCursor !== undefined) {
    const rawAlpha = alphaCursor.attrsObject(["val"]);
    const parsedAlpha = AlphaModifierSchema.safeParse(rawAlpha);
    if (parsedAlpha.success) {
      return { color: baseColor, alpha: parsedAlpha.data.val / 100_000 };
    }
  }
  return baseColor;
}

type ParsedRunStyle = NonNullable<SpanDeclarationProps["style"]>;

function readRunProperties(
  rPrCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
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
  ]);
  const parsed = RunPrFragmentSchema.safeParse(raw);
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

  let color: NonNullable<TextStyle["color"]> | undefined;
  const solidFill = rPrCursor.child(A_NS, "solidFill");
  if (solidFill !== undefined) {
    color = readSolidFillColor(solidFill, part, slideIndex, diagnostics);
  }

  let fontFamily: NonNullable<TextStyle["fontFamily"]> | undefined;
  const latin = rPrCursor.child(A_NS, "latin");
  const ea = rPrCursor.child(A_NS, "ea");
  const cs = rPrCursor.child(A_NS, "cs");
  const fontCursor = latin ?? ea ?? cs;
  if (latin !== undefined) {
    const rawLatin = latin.attrsObject(["typeface"]);
    const parsedLatin = TypefaceFragmentSchema.safeParse(rawLatin);
    if (
      parsedLatin.success &&
      !parsedLatin.data.typeface.startsWith("+mj-") &&
      !parsedLatin.data.typeface.startsWith("+mn-")
    ) {
      fontFamily = parsedLatin.data.typeface;
    }
  } else if (fontCursor !== undefined) {
    const rawFont = fontCursor.attrsObject(["typeface"]);
    const parsedFont = TypefaceFragmentSchema.safeParse(rawFont);
    if (
      parsedFont.success &&
      !parsedFont.data.typeface.startsWith("+mj-") &&
      !parsedFont.data.typeface.startsWith("+mn-")
    ) {
      fontFamily = parsedFont.data.typeface;
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
    ...(lang === undefined ? {} : { lang }),
  };
}

type ImportedRun = {
  readonly text: string;
  readonly style: ParsedRunStyle;
};

type ImportedParagraph = {
  readonly runs: readonly ImportedRun[];
  readonly style: NonNullable<ParagraphDeclarationProps["style"]>;
};

function readParagraph(
  pCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): ImportedParagraph {
  const pStyle: {
    align?: TextAlignment;
  } = {};
  const pPr = pCursor.child(A_NS, "pPr");
  if (pPr !== undefined) {
    const rawPPr = pPr.attrsObject(["algn", "rtl"]);
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
    } else if (parsedPPr.data.algn !== undefined) {
      pStyle.align = algnToTextAlign[parsedPPr.data.algn];
    }
  }

  const runs: ImportedRun[] = [];
  for (const child of pCursor.remainingElementChildren()) {
    if (child.namespaceUri === A_NS && child.localName === "r") {
      const rPr = child.child(A_NS, "rPr");
      const runStyle =
        rPr === undefined
          ? {}
          : readRunProperties(rPr, part, slideIndex, diagnostics);
      const t = child.child(A_NS, "t");
      if (t !== undefined) {
        t.consumeAttr("space", XML_NS);
        runs.push({ text: t.textContent, style: runStyle });
      }
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "br") {
      const rPr = child.child(A_NS, "rPr");
      const runStyle =
        rPr === undefined
          ? {}
          : readRunProperties(rPr, part, slideIndex, diagnostics);
      runs.push({ text: "\n", style: runStyle });
      continue;
    }
    if (child.namespaceUri === A_NS && child.localName === "endParaRPr") {
      child.consumeAttr("lang");
      child.consumeAttr("sz");
      child.consumeAttr("b");
      child.consumeAttr("i");
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

  return { runs, style: pStyle };
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
  nvSpPr.ignoreEmptyChild(P_NS, "nvPr");

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
    backgroundColor = readSolidFillColor(
      spSolidFill,
      part,
      slideIndex,
      diagnostics,
    );
  }

  let paddingTop: number | undefined;
  let paddingRight: number | undefined;
  let paddingBottom: number | undefined;
  let paddingLeft: number | undefined;
  let verticalAlign: NonNullable<TextStyle["verticalAlign"]> | undefined;
  let textWrap: NonNullable<TextStyle["textWrap"]> | undefined;
  const rotate =
    parsedXfrm.success &&
    parsedXfrm.data.rot !== undefined &&
    parsedXfrm.data.rot !== 0
      ? parsedXfrm.data.rot / 60_000
      : undefined;

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
  }

  txBody.ignoreEmptyChild(A_NS, "lstStyle");

  const importedParagraphs = txBody
    .children(A_NS, "p")
    .map((pCursor) => readParagraph(pCursor, part, slideIndex, diagnostics));

  const singleParagraph =
    importedParagraphs.length === 1 ? importedParagraphs[0] : undefined;
  const singleRun =
    singleParagraph !== undefined && singleParagraph.runs.length === 1
      ? singleParagraph.runs[0]
      : undefined;

  const boxStyle: NonNullable<TextDeclarationProps["style"]> = {
    ...(paddingTop === undefined ? {} : { paddingTop }),
    ...(paddingRight === undefined ? {} : { paddingRight }),
    ...(paddingBottom === undefined ? {} : { paddingBottom }),
    ...(paddingLeft === undefined ? {} : { paddingLeft }),
    ...(verticalAlign === undefined ? {} : { verticalAlign }),
    ...(textWrap === undefined ? {} : { textWrap }),
    ...(rotate === undefined ? {} : { rotate }),
    ...(backgroundColor === undefined ? {} : { backgroundColor }),
    ...(singleParagraph === undefined ? {} : singleParagraph.style),
    ...(singleRun === undefined ? {} : singleRun.style),
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
  if (singleRun !== undefined) {
    node.content(singleRun.text);
  } else if (importedParagraphs.length > 0) {
    const paragraphNodes = importedParagraphs.map((para) => {
      const pNode = createParagraphNode(
        Object.keys(para.style).length === 0 ? {} : { style: para.style },
      );
      if (
        para.runs.length === 1 &&
        Object.keys(para.runs[0]!.style).length === 0
      ) {
        pNode.content(para.runs[0]!.text);
      } else {
        pNode.content(
          ...para.runs.map((run) =>
            Object.keys(run.style).length === 0
              ? run.text
              : createSpanNode({ style: run.style }).content(run.text),
          ),
        );
      }
      return pNode;
    });
    node.content(...paragraphNodes);
  }

  return { node, diagnostics };
}

export const textBoxImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: isTextBoxSpElement,
  read: readTextBoxSp,
};
