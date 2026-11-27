import type {
  ColorValue,
  Geometry,
  GradientStop,
  HexColor,
  LinearGradient,
  RadialGradient,
  SchemeColorName,
} from "@pptxdk/core/capability";
import {
  presetGeometryNameSchema,
  schemeColorNameSchema,
} from "@pptxdk/core/capability";
import { z } from "zod";

import type { ImportDiagnostic, ImportLocation } from "#/diagnostics.ts";
import type { ImportedPlaceholderSpPr } from "#/mechanism.ts";
import { A_NS, R_NS, emuToPx } from "#/namespaces.ts";
import { resolvePartPath, type PptxPackage } from "#/opc.ts";
import {
  extentEmuFragmentSchema,
  hexRgb6Schema,
  offsetEmuFragmentSchema,
  ooxmlIntSchema,
  ooxmlNonNegativeIntSchema,
  transformAttributesFragmentSchema,
} from "#/xml-fragments.ts";
import type { XmlElementCursor } from "#/xml-reader.ts";

export const arrowTypes = [
  "none",
  "triangle",
  "stealth",
  "arrow",
  "diamond",
  "oval",
] as const;

export type ArrowType = (typeof arrowTypes)[number];

export const arrowTypeSchema: z.ZodType<ArrowType> = z.enum(arrowTypes);

const percentValAttrSchema = z.object({
  val: ooxmlNonNegativeIntSchema,
});

const alphaModFixAttrSchema = z.object({
  amt: ooxmlNonNegativeIntSchema,
});

const srgbColorAttrSchema = z.object({
  val: hexRgb6Schema,
});

const schemeColorAttrSchema = z.object({
  val: schemeColorNameSchema,
});

const sysColorAttrSchema = z.object({
  val: z.string().min(1),
  lastClr: hexRgb6Schema.optional(),
});

const gradStopAttrSchema = z.object({
  pos: ooxmlNonNegativeIntSchema,
});

const gradLinAttrSchema = z.object({
  ang: ooxmlIntSchema.optional(),
  scaled: z.enum(["0", "1", "true", "false"]).optional(),
});

const lineAttrSchema = z.object({
  w: ooxmlNonNegativeIntSchema.optional(),
  cap: z.enum(["rnd", "sq", "flat"]).optional(),
  cmpd: z.enum(["sng", "dbl", "thickThin", "thinThick", "tri"]).optional(),
  algn: z.enum(["ctr", "in"]).optional(),
});

const dashPresetAttrSchema = z.object({
  val: z.string().min(1),
});

const lineEndAttrSchema = z.object({
  type: arrowTypeSchema.optional(),
  w: z.enum(["sm", "med", "lg"]).optional(),
  len: z.enum(["sm", "med", "lg"]).optional(),
});

const outerShadowAttrSchema = z.object({
  blurRad: ooxmlNonNegativeIntSchema.optional(),
  dist: ooxmlNonNegativeIntSchema.optional(),
  dir: ooxmlIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: z.enum(["0", "1", "true", "false"]).optional(),
});

const softEdgeAttrSchema = z.object({
  rad: ooxmlNonNegativeIntSchema,
});

/** Reads a DrawingML color child (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`) with optional alpha and luminance transforms. */
export function readDrawingColor(
  parentCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ColorValue | undefined {
  let colorCursor: XmlElementCursor | undefined;
  let baseHex: HexColor | undefined;
  let slot: SchemeColorName | undefined;

  for (const child of parentCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    if (child.localName === "srgbClr" && colorCursor === undefined) {
      colorCursor = child;
      const raw = child.attrsObject(["val"]);
      const parsed = srgbColorAttrSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "a:srgbClr/@val",
          rawValue: raw["val"] ?? "",
        });
        return undefined;
      }
      baseHex = `#${parsed.data.val.toUpperCase()}`;
    } else if (child.localName === "schemeClr" && colorCursor === undefined) {
      colorCursor = child;
      const raw = child.attrsObject(["val"]);
      const parsed = schemeColorAttrSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "a:schemeClr/@val",
          rawValue: raw["val"] ?? "",
        });
        return undefined;
      }
      slot = parsed.data.val;
    } else if (child.localName === "sysClr" && colorCursor === undefined) {
      colorCursor = child;
      const raw = child.attrsObject(["val", "lastClr"]);
      const parsed = sysColorAttrSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "a:sysClr",
          rawValue: JSON.stringify(raw),
        });
        return undefined;
      }
      const fallback =
        parsed.data.val === "windowText"
          ? "000000"
          : parsed.data.val === "window"
            ? "FFFFFF"
            : undefined;
      const hex = parsed.data.lastClr ?? fallback;
      if (hex === undefined) return undefined;
      baseHex = `#${hex.toUpperCase()}`;
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `color:${child.localName}`,
      });
    }
  }

  if (colorCursor === undefined) return undefined;

  let alpha: number | undefined;
  let lumMod: number | undefined;
  let lumOff: number | undefined;

  for (const mod of colorCursor.remainingElementChildren()) {
    if (mod.namespaceUri === A_NS && mod.localName === "alpha") {
      const parsed = percentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        alpha = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
    } else if (mod.namespaceUri === A_NS && mod.localName === "lumMod") {
      const parsed = percentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        lumMod = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
    } else if (mod.namespaceUri === A_NS && mod.localName === "lumOff") {
      const parsed = percentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        lumOff = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
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
      : baseHex;

  if (baseColor === undefined) return undefined;
  if (alpha !== undefined && alpha < 1) {
    return { color: baseColor, alpha };
  }
  return baseColor;
}

export type ImportedTransform = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly xEmu: number;
  readonly yEmu: number;
  readonly cxEmu: number;
  readonly cyEmu: number;
  readonly rotate?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
};

/** Parses `<a:xfrm>` offset, extent, rotation, and flips into pixels and raw EMU. */
export function readTransformXml(
  xfrmCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ImportedTransform | undefined {
  const rawXfrm = xfrmCursor.attrsObject(["rot", "flipH", "flipV"]);
  const parsedXfrm = transformAttributesFragmentSchema.safeParse(rawXfrm);
  const off = xfrmCursor.child(A_NS, "off");
  const ext = xfrmCursor.child(A_NS, "ext");
  if (off === undefined || ext === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: xfrmCursor.xpath },
    });
    return undefined;
  }
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
      location: { ...location, xpath: xfrmCursor.xpath },
      attributeOrElement: "a:xfrm",
    });
    return undefined;
  }
  const rotate =
    parsedXfrm.success &&
    parsedXfrm.data.rot !== undefined &&
    parsedXfrm.data.rot !== 0
      ? parsedXfrm.data.rot / 60_000
      : undefined;
  const flipH = parsedXfrm.success && parsedXfrm.data.flipH === true;
  const flipV = parsedXfrm.success && parsedXfrm.data.flipV === true;

  return {
    x: emuToPx(parsedOff.data.x),
    y: emuToPx(parsedOff.data.y),
    width: emuToPx(parsedExt.data.cx),
    height: emuToPx(parsedExt.data.cy),
    xEmu: parsedOff.data.x,
    yEmu: parsedOff.data.y,
    cxEmu: parsedExt.data.cx,
    cyEmu: parsedExt.data.cy,
    ...(rotate === undefined ? {} : { rotate }),
    ...(flipH ? { flipH: true } : {}),
    ...(flipV ? { flipV: true } : {}),
  };
}

/** Reads DrawingML `<a:gradFill>` into a `LinearGradient` or `RadialGradient`. */
export function readGradientFill(
  gradFillCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): LinearGradient | RadialGradient | undefined {
  gradFillCursor.consumeAttr("rotWithShape");
  gradFillCursor.consumeAttr("flip");
  const gsLst = gradFillCursor.child(A_NS, "gsLst");
  const stops: GradientStop[] = [];
  if (gsLst !== undefined) {
    for (const gs of gsLst.children(A_NS, "gs")) {
      const rawPos = gs.attrsObject(["pos"]);
      const parsedPos = gradStopAttrSchema.safeParse(rawPos);
      const color = readDrawingColor(gs, location, diagnostics);
      if (!parsedPos.success || color === undefined) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: gs.xpath },
          attributeOrElement: "a:gs",
          rawValue: JSON.stringify(rawPos),
        });
        continue;
      }
      stops.push({
        position: Math.min(1, Math.max(0, parsedPos.data.pos / 100_000)),
        color,
      });
    }
  }
  if (stops.length < 2) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: gradFillCursor.xpath },
      attributeOrElement: "a:gradFill/a:gsLst",
    });
    return undefined;
  }

  let result: LinearGradient | RadialGradient | undefined;
  const lin = gradFillCursor.child(A_NS, "lin");
  if (lin !== undefined) {
    const rawLin = lin.attrsObject(["ang", "scaled"]);
    const parsedLin = gradLinAttrSchema.safeParse(rawLin);
    const drawingAngle = parsedLin.success ? (parsedLin.data.ang ?? 0) : 0;
    const authorAngle = (((drawingAngle / 60_000 + 90) % 360) + 360) % 360;
    result = { type: "linear-gradient", angle: authorAngle, stops };
  } else {
    const pathEl = gradFillCursor.child(A_NS, "path");
    if (pathEl !== undefined) {
      pathEl.consumeAttr("path");
      const fillToRect = pathEl.child(A_NS, "fillToRect");
      if (fillToRect !== undefined) {
        fillToRect.consumeAttr("l");
        fillToRect.consumeAttr("t");
        fillToRect.consumeAttr("r");
        fillToRect.consumeAttr("b");
      }
      result = { type: "radial-gradient", stops };
    }
  }
  const tileRect = gradFillCursor.child(A_NS, "tileRect");
  if (tileRect !== undefined) {
    tileRect.consumeAttr("l");
    tileRect.consumeAttr("t");
    tileRect.consumeAttr("r");
    tileRect.consumeAttr("b");
  }
  return result;
}

/** Reads DrawingML `<a:blipFill>` embedded image bytes and optional `<a:alphaModFix>` opacity. */
export function readBlipFill(
  blipFillCursor: XmlElementCursor,
  pkg: PptxPackage,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
  validateImageBytes?: (bytes: Uint8Array) => void,
): { readonly image: Uint8Array; readonly opacity?: number } | undefined {
  blipFillCursor.consumeAttr("rotWithShape");
  blipFillCursor.consumeAttr("dpi");
  const srcRect = blipFillCursor.child(A_NS, "srcRect");
  if (srcRect !== undefined) {
    srcRect.consumeAttr("l");
    srcRect.consumeAttr("t");
    srcRect.consumeAttr("r");
    srcRect.consumeAttr("b");
  }
  const stretch = blipFillCursor.child(A_NS, "stretch");
  if (stretch !== undefined) {
    stretch.ignoreEmptyChild(A_NS, "fillRect");
  }
  const blip = blipFillCursor.child(A_NS, "blip");
  if (blip === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: blipFillCursor.xpath },
    });
    return undefined;
  }
  blip.consumeAttr("cstate");
  const embedId = blip.attr("embed", R_NS);
  const linkId = blip.attr("link", R_NS);
  let opacity: number | undefined;
  const alphaModFix = blip.child(A_NS, "alphaModFix");
  if (alphaModFix !== undefined) {
    const parsedAmt = alphaModFixAttrSchema.safeParse(
      alphaModFix.attrsObject(["amt"]),
    );
    if (parsedAmt.success) {
      opacity = Math.min(1, Math.max(0, parsedAmt.data.amt / 100_000));
    }
  }
  if (embedId === undefined) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: blip.xpath },
      namespaceUri: A_NS,
      localName: "blip",
      mechanismHint: linkId !== undefined ? "linked-image-fill" : "blip-fill",
    });
    return undefined;
  }
  const rels = pkg.readRelationships(location.part);
  diagnostics.push(...rels.diagnostics);
  const rel = rels.entries.find((entry) => entry.Id === embedId);
  if (rel === undefined || rel.TargetMode === "External") {
    diagnostics.push({
      code: "broken-relationship",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: blip.xpath },
      relationshipId: embedId,
    });
    return undefined;
  }
  const targetPart = resolvePartPath(location.part, rel.Target);
  if (targetPart === undefined || !pkg.hasPart(targetPart)) {
    diagnostics.push({
      code: "missing-target-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: blip.xpath },
      relationshipId: embedId,
      targetPart: targetPart ?? rel.Target,
    });
    return undefined;
  }
  const detachedBytes = pkg.readDetachedBytes(targetPart);
  if (detachedBytes === undefined) return undefined;
  if (validateImageBytes !== undefined) {
    try {
      validateImageBytes(detachedBytes);
    } catch {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: targetPart },
      });
      return undefined;
    }
  }
  return {
    image: detachedBytes,
    ...(opacity === undefined ? {} : { opacity }),
  };
}

export type ImportedLineBorder = {
  readonly noFill?: boolean;
  readonly borderColor?: ColorValue;
  readonly borderWidth?: number;
  readonly borderStyle?: "solid" | "dashed" | "dotted";
  readonly startArrow?: ArrowType;
  readonly endArrow?: ArrowType;
};

/** Reads DrawingML `<a:ln>` stroke width, fill, dash preset, join, and line end arrows. */
export function readLineBorder(
  lnCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ImportedLineBorder {
  const rawLn = lnCursor.attrsObject(["w", "cap", "cmpd", "algn"]);
  const parsedLn = lineAttrSchema.safeParse(rawLn);
  const width =
    parsedLn.success && parsedLn.data.w !== undefined
      ? emuToPx(parsedLn.data.w)
      : undefined;

  let borderColor: ColorValue | undefined;
  let borderStyle: "solid" | "dashed" | "dotted" | undefined;
  let startArrow: ArrowType | undefined;
  let endArrow: ArrowType | undefined;
  let hasNoFill = false;

  for (const child of lnCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    switch (child.localName) {
      case "noFill":
        hasNoFill = true;
        break;
      case "solidFill":
        borderColor = readDrawingColor(child, location, diagnostics);
        break;
      case "prstDash": {
        const rawDash = child.attrsObject(["val"]);
        const parsedDash = dashPresetAttrSchema.safeParse(rawDash);
        if (parsedDash.success) {
          const val = parsedDash.data.val;
          if (val === "solid") borderStyle = "solid";
          else if (val === "dot" || val === "sysDot") borderStyle = "dotted";
          else if (
            val === "dash" ||
            val === "lgDash" ||
            val === "dashDot" ||
            val === "lgDashDot" ||
            val === "lgDashDotDot" ||
            val === "sysDash" ||
            val === "sysDashDot" ||
            val === "sysDashDotDot"
          ) {
            borderStyle = "dashed";
          } else {
            diagnostics.push({
              code: "unsupported-mechanism",
              category: "unsupported",
              disposition: "omitted",
              location: { ...location, xpath: child.xpath },
              namespaceUri: A_NS,
              localName: "prstDash",
              mechanismHint: `dash:${val}`,
            });
          }
        }
        break;
      }
      case "headEnd":
      case "tailEnd": {
        const rawEnd = child.attrsObject(["type", "w", "len"]);
        const parsedEnd = lineEndAttrSchema.safeParse(rawEnd);
        if (!parsedEnd.success) {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "omitted",
            location: { ...location, xpath: child.xpath },
            attributeOrElement: `a:${child.localName}`,
            rawValue: JSON.stringify(rawEnd),
          });
          break;
        }
        if (parsedEnd.data.type !== undefined) {
          if (child.localName === "headEnd") startArrow = parsedEnd.data.type;
          else endArrow = parsedEnd.data.type;
        }
        break;
      }
      case "round":
      case "miter":
      case "bevel":
        if (child.localName === "miter") child.consumeAttr("lim");
        break;
      default:
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          namespaceUri: child.namespaceUri,
          localName: child.localName,
          mechanismHint: `line:${child.localName}`,
        });
        break;
    }
  }

  if (hasNoFill) {
    return { noFill: true };
  }
  if (borderColor === undefined) {
    return {};
  }

  return {
    borderColor,
    borderWidth: width ?? 1,
    ...(borderStyle === undefined ? {} : { borderStyle }),
    ...(startArrow === undefined ? {} : { startArrow }),
    ...(endArrow === undefined ? {} : { endArrow }),
  };
}

/** Reads DrawingML `<a:effectLst>` outer shadow and soft edge while reporting unsupported effects cleanly. */
export function readEffectList(
  effectLstCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): {
  readonly boxShadow?: {
    readonly color?: ColorValue;
    readonly blur?: number;
    readonly distance?: number;
    readonly angle?: number;
  };
  readonly softEdge?: number;
} {
  let boxShadow:
    | {
        readonly color?: ColorValue;
        readonly blur?: number;
        readonly distance?: number;
        readonly angle?: number;
      }
    | undefined;
  let softEdge: number | undefined;

  for (const child of effectLstCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    if (child.localName === "outerShdw") {
      const raw = child.attrsObject([
        "blurRad",
        "dist",
        "dir",
        "algn",
        "rotWithShape",
      ]);
      const parsed = outerShadowAttrSchema.safeParse(raw);
      const color = readDrawingColor(child, location, diagnostics);
      if (parsed.success) {
        boxShadow = {
          ...(color === undefined ? {} : { color }),
          blur:
            parsed.data.blurRad === undefined
              ? 0
              : emuToPx(parsed.data.blurRad),
          distance:
            parsed.data.dist === undefined ? 0 : emuToPx(parsed.data.dist),
          angle: parsed.data.dir === undefined ? 0 : parsed.data.dir / 60_000,
        };
      }
    } else if (child.localName === "softEdge") {
      const raw = child.attrsObject(["rad"]);
      const parsed = softEdgeAttrSchema.safeParse(raw);
      if (parsed.success && parsed.data.rad > 0) {
        softEdge = emuToPx(parsed.data.rad);
      }
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `effect:${child.localName}`,
      });
    }
  }

  return {
    ...(boxShadow === undefined ? {} : { boxShadow }),
    ...(softEdge === undefined ? {} : { softEdge }),
  };
}

export type PlaceholderSpPrVisualOptions = {
  readonly readAppearance?: (
    spPr: XmlElementCursor,
    pkg: PptxPackage,
    fallbackUnits: { readonly w: number; readonly h: number },
    location: ImportLocation,
    diagnostics: ImportDiagnostic[],
  ) => {
    readonly geometry?: Geometry;
    readonly style: Record<string, unknown>;
  };
  readonly allowedStyleKeys?: ReadonlySet<string>;
};

function readBasicPlaceholderAppearance(
  spPr: XmlElementCursor,
  pkg: PptxPackage,
  _fallbackUnits: { readonly w: number; readonly h: number },
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): {
  readonly geometry?: Geometry;
  readonly style: Record<string, unknown>;
} {
  let prstGeomName: string | undefined;
  let roundRectAdj: number | undefined;
  let hasNoFill = false;
  let hasEmptyEffectLst = false;
  let backgroundColor: ColorValue | undefined;
  let backgroundImage: unknown;
  let opacity: number | undefined;
  let border: ImportedLineBorder = {};
  let effects: ReturnType<typeof readEffectList> = {};

  for (const child of spPr.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) continue;
    switch (child.localName) {
      case "prstGeom": {
        prstGeomName = child.attr("prst");
        const avLst = child.child(A_NS, "avLst");
        if (avLst !== undefined) {
          for (const gd of avLst.children(A_NS, "gd")) {
            const name = gd.attr("name");
            const fmla = gd.attr("fmla");
            if (
              name === "adj" &&
              fmla !== undefined &&
              fmla.startsWith("val ")
            ) {
              const num = Number(fmla.slice(4));
              if (Number.isSafeInteger(num)) roundRectAdj = num;
            }
          }
        }
        break;
      }
      case "noFill":
        hasNoFill = true;
        break;
      case "solidFill":
        backgroundColor = readDrawingColor(child, location, diagnostics);
        break;
      case "gradFill":
        backgroundImage = readGradientFill(child, location, diagnostics);
        break;
      case "blipFill": {
        const blip = readBlipFill(child, pkg, location, diagnostics);
        if (blip !== undefined) {
          backgroundImage = { image: blip.image };
          opacity = blip.opacity;
        }
        break;
      }
      case "ln":
        border = readLineBorder(child, location, diagnostics);
        break;
      case "effectLst":
        hasEmptyEffectLst = child.element.children.length === 0;
        effects = readEffectList(child, location, diagnostics);
        break;
      case "extLst":
        child.ignoreEmptyChild(A_NS, "ext");
        break;
      default:
        break;
    }
  }

  const parsedPrst =
    prstGeomName === undefined
      ? undefined
      : presetGeometryNameSchema.safeParse(prstGeomName);
  const geometry: Geometry | undefined = parsedPrst?.success
    ? roundRectAdj !== undefined
      ? { type: "preset", name: parsedPrst.data, adjust: { adj: roundRectAdj } }
      : parsedPrst.data
    : undefined;

  const background =
    backgroundColor ?? backgroundImage ?? (hasNoFill ? "none" : undefined);
  const resolvedBorder =
    border.borderColor !== undefined
      ? {
          color: border.borderColor,
          ...(border.borderWidth !== undefined
            ? { width: border.borderWidth }
            : {}),
          ...(border.borderStyle !== undefined
            ? { style: border.borderStyle }
            : {}),
        }
      : border.noFill === true
        ? "none"
        : undefined;
  const resolvedBoxShadow =
    effects.boxShadow ??
    (hasEmptyEffectLst && effects.softEdge === undefined ? "none" : undefined);
  const style: Record<string, unknown> = {
    ...(background === undefined ? {} : { background }),
    ...(resolvedBorder === undefined ? {} : { border: resolvedBorder }),
    ...(opacity === undefined ? {} : { opacity }),
    ...(resolvedBoxShadow === undefined
      ? {}
      : { boxShadow: resolvedBoxShadow }),
    ...(effects.softEdge === undefined ? {} : { softEdge: effects.softEdge }),
  };

  return {
    ...(geometry === undefined ? {} : { geometry }),
    style,
  };
}

/**
 * Shared reader for master and layout placeholder `<p:spPr>` transforms, geometry, fills, borders, and effects.
 */
export function readPlaceholderSpPrVisuals(
  args: {
    readonly spPr: XmlElementCursor;
    readonly pkg: PptxPackage;
    readonly location: ImportLocation;
    readonly diagnostics: ImportDiagnostic[];
  },
  options?: PlaceholderSpPrVisualOptions,
): ImportedPlaceholderSpPr {
  const { spPr, pkg, location, diagnostics } = args;
  let x: number | undefined;
  let y: number | undefined;
  let width: number | undefined;
  let height: number | undefined;
  let rotate: number | undefined;
  let flipH: boolean | undefined;
  let flipV: boolean | undefined;
  let fallbackUnits = { w: 0, h: 0 };

  const xfrm = spPr.child(A_NS, "xfrm");
  if (xfrm !== undefined) {
    const hasExplicitFlipH = xfrm.element.hasAttribute("flipH");
    const hasExplicitFlipV = xfrm.element.hasAttribute("flipV");
    if (
      xfrm.child(A_NS, "off") !== undefined &&
      xfrm.child(A_NS, "ext") !== undefined
    ) {
      const tf = readTransformXml(xfrm, location, diagnostics);
      if (tf !== undefined) {
        x = tf.x;
        y = tf.y;
        width = tf.width;
        height = tf.height;
        rotate = tf.rotate;
        if (hasExplicitFlipH) flipH = Boolean(tf.flipH);
        if (hasExplicitFlipV) flipV = Boolean(tf.flipV);
        fallbackUnits = { w: tf.cxEmu, h: tf.cyEmu };
      }
    } else {
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
    }
  }

  const appearance = (
    options?.readAppearance ?? readBasicPlaceholderAppearance
  )(spPr, pkg, fallbackUnits, location, diagnostics);

  let borderRadius: number | undefined;
  let exportedGeometry: Geometry | undefined;
  const geom = appearance.geometry;
  if (geom === "roundRect") {
    if (width !== undefined && height !== undefined) {
      borderRadius = (16667 / 100_000) * Math.min(width, height);
    } else {
      exportedGeometry = geom;
    }
  } else if (typeof geom === "object" && geom.type === "preset") {
    if (geom.name === "roundRect") {
      if (width !== undefined && height !== undefined) {
        const adj = geom.adjust?.["adj"] ?? 16667;
        if (adj > 0) {
          borderRadius = (adj / 100_000) * Math.min(width, height);
        }
      } else {
        exportedGeometry = geom;
      }
    } else if (geom.name !== "rect") {
      exportedGeometry = geom;
    }
  } else if (geom !== undefined && geom !== "rect") {
    exportedGeometry = geom;
  }

  const style: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(appearance.style)) {
    if (
      value !== undefined &&
      (options?.allowedStyleKeys === undefined ||
        options.allowedStyleKeys.has(key))
    ) {
      style[key] = value;
    }
  }
  if (x !== undefined && y !== undefined) {
    style["position"] = "absolute";
    style["left"] = x;
    style["top"] = y;
  }
  if (width !== undefined) {
    style["width"] = width;
  }
  if (height !== undefined) {
    style["height"] = height;
  }
  if (borderRadius !== undefined) {
    style["borderRadius"] = borderRadius;
  }
  if (rotate !== undefined) {
    style["rotate"] = rotate;
  }
  if (flipH !== undefined) {
    style["flipH"] = flipH;
  }
  if (flipV !== undefined) {
    style["flipV"] = flipV;
  }

  return {
    ...(exportedGeometry === undefined ? {} : { geometry: exportedGeometry }),
    ...(Object.keys(style).length === 0 ? {} : { style }),
  };
}
