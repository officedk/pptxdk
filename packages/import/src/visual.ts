import type {
  ColorValue,
  GradientStop,
  LinearGradient,
  RadialGradient,
  SchemeColorName,
} from "@pptxdk/core/capability";
import { schemeColorNameSchema } from "@pptxdk/core/capability";
import { z } from "zod";

import type { ImportDiagnostic, ImportLocation } from "#/diagnostics.ts";
import { A_NS, R_NS, emuToPx } from "#/namespaces.ts";
import { resolvePartPath, type PptxPackage } from "#/opc.ts";
import {
  ExtentEmuFragmentSchema,
  HexRgb6Schema,
  OffsetEmuFragmentSchema,
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  TransformAttributesFragmentSchema,
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

export const arrowTypeSchema = z.enum(arrowTypes);
export type ArrowType = z.infer<typeof arrowTypeSchema>;

const PercentValAttrSchema = z.object({
  val: OoxmlNonNegativeIntSchema,
});

const AlphaModFixAttrSchema = z.object({
  amt: OoxmlNonNegativeIntSchema,
});

const SrgbColorAttrSchema = z.object({
  val: HexRgb6Schema,
});

const SchemeColorAttrSchema = z.object({
  val: schemeColorNameSchema,
});

const SysColorAttrSchema = z.object({
  val: z.string().min(1),
  lastClr: HexRgb6Schema.optional(),
});

const GradStopAttrSchema = z.object({
  pos: OoxmlNonNegativeIntSchema,
});

const GradLinAttrSchema = z.object({
  ang: OoxmlIntSchema.optional(),
  scaled: z.enum(["0", "1", "true", "false"]).optional(),
});

const LineAttrSchema = z.object({
  w: OoxmlNonNegativeIntSchema.optional(),
  cap: z.enum(["rnd", "sq", "flat"]).optional(),
  cmpd: z.enum(["sng", "dbl", "thickThin", "thinThick", "tri"]).optional(),
  algn: z.enum(["ctr", "in"]).optional(),
});

const DashPresetAttrSchema = z.object({
  val: z.string().min(1),
});

const LineEndAttrSchema = z.object({
  type: arrowTypeSchema.optional(),
  w: z.enum(["sm", "med", "lg"]).optional(),
  len: z.enum(["sm", "med", "lg"]).optional(),
});

const OuterShadowAttrSchema = z.object({
  blurRad: OoxmlNonNegativeIntSchema.optional(),
  dist: OoxmlNonNegativeIntSchema.optional(),
  dir: OoxmlIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: z.enum(["0", "1", "true", "false"]).optional(),
});

const SoftEdgeAttrSchema = z.object({
  rad: OoxmlNonNegativeIntSchema,
});

/** Reads a DrawingML color child (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`) with optional alpha and luminance transforms. */
export function readDrawingColor(
  parentCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ColorValue | undefined {
  let colorCursor: XmlElementCursor | undefined;
  let baseHex: string | undefined;
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
      const parsed = SrgbColorAttrSchema.safeParse(raw);
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
      const parsed = SchemeColorAttrSchema.safeParse(raw);
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
      const parsed = SysColorAttrSchema.safeParse(raw);
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
      const parsed = PercentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        alpha = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
    } else if (mod.namespaceUri === A_NS && mod.localName === "lumMod") {
      const parsed = PercentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        lumMod = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
    } else if (mod.namespaceUri === A_NS && mod.localName === "lumOff") {
      const parsed = PercentValAttrSchema.safeParse(mod.attrsObject(["val"]));
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
  const parsedXfrm = TransformAttributesFragmentSchema.safeParse(rawXfrm);
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
      const parsedPos = GradStopAttrSchema.safeParse(rawPos);
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
    const parsedLin = GradLinAttrSchema.safeParse(rawLin);
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
    const parsedAmt = AlphaModFixAttrSchema.safeParse(
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
  const parsedLn = LineAttrSchema.safeParse(rawLn);
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
        const parsedDash = DashPresetAttrSchema.safeParse(rawDash);
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
        const parsedEnd = LineEndAttrSchema.safeParse(rawEnd);
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

  if (hasNoFill || borderColor === undefined) {
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
      const parsed = OuterShadowAttrSchema.safeParse(raw);
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
      const parsed = SoftEdgeAttrSchema.safeParse(raw);
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
