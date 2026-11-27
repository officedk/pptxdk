import {
  matchLayoutPlaceholderKey,
  schemeColorNameSchema,
  type ColorValue,
  type NativePlaceholderRef,
  type Shadow,
} from "@pptxdk/core/capability";
import {
  A_NS,
  P_NS,
  REL_IMAGE,
  R_NS,
  emuToPx,
  extentEmuFragmentSchema,
  hexRgb6Schema,
  offsetEmuFragmentSchema,
  ooxmlNonNegativeIntSchema,
  transformAttributesFragmentSchema,
  resolveExternalRelationshipTarget,
  resolveInternalRelationshipTarget,
  type CascadedOmission,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  preset,
  presetGeometryNames,
  type CustomGeometry,
  type Geometry,
  type GeometryPath,
  type PathCommand,
  type PresetGeometryName,
} from "../shared/geometry.ts";
import { readClickLink, type ClickLinkReadResult } from "../shared/identity.ts";
import {
  isPlaceholderPicElement,
  readNvPrPlaceholder,
} from "../shared/placeholder-import.ts";
import {
  createImageNode,
  type ImageDeclarationProps,
  type ImageSource,
  type ImageVisualStyle,
} from "./image-declaration.ts";
import { readImage } from "./image-decode.ts";
import {
  alphaModFixFragmentSchema,
  customPathAttributesSchema,
  customPathPointSchema,
  geomGuideFragmentSchema,
  lineAttributesFragmentSchema,
  nonVisualPicIdSchema,
  outerShadowFragmentSchema,
  presetDashFragmentSchema,
  presetGeomFragmentSchema,
  softEdgeFragmentSchema,
  srcRectFragmentSchema,
} from "./image-schemas.ts";

const SchemeColorSchema = z.object({
  val: schemeColorNameSchema,
});

const SrgbColorSchema = z.object({
  val: hexRgb6Schema,
});

const AlphaModifierSchema = z.object({
  val: ooxmlNonNegativeIntSchema,
});

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
    baseColor = parsed.data.val;
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

function readCustomPathPoint(
  ptCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): { readonly x: number; readonly y: number } | undefined {
  const raw = ptCursor.attrsObject(["x", "y"]);
  const parsed = customPathPointSchema.safeParse(raw);
  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part, xpath: ptCursor.xpath, slideIndex },
      attributeOrElement: "a:pt",
      rawValue: JSON.stringify(raw),
    });
    return undefined;
  }
  return { x: parsed.data.x / 1000, y: parsed.data.y / 1000 };
}

function readCustomGeometry(
  custGeom: XmlElementCursor,
  part: string,
  slideIndex: number,
  fallbackSizePx: { readonly width: number; readonly height: number },
  diagnostics: ImportDiagnostic[],
): CustomGeometry | undefined {
  custGeom.ignoreEmptyChild(A_NS, "avLst");
  const gdLst = custGeom.child(A_NS, "gdLst");
  if (gdLst !== undefined && gdLst.remainingElementChildren().length > 0) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { part, xpath: gdLst.xpath, slideIndex },
      namespaceUri: A_NS,
      localName: "gdLst",
    });
    custGeom.suppressUnconsumed();
    return undefined;
  }
  custGeom.ignoreEmptyChild(A_NS, "ahLst");
  custGeom.ignoreEmptyChild(A_NS, "cxnLst");
  const rect = custGeom.child(A_NS, "rect");
  if (rect !== undefined) {
    rect.consumeAttr("l");
    rect.consumeAttr("t");
    rect.consumeAttr("r");
    rect.consumeAttr("b");
  }

  const pathLst = custGeom.child(A_NS, "pathLst");
  if (pathLst === undefined) return undefined;

  const paths: GeometryPath[] = [];
  let viewWidth: number | undefined;
  let viewHeight: number | undefined;

  for (const pathCursor of pathLst.children(A_NS, "path")) {
    const rawPath = pathCursor.attrsObject(["w", "h", "fill", "stroke"]);
    const parsedPath = customPathAttributesSchema.safeParse(rawPath);
    if (!parsedPath.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part, xpath: pathCursor.xpath, slideIndex },
        attributeOrElement: "a:path",
        rawValue: JSON.stringify(rawPath),
      });
      pathCursor.suppressUnconsumed();
      return undefined;
    }

    const pathW =
      parsedPath.data.w !== undefined
        ? parsedPath.data.w / 1000
        : fallbackSizePx.width;
    const pathH =
      parsedPath.data.h !== undefined
        ? parsedPath.data.h / 1000
        : fallbackSizePx.height;
    viewWidth ??= pathW;
    viewHeight ??= pathH;

    const commands: PathCommand[] = [];
    for (const cmdCursor of pathCursor.remainingElementChildren()) {
      if (cmdCursor.namespaceUri !== A_NS) {
        diagnostics.push({
          code: "unconsumed-xml",
          category: "unsupported",
          disposition: "omitted",
          kind: "element",
          location: { part, xpath: cmdCursor.xpath, slideIndex },
          namespaceUri: cmdCursor.namespaceUri,
          localName: cmdCursor.localName,
        });
        continue;
      }
      switch (cmdCursor.localName) {
        case "moveTo": {
          const pt = cmdCursor.child(A_NS, "pt");
          const coords =
            pt === undefined
              ? undefined
              : readCustomPathPoint(pt, part, slideIndex, diagnostics);
          if (coords === undefined) return undefined;
          commands.push({ op: "M", x: coords.x, y: coords.y });
          break;
        }
        case "lnTo": {
          const pt = cmdCursor.child(A_NS, "pt");
          const coords =
            pt === undefined
              ? undefined
              : readCustomPathPoint(pt, part, slideIndex, diagnostics);
          if (coords === undefined) return undefined;
          commands.push({ op: "L", x: coords.x, y: coords.y });
          break;
        }
        case "quadBezTo": {
          const pts = cmdCursor.children(A_NS, "pt");
          if (pts.length !== 2) return undefined;
          const p1 = readCustomPathPoint(
            pts[0]!,
            part,
            slideIndex,
            diagnostics,
          );
          const p2 = readCustomPathPoint(
            pts[1]!,
            part,
            slideIndex,
            diagnostics,
          );
          if (p1 === undefined || p2 === undefined) return undefined;
          commands.push({ op: "Q", x1: p1.x, y1: p1.y, x: p2.x, y: p2.y });
          break;
        }
        case "cubicBezTo": {
          const pts = cmdCursor.children(A_NS, "pt");
          if (pts.length !== 3) return undefined;
          const p1 = readCustomPathPoint(
            pts[0]!,
            part,
            slideIndex,
            diagnostics,
          );
          const p2 = readCustomPathPoint(
            pts[1]!,
            part,
            slideIndex,
            diagnostics,
          );
          const p3 = readCustomPathPoint(
            pts[2]!,
            part,
            slideIndex,
            diagnostics,
          );
          if (p1 === undefined || p2 === undefined || p3 === undefined) {
            return undefined;
          }
          commands.push({
            op: "C",
            x1: p1.x,
            y1: p1.y,
            x2: p2.x,
            y2: p2.y,
            x: p3.x,
            y: p3.y,
          });
          break;
        }
        case "close": {
          commands.push({ op: "Z" });
          break;
        }
        default: {
          diagnostics.push({
            code: "unsupported-mechanism",
            category: "unsupported",
            disposition: "omitted",
            location: { part, xpath: cmdCursor.xpath, slideIndex },
            namespaceUri: cmdCursor.namespaceUri,
            localName: cmdCursor.localName,
          });
          return undefined;
        }
      }
    }

    if (commands.length === 0) continue;
    const fill = parsedPath.data.fill === "none" ? false : true;
    const stroke =
      parsedPath.data.stroke === "0" || parsedPath.data.stroke === "false"
        ? false
        : true;
    paths.push({
      commands,
      ...(fill ? {} : { fill: false }),
      ...(stroke ? {} : { stroke: false }),
    });
  }

  if (
    paths.length === 0 ||
    viewWidth === undefined ||
    viewHeight === undefined
  ) {
    return undefined;
  }

  return {
    type: "custom",
    width: viewWidth,
    height: viewHeight,
    paths,
  };
}

function readPictureGeometry(
  spPr: XmlElementCursor,
  part: string,
  slideIndex: number,
  sizePx: { readonly width: number; readonly height: number },
  diagnostics: ImportDiagnostic[],
): Geometry | undefined {
  const prstGeom = spPr.child(A_NS, "prstGeom");
  if (prstGeom !== undefined) {
    const rawGeom = prstGeom.attrsObject(["prst"]);
    const parsedGeom = presetGeomFragmentSchema.safeParse(rawGeom);
    const avLst = prstGeom.child(A_NS, "avLst");
    const adjust: Record<string, number> = {};
    if (avLst !== undefined) {
      for (const gd of avLst.children(A_NS, "gd")) {
        const rawGd = gd.attrsObject(["name", "fmla"]);
        const parsedGd = geomGuideFragmentSchema.safeParse(rawGd);
        if (!parsedGd.success) {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "recovered",
            location: { part, xpath: gd.xpath, slideIndex },
            attributeOrElement: "a:gd",
            rawValue: JSON.stringify(rawGd),
          });
          continue;
        }
        const match = /^val (-?\d+)$/u.exec(parsedGd.data.fmla);
        if (match !== null) {
          adjust[parsedGd.data.name] = Number(match[1]);
        }
      }
    }
    if (!parsedGeom.success || !presetGeometryNames.has(parsedGeom.data.prst)) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: prstGeom.xpath, slideIndex },
        attributeOrElement: "a:prstGeom/@prst",
        rawValue: rawGeom["prst"] ?? "",
      });
      return undefined;
    }
    const prstName = parsedGeom.data.prst as PresetGeometryName;
    if (Object.keys(adjust).length > 0) {
      return preset(prstName, adjust);
    }
    return prstName === "rect" ? undefined : prstName;
  }

  const custGeom = spPr.child(A_NS, "custGeom");
  if (custGeom !== undefined) {
    return readCustomGeometry(custGeom, part, slideIndex, sizePx, diagnostics);
  }
  return undefined;
}

function readBorderStyle(
  ln: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): {
  readonly border?: ImageVisualStyle["border"];
} {
  if (ln.ignoreEmptyChild(A_NS, "noFill")) {
    return {};
  }
  const rawLn = ln.attrsObject(["w"]);
  const parsedLn = lineAttributesFragmentSchema.safeParse(rawLn);
  let borderWidth: number | undefined;
  if (!parsedLn.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { part, xpath: ln.xpath, slideIndex },
      attributeOrElement: "a:ln/@w",
      rawValue: rawLn["w"] ?? "",
    });
  } else if (parsedLn.data.w !== undefined) {
    borderWidth = emuToPx(parsedLn.data.w);
  }

  let borderColor: ColorValue | undefined;
  const solidFill = ln.child(A_NS, "solidFill");
  if (solidFill !== undefined) {
    borderColor = readColorElement(solidFill, part, slideIndex, diagnostics);
  }

  let borderStyle: "solid" | "dashed" | "dotted" | undefined;
  const prstDash = ln.child(A_NS, "prstDash");
  if (prstDash !== undefined) {
    const rawDash = prstDash.attrsObject(["val"]);
    const parsedDash = presetDashFragmentSchema.safeParse(rawDash);
    if (!parsedDash.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: prstDash.xpath, slideIndex },
        attributeOrElement: "a:prstDash/@val",
        rawValue: rawDash["val"] ?? "",
      });
    } else {
      switch (parsedDash.data.val) {
        case "solid":
          borderStyle = "solid";
          break;
        case "dash":
          borderStyle = "dashed";
          break;
        case "sysDot":
        case "dot":
          borderStyle = "dotted";
          break;
      }
    }
  }

  if (borderColor === undefined) {
    return {};
  }

  return {
    border: {
      color: borderColor,
      ...(borderWidth === undefined ? {} : { width: borderWidth }),
      ...(borderStyle === undefined ? {} : { style: borderStyle }),
    },
  };
}

function readEffectList(
  effectLst: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): {
  readonly boxShadow?: Shadow;
  readonly softEdge?: number;
} {
  let boxShadow: Shadow | undefined;
  let softEdge: number | undefined;

  const outerShdw = effectLst.child(A_NS, "outerShdw");
  if (outerShdw !== undefined) {
    const rawShadow = outerShdw.attrsObject([
      "blurRad",
      "dist",
      "dir",
      "algn",
      "rotWithShape",
    ]);
    const parsedShadow = outerShadowFragmentSchema.safeParse(rawShadow);
    if (!parsedShadow.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: outerShdw.xpath, slideIndex },
        attributeOrElement: "a:outerShdw",
        rawValue: JSON.stringify(rawShadow),
      });
    } else {
      const color = readColorElement(outerShdw, part, slideIndex, diagnostics);
      boxShadow = {
        ...(color === undefined ? {} : { color }),
        ...(parsedShadow.data.blurRad === undefined
          ? {}
          : { blur: emuToPx(parsedShadow.data.blurRad) }),
        ...(parsedShadow.data.dist === undefined
          ? {}
          : { distance: emuToPx(parsedShadow.data.dist) }),
        ...(parsedShadow.data.dir === undefined
          ? {}
          : { angle: parsedShadow.data.dir / 60_000 }),
      };
    }
  }

  const softEdgeCursor = effectLst.child(A_NS, "softEdge");
  if (softEdgeCursor !== undefined) {
    const rawSoft = softEdgeCursor.attrsObject(["rad"]);
    const parsedSoft = softEdgeFragmentSchema.safeParse(rawSoft);
    if (!parsedSoft.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part, xpath: softEdgeCursor.xpath, slideIndex },
        attributeOrElement: "a:softEdge/@rad",
        rawValue: rawSoft["rad"] ?? "",
      });
    } else if (parsedSoft.data.rad > 0) {
      softEdge = emuToPx(parsedSoft.data.rad);
    }
  }

  return {
    ...(boxShadow === undefined ? {} : { boxShadow }),
    ...(softEdge === undefined ? {} : { softEdge }),
  };
}

/**
 * Reads a Transitional `<p:pic>` element into an editable `ImageNode`, supporting embedded media,
 * external image links (without network access), crops, clipping geometries, visual effects, and click links.
 */
export function readImagePic(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, layoutPlaceholders } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const nvPicPr = cursor.child(P_NS, "nvPicPr");
  if (nvPicPr === undefined) {
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
  let alt: string | undefined;
  let clickResult: ClickLinkReadResult | undefined;

  const cNvPr = nvPicPr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    const rawCNvPr = cNvPr.attrsObject(["id", "name", "descr"]);
    const parsedCNvPr = nonVisualPicIdSchema.safeParse(rawCNvPr);
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
      if (
        parsedCNvPr.data.descr !== undefined &&
        parsedCNvPr.data.descr !== ""
      ) {
        alt = parsedCNvPr.data.descr;
      }
    }
    const hlinkClick = cNvPr.child(A_NS, "hlinkClick");
    if (hlinkClick !== undefined) {
      clickResult = readClickLink(hlinkClick, context, {
        ...(objectId === undefined ? {} : { objectId }),
        ...(objectName === undefined ? {} : { objectName }),
      });
    }
    cNvPr.ignoreEmptyChild(A_NS, "extLst");
  }

  const cNvPicPr = nvPicPr.child(P_NS, "cNvPicPr");
  if (cNvPicPr !== undefined) {
    const picLocks = cNvPicPr.child(A_NS, "picLocks");
    if (picLocks !== undefined) {
      picLocks.consumeAttr("noChangeAspect");
      picLocks.consumeAttr("noGrp");
      picLocks.consumeAttr("preferRelativeResize");
    }
  }
  const picLocation: ImportLocation = {
    part,
    xpath: cursor.xpath,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };
  const nvPr = nvPicPr.child(P_NS, "nvPr");
  const phResult = readNvPrPlaceholder(nvPr, picLocation, diagnostics);
  if (phResult.kind === "error") {
    return { diagnostics };
  }
  const nativeRef: NativePlaceholderRef | undefined =
    phResult.kind === "ok" ? phResult.nativeRef : undefined;

  const cascadedForPicture: CascadedOmission[] =
    clickResult?.clickLocation !== undefined
      ? [{ location: clickResult.clickLocation, role: "link" }]
      : [];

  const blipFill = cursor.child(P_NS, "blipFill");
  const spPr = cursor.child(P_NS, "spPr");
  if (blipFill === undefined || spPr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: picLocation,
      ...(cascadedForPicture.length === 0
        ? {}
        : { cascaded: cascadedForPicture }),
    });
    return { diagnostics };
  }

  blipFill.consumeAttr("rotWithShape");
  blipFill.consumeAttr("dpi");

  const blip = blipFill.child(A_NS, "blip");
  if (blip === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: picLocation,
      ...(cascadedForPicture.length === 0
        ? {}
        : { cascaded: cascadedForPicture }),
    });
    return { diagnostics };
  }

  const embedId = blip.attr("embed", R_NS);
  const linkId = blip.attr("link", R_NS);
  blip.consumeAttr("cstate");

  let opacity: number | undefined;
  const alphaModFix = blip.child(A_NS, "alphaModFix");
  if (alphaModFix !== undefined) {
    const rawAlpha = alphaModFix.attrsObject(["amt"]);
    const parsedAlpha = alphaModFixFragmentSchema.safeParse(rawAlpha);
    if (!parsedAlpha.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { ...picLocation, xpath: alphaModFix.xpath },
        attributeOrElement: "a:alphaModFix/@amt",
        rawValue: rawAlpha["amt"] ?? "",
      });
    } else if (parsedAlpha.data.amt < 100_000) {
      opacity = parsedAlpha.data.amt / 100_000;
    }
  }
  const blipExtLst = blip.child(A_NS, "extLst");
  if (blipExtLst !== undefined) {
    for (const ext of blipExtLst.children(A_NS, "ext")) {
      if (ext.attr("uri") === "{28A0092B-C50C-407E-A947-70E740481C1C}") {
        for (const child of ext.remainingElementChildren()) {
          child.consumeAttr("val");
        }
      }
    }
  }

  let src: ImageSource;
  const relationships = context.relationships ?? new Map();
  if (embedId !== undefined && embedId !== "") {
    const resolved = resolveInternalRelationshipTarget(
      context.package,
      relationships,
      part,
      embedId,
      picLocation,
      REL_IMAGE,
    );
    if (!resolved.ok) {
      diagnostics.push({
        ...resolved.diagnostic,
        ...(cascadedForPicture.length === 0
          ? {}
          : { cascaded: cascadedForPicture }),
      });
      return { diagnostics };
    }
    try {
      readImage(resolved.bytes);
    } catch {
      const cascaded: CascadedOmission[] = [
        { location: { part: resolved.targetPart }, role: "media" },
        ...cascadedForPicture,
      ];
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: picLocation,
        cascaded,
      });
      return { diagnostics };
    }
    src = resolved.bytes;
  } else if (linkId !== undefined && linkId !== "") {
    const resolved = resolveExternalRelationshipTarget(
      relationships,
      linkId,
      picLocation,
      REL_IMAGE,
    );
    if (!resolved.ok) {
      diagnostics.push({
        ...resolved.diagnostic,
        ...(cascadedForPicture.length === 0
          ? {}
          : { cascaded: cascadedForPicture }),
      });
      return { diagnostics };
    }
    src = { externalUrl: resolved.target };
  } else {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: picLocation,
      ...(cascadedForPicture.length === 0
        ? {}
        : { cascaded: cascadedForPicture }),
    });
    return { diagnostics };
  }

  let crop: ImageDeclarationProps["crop"] | undefined;
  const srcRect = blipFill.child(A_NS, "srcRect");
  if (srcRect !== undefined) {
    const rawCrop = srcRect.attrsObject(["l", "t", "r", "b"]);
    const parsedCrop = srcRectFragmentSchema.safeParse(rawCrop);
    if (!parsedCrop.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { ...picLocation, xpath: srcRect.xpath },
        attributeOrElement: "a:srcRect",
        rawValue: JSON.stringify(rawCrop),
      });
    } else {
      const { l, t, r, b } = parsedCrop.data;
      const nextCrop = {
        ...(l !== undefined && l !== 0 ? { left: l / 100_000 } : {}),
        ...(t !== undefined && t !== 0 ? { top: t / 100_000 } : {}),
        ...(r !== undefined && r !== 0 ? { right: r / 100_000 } : {}),
        ...(b !== undefined && b !== 0 ? { bottom: b / 100_000 } : {}),
      };
      if (Object.keys(nextCrop).length > 0) {
        crop = nextCrop;
      }
    }
  }

  const stretch = blipFill.child(A_NS, "stretch");
  if (stretch !== undefined) {
    stretch.ignoreEmptyChild(A_NS, "fillRect");
  }

  const matchedKey =
    nativeRef === undefined
      ? undefined
      : matchLayoutPlaceholderKey(layoutPlaceholders, {
          type: nativeRef.type,
          idx: nativeRef.idx,
        });
  const matchedLayoutPh =
    matchedKey === undefined ? undefined : layoutPlaceholders?.[matchedKey];

  let localX: number | undefined;
  let localY: number | undefined;
  let localWidth: number | undefined;
  let localHeight: number | undefined;
  let rotate: number | undefined;
  let flipH: true | undefined;
  let flipV: true | undefined;

  const xfrm = spPr.child(A_NS, "xfrm");
  if (xfrm === undefined) {
    if (nativeRef === undefined) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: picLocation,
        ...(cascadedForPicture.length === 0
          ? {}
          : { cascaded: cascadedForPicture }),
      });
      return { diagnostics };
    }
  } else {
    const off = xfrm.child(A_NS, "off");
    const ext = xfrm.child(A_NS, "ext");
    if (off === undefined || ext === undefined) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: picLocation,
        ...(cascadedForPicture.length === 0
          ? {}
          : { cascaded: cascadedForPicture }),
      });
      return { diagnostics };
    }

    const rawXfrm = xfrm.attrsObject(["rot", "flipH", "flipV"]);
    const parsedXfrm = transformAttributesFragmentSchema.safeParse(rawXfrm);
    const rawOff = off.attrsObject(["x", "y"]);
    const parsedOff = offsetEmuFragmentSchema.safeParse(rawOff);
    const rawExt = ext.attrsObject(["cx", "cy"]);
    const parsedExt = extentEmuFragmentSchema.safeParse(rawExt);
    if (!parsedOff.success || !parsedExt.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...picLocation, xpath: xfrm.xpath },
        attributeOrElement: "a:xfrm",
        ...(cascadedForPicture.length === 0
          ? {}
          : { cascaded: cascadedForPicture }),
      });
      return { diagnostics };
    }

    localX = emuToPx(parsedOff.data.x);
    localY = emuToPx(parsedOff.data.y);
    localWidth = emuToPx(parsedExt.data.cx);
    localHeight = emuToPx(parsedExt.data.cy);
    rotate =
      parsedXfrm.success &&
      parsedXfrm.data.rot !== undefined &&
      parsedXfrm.data.rot !== 0
        ? parsedXfrm.data.rot / 60_000
        : undefined;
    flipH =
      parsedXfrm.success && parsedXfrm.data.flipH === true ? true : undefined;
    flipV =
      parsedXfrm.success && parsedXfrm.data.flipV === true ? true : undefined;
  }

  const layoutPhWidth =
    typeof matchedLayoutPh?.style?.width === "number"
      ? matchedLayoutPh.style.width
      : undefined;
  const layoutPhHeight =
    typeof matchedLayoutPh?.style?.height === "number"
      ? matchedLayoutPh.style.height
      : undefined;

  const geometry = readPictureGeometry(
    spPr,
    part,
    slideIndex,
    {
      width: localWidth ?? layoutPhWidth ?? 100,
      height: localHeight ?? layoutPhHeight ?? 100,
    },
    diagnostics,
  );
  spPr.ignoreEmptyChild(A_NS, "noFill");

  let border: ReturnType<typeof readBorderStyle> = {};
  const ln = spPr.child(A_NS, "ln");
  if (ln !== undefined) {
    border = readBorderStyle(ln, part, slideIndex, diagnostics);
  }

  let effects: ReturnType<typeof readEffectList> = {};
  const effectLst = spPr.child(A_NS, "effectLst");
  if (effectLst !== undefined) {
    effects = readEffectList(effectLst, part, slideIndex, diagnostics);
  }

  if (clickResult !== undefined) {
    diagnostics.push(...clickResult.diagnostics);
  }

  const style: NonNullable<ImageDeclarationProps["style"]> = {
    ...(localX !== undefined && localY !== undefined
      ? { position: "absolute", left: localX, top: localY }
      : {}),
    ...(localWidth === undefined ? {} : { width: localWidth }),
    ...(localHeight === undefined ? {} : { height: localHeight }),
    ...(flipH === undefined ? {} : { flipH }),
    ...(flipV === undefined ? {} : { flipV }),
    ...(opacity === undefined ? {} : { opacity }),
    ...(rotate === undefined ? {} : { rotate }),
    ...border,
    ...effects,
  };

  const props: ImageDeclarationProps = {
    src,
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(matchedKey === undefined ? {} : { placeholder: matchedKey }),
    ...(nativeRef === undefined ? {} : { nativePlaceholder: nativeRef }),
    ...(alt === undefined ? {} : { alt }),
    ...(clickResult?.link === undefined ? {} : { link: clickResult.link }),
    ...(geometry === undefined ? {} : { geometry }),
    ...(crop === undefined ? {} : { crop }),
    ...(Object.keys(style).length === 0 ? {} : { style }),
  };

  const node = createImageNode(props);

  if (
    clickResult?.deferredSlideLink !== undefined &&
    context.registerDeferredSlideLink !== undefined
  ) {
    const pending = clickResult.deferredSlideLink;
    context.registerDeferredSlideLink({
      ...pending,
      apply(targetSlideId) {
        node.set({ link: { id: targetSlideId } });
      },
    });
  }

  return { node, diagnostics };
}

export const imageImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "pic",
  matches: (cursor) => !isPlaceholderPicElement(cursor),
  read: readImagePic,
};

export const picturePlaceholderImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "pic",
  matches: isPlaceholderPicElement,
  read: readImagePic,
};
