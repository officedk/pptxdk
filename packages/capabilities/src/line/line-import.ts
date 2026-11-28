import type { ColorValue } from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";
import { pxToEmu } from "@pptxdk/emit";
import {
  A_NS,
  P_NS,
  allocateScopedAuthorId,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
  readDrawingColor,
  readTransformXml,
  type ImportDiagnostic,
  type ImportLocation,
  type ScopeConnectorResolutionContext,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  readNonVisualCommonProperties,
  type BoxCoordinateMapper,
} from "../shape/shape-import.ts";
import {
  getGeometryConnectionSiteCount,
  type Geometry,
} from "../shared/geometry.ts";
import {
  attachDeferredSlideLink,
  consumeNonPlaceholderNvPr,
} from "../shared/identity.ts";
import { isPlaceholderSpElement } from "../shared/placeholder-import.ts";
import {
  readSpPrAppearance,
  type ImportedTransform,
} from "../shared/visual.ts";
import {
  createLineNode,
  type LineAdjust,
  type LineDeclarationProps,
  type LineEndpoint,
  type LinePort,
  type LineRouting,
  type LineStrokeValue,
  type LineStyle,
} from "./line-declaration.ts";

export type LineBoxCoordinateMapper = BoxCoordinateMapper;

type XmlElement = XmlElementCursor["element"];

const linePresetNames = new Set<string>([
  "line",
  "lineInv",
  "straightConnector1",
  "bentConnector2",
  "bentConnector3",
  "bentConnector4",
  "bentConnector5",
  "curvedConnector2",
  "curvedConnector3",
  "curvedConnector4",
  "curvedConnector5",
]);

const fillElementNames = new Set<string>([
  "solidFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
]);

const connectionAttrSchema = z.object({
  id: ooxmlPositiveIntSchema,
  idx: ooxmlNonNegativeIntSchema,
});

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

function hasTextBoxFlag(spElement: XmlElement): boolean {
  const nvSpPr = directElementChild(spElement, P_NS, "nvSpPr");
  if (nvSpPr === undefined) return false;
  const cNvSpPr = directElementChild(nvSpPr, P_NS, "cNvSpPr");
  if (cNvSpPr === undefined) return false;
  const txBox = cNvSpPr.getAttribute("txBox");
  return txBox === "1" || txBox === "true";
}

function hasEffectiveSpText(spElement: XmlElement): boolean {
  const txBody = directElementChild(spElement, P_NS, "txBody");
  if (txBody === undefined) return false;
  for (
    let child = txBody.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const pEl = child as XmlElement;
    if (
      (pEl.namespaceURI ?? "") !== A_NS ||
      (pEl.localName ?? pEl.tagName) !== "p"
    ) {
      continue;
    }
    const pPr = directElementChild(pEl, A_NS, "pPr");
    if (pPr !== undefined) {
      if (
        directElementChild(pPr, A_NS, "buChar") !== undefined ||
        directElementChild(pPr, A_NS, "buAutoNum") !== undefined ||
        directElementChild(pPr, A_NS, "defRPr") !== undefined
      ) {
        return true;
      }
    }
    if (
      directElementChild(pEl, A_NS, "r") !== undefined ||
      directElementChild(pEl, A_NS, "fld") !== undefined ||
      directElementChild(pEl, A_NS, "br") !== undefined
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Side-effect-free predicate matching `<p:sp>` elements that represent pure 1D lines or connectors
 * (no fill, no effective text, and a `line` / `lineInv` / `*Connector*` preset geometry).
 */
export function isLineNormalizedSpElement(cursor: XmlElementCursor): boolean {
  if (isPlaceholderSpElement(cursor)) return false;
  const spEl = cursor.element;
  if (hasTextBoxFlag(spEl)) return false;

  const spPr = directElementChild(spEl, P_NS, "spPr");
  if (spPr === undefined) return false;
  if (directElementChild(spPr, A_NS, "custGeom") !== undefined) return false;

  const prstGeom = directElementChild(spPr, A_NS, "prstGeom");
  if (prstGeom === undefined) return false;
  const prst = prstGeom.getAttribute("prst") ?? "";
  if (!linePresetNames.has(prst)) return false;

  for (let child = spPr.firstChild; child !== null; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if (
      (el.namespaceURI ?? "") === A_NS &&
      fillElementNames.has(el.localName ?? el.tagName)
    ) {
      return false;
    }
  }

  const pStyle = directElementChild(spEl, P_NS, "style");
  if (pStyle !== undefined) {
    const fillRef = directElementChild(pStyle, A_NS, "fillRef");
    if (fillRef !== undefined) {
      const idx = fillRef.getAttribute("idx");
      if (idx !== null && idx !== "" && idx !== "0") {
        return false;
      }
    }
  }

  if (hasEffectiveSpText(spEl)) return false;
  return true;
}

function readStyleLnRefColor(
  parentCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ColorValue | undefined {
  const styleCursor = parentCursor.child(P_NS, "style");
  if (styleCursor === undefined) return undefined;
  const lnRef = styleCursor.child(A_NS, "lnRef");
  let lnColor: ColorValue | undefined;
  if (lnRef !== undefined) {
    const idxStr = lnRef.attr("idx") ?? "0";
    const parsedColor = readDrawingColor(lnRef, location, diagnostics);
    if (idxStr !== "0") {
      lnColor = parsedColor;
    }
  }
  styleCursor.suppressUnconsumed();
  return lnColor;
}

function mapLineGeometry(
  geometry: ReturnType<typeof readSpPrAppearance>["geometry"],
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): {
  readonly routing: LineRouting;
  readonly adjust?: LineAdjust;
  readonly invertFlipV: boolean;
} {
  if (geometry === undefined) {
    return { routing: "straight", invertFlipV: false };
  }
  if (typeof geometry !== "string" && geometry.type === "custom") {
    diagnostics.push({
      code: "compatibility-fallback",
      category: "unsupported",
      disposition: "fallback",
      location,
      preferredMechanism: "custom-geometry-connector",
      selectedFallback: "line",
      loss: { lost: "native-mechanism", retainedAs: "line" },
    });
    return { routing: "straight", invertFlipV: false };
  }
  const name = typeof geometry === "string" ? geometry : geometry.name;
  const rawAdjust = typeof geometry === "string" ? undefined : geometry.adjust;

  if (name === "line" || name === "straightConnector1") {
    return { routing: "straight", invertFlipV: false };
  }
  if (name === "lineInv") {
    return { routing: "straight", invertFlipV: true };
  }
  if (name.startsWith("bentConnector") || name.startsWith("curvedConnector")) {
    const routing: LineRouting = name.startsWith("bentConnector")
      ? "elbow"
      : "curved";
    if (name.endsWith("2")) {
      return { routing, adjust: {}, invertFlipV: false };
    }
    const adjust: LineAdjust = {
      ...(rawAdjust?.["adj1"] === undefined ? {} : { adj1: rawAdjust["adj1"] }),
      ...(rawAdjust?.["adj2"] === undefined || name.endsWith("3")
        ? {}
        : { adj2: rawAdjust["adj2"] }),
      ...(rawAdjust?.["adj3"] === undefined || !name.endsWith("5")
        ? {}
        : { adj3: rawAdjust["adj3"] }),
    };
    return {
      routing,
      ...(Object.keys(adjust).length === 0 ? {} : { adjust }),
      invertFlipV: false,
    };
  }

  diagnostics.push({
    code: "compatibility-fallback",
    category: "unsupported",
    disposition: "fallback",
    location,
    preferredMechanism: `preset:${name}`,
    selectedFallback: "line",
    loss: { lost: "native-mechanism", retainedAs: "line" },
  });
  return { routing: "straight", invertFlipV: false };
}

function readEndpointRef(
  cNvCxnSpPr: XmlElementCursor,
  tagName: "stCxn" | "endCxn",
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): { readonly id: number; readonly idx: number } | undefined {
  const el = cNvCxnSpPr.child(A_NS, tagName);
  if (el === undefined) return undefined;
  const raw = el.attrsObject(["id", "idx"]);
  const parsed = connectionAttrSchema.safeParse(raw);
  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: el.xpath },
      attributeOrElement: `a:${tagName}`,
      rawValue: JSON.stringify(raw),
    });
    return undefined;
  }
  return parsed.data;
}

function indexToPort(idx: number): LinePort {
  switch (idx) {
    case 0:
      return "top";
    case 1:
      return "left";
    case 2:
      return "bottom";
    case 3:
      return "right";
    default:
      return idx;
  }
}

function buildLineNodeFromSpPr(args: {
  readonly context: ShapeTreeImportContext;
  readonly spPr: XmlElementCursor;
  readonly lnRefColor: ColorValue | undefined;
  readonly objectId: number | undefined;
  readonly objectName: string | undefined;
  readonly link: LineDeclarationProps["link"] | undefined;
  readonly mapBox?: LineBoxCoordinateMapper;
  readonly diagnostics: ImportDiagnostic[];
}):
  | {
      readonly node: CommonNode;
      readonly mappedBox: {
        readonly x: number;
        readonly y: number;
        readonly width: number;
        readonly height: number;
      };
      readonly flipH: boolean;
      readonly flipV: boolean;
      readonly rotate: number | undefined;
    }
  | undefined {
  const {
    context,
    spPr,
    lnRefColor,
    objectId,
    objectName,
    link,
    mapBox,
    diagnostics,
  } = args;
  const location: ImportLocation = {
    part: context.part,
    slideIndex: context.slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };

  let transform: ImportedTransform | undefined;
  const xfrmCursor = spPr.child(A_NS, "xfrm");
  if (xfrmCursor !== undefined) {
    transform = readTransformXml(xfrmCursor, location, diagnostics, {
      allowSingleZeroExtent: true,
    });
  } else {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: spPr.xpath },
    });
  }
  if (transform === undefined) return undefined;

  const appearance = readSpPrAppearance(
    spPr,
    context.package,
    { w: transform.cxEmu, h: transform.cyEmu },
    location,
    diagnostics,
  );

  const mappedBox = mapBox
    ? mapBox({
        x: transform.x,
        y: transform.y,
        width: transform.width,
        height: transform.height,
      })
    : {
        x: transform.x,
        y: transform.y,
        width: transform.width,
        height: transform.height,
      };

  const { routing, adjust, invertFlipV } = mapLineGeometry(
    appearance.geometry,
    { ...location, xpath: spPr.xpath },
    diagnostics,
  );
  const effectiveFlipH = transform.flipH === true;
  const effectiveFlipV = invertFlipV
    ? transform.flipV !== true
    : transform.flipV === true;

  const rawBorder = appearance.style.border;
  const borderArrows = appearance.borderArrows;
  let stroke: LineStrokeValue | undefined;
  if (rawBorder === "none") {
    stroke = "none";
  } else if (rawBorder !== undefined) {
    stroke = {
      color: rawBorder.color,
      ...(rawBorder.width === undefined ? {} : { width: rawBorder.width }),
      ...(rawBorder.style === undefined ? {} : { style: rawBorder.style }),
      ...(borderArrows?.startArrow === undefined
        ? {}
        : { startArrow: borderArrows.startArrow }),
      ...(borderArrows?.endArrow === undefined
        ? {}
        : { endArrow: borderArrows.endArrow }),
    };
  } else if (lnRefColor !== undefined) {
    stroke = { color: lnRefColor, width: 1, style: "solid" };
  } else {
    stroke = "none";
  }

  const lineStyle: LineStyle = {
    position: "absolute",
    left: mappedBox.x,
    top: mappedBox.y,
    width: mappedBox.width,
    height: mappedBox.height,
    ...(transform.rotate === undefined ? {} : { rotate: transform.rotate }),
    ...(effectiveFlipH ? { flipH: true } : {}),
    ...(effectiveFlipV ? { flipV: true } : {}),
    ...(appearance.style.boxShadow === undefined
      ? {}
      : { boxShadow: appearance.style.boxShadow }),
  };

  const props: LineDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    ...(routing === "straight" ? {} : { routing }),
    ...(adjust === undefined ? {} : { adjust }),
    stroke,
    style: lineStyle,
  };

  return {
    node: createLineNode(props),
    mappedBox,
    flipH: effectiveFlipH,
    flipV: effectiveFlipV,
    rotate: transform.rotate,
  };
}

/** Imports `<p:cxnSp>` connectors and restores `from`/`to` endpoints when targets resolve in scope. */
export function readLineCxnSp(
  context: ShapeTreeImportContext,
  mapBox?: LineBoxCoordinateMapper,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const nvCxnSpPr = cursor.child(P_NS, "nvCxnSpPr");
  if (nvCxnSpPr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part, xpath: cursor.xpath, slideIndex },
    });
    return { diagnostics };
  }

  const cNvPr = nvCxnSpPr.child(P_NS, "cNvPr");
  const { objectId, objectName, link, deferredSlideLink } =
    readNonVisualCommonProperties(cNvPr, context, diagnostics);

  const location: ImportLocation = {
    part,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };

  const cNvCxnSpPr = nvCxnSpPr.child(P_NS, "cNvCxnSpPr");
  const stCxn =
    cNvCxnSpPr === undefined
      ? undefined
      : readEndpointRef(cNvCxnSpPr, "stCxn", location, diagnostics);
  const endCxn =
    cNvCxnSpPr === undefined
      ? undefined
      : readEndpointRef(cNvCxnSpPr, "endCxn", location, diagnostics);
  cNvCxnSpPr?.suppressUnconsumed();
  consumeNonPlaceholderNvPr(nvCxnSpPr);

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

  const lnRefColor = readStyleLnRefColor(cursor, location, diagnostics);
  const built = buildLineNodeFromSpPr({
    context,
    spPr,
    lnRefColor,
    objectId,
    objectName,
    link,
    ...(mapBox === undefined ? {} : { mapBox }),
    diagnostics,
  });
  if (built === undefined) {
    return { diagnostics };
  }

  attachDeferredSlideLink(built.node, deferredSlideLink, context);

  if (
    (stCxn !== undefined || endCxn !== undefined) &&
    context.registerDeferredConnector !== undefined
  ) {
    const { node, mappedBox, flipH, flipV, rotate } = built;
    const startPt = {
      x: flipH ? mappedBox.x + mappedBox.width : mappedBox.x,
      y: flipV ? mappedBox.y + mappedBox.height : mappedBox.y,
    };
    const endPt = {
      x: flipH ? mappedBox.x : mappedBox.x + mappedBox.width,
      y: flipV ? mappedBox.y : mappedBox.y + mappedBox.height,
    };

    context.registerDeferredConnector(
      (scopeCtx: ScopeConnectorResolutionContext) => {
        const emitFallbackDiagnostic = (): void => {
          scopeCtx.diagnostics.push({
            code: "compatibility-fallback",
            category: "unsupported",
            disposition: "fallback",
            location: { ...location, xpath: cursor.xpath },
            preferredMechanism: "connector-anchor",
            selectedFallback: "line",
            loss: { lost: "native-mechanism", retainedAs: "line" },
          });
        };

        if (
          (rotate !== undefined && rotate !== 0) ||
          (pxToEmu(mappedBox.width) === 0 && pxToEmu(mappedBox.height) === 0)
        ) {
          emitFallbackDiagnostic();
          return;
        }

        const targetSiteCount = (target: CommonNode): number =>
          target.type === "shape"
            ? getGeometryConnectionSiteCount(
                target.props["geometry"] as Geometry | undefined,
              )
            : 4;

        const resolveImportedSide = (
          cxn: { readonly id: number; readonly idx: number } | undefined,
          fallbackPoint: { readonly x: number; readonly y: number },
        ): LineEndpoint | undefined => {
          if (cxn === undefined) {
            return { x: fallbackPoint.x, y: fallbackPoint.y };
          }
          const target = scopeCtx.scopeObjectsBySpid.get(cxn.id);
          if (target === undefined || target === node) return undefined;
          if (cxn.idx < 0 || cxn.idx >= targetSiteCount(target)) {
            return undefined;
          }
          let targetId =
            typeof target.props["id"] === "string" &&
            target.props["id"].trim().length > 0
              ? target.props["id"]
              : undefined;
          if (targetId === undefined) {
            const targetName =
              typeof target.props["name"] === "string"
                ? target.props["name"]
                : undefined;
            targetId = allocateScopedAuthorId(
              targetName,
              scopeCtx.usedAuthorIds,
              `obj-${cxn.id}`,
            );
            target.set({ id: targetId });
          }
          return {
            id: targetId,
            port: indexToPort(cxn.idx),
          };
        };

        const fromEndpoint = resolveImportedSide(stCxn, startPt);
        const toEndpoint = resolveImportedSide(endCxn, endPt);
        if (fromEndpoint === undefined || toEndpoint === undefined) {
          emitFallbackDiagnostic();
          return;
        }

        node.style({
          position: undefined,
          left: undefined,
          top: undefined,
          width: undefined,
          height: undefined,
          flipH: undefined,
          flipV: undefined,
          rotate: undefined,
        });
        node.set({
          from: fromEndpoint,
          to: toEndpoint,
        });
      },
    );
  }

  return { node: built.node, diagnostics };
}

/** Normalizes no-fill, no-text `line` / `lineInv` / `*Connector*` `<p:sp>` elements into `<Line>`. */
export function readLineSp(
  context: ShapeTreeImportContext,
  mapBox?: LineBoxCoordinateMapper,
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

  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  const { objectId, objectName, link, deferredSlideLink } =
    readNonVisualCommonProperties(cNvPr, context, diagnostics);
  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  cNvSpPr?.consumeAttr("txBox");
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

  const lnRefColor = readStyleLnRefColor(cursor, location, diagnostics);
  const txBody = cursor.child(P_NS, "txBody");
  txBody?.suppressUnconsumed();

  const built = buildLineNodeFromSpPr({
    context,
    spPr,
    lnRefColor,
    objectId,
    objectName,
    link,
    ...(mapBox === undefined ? {} : { mapBox }),
    diagnostics,
  });
  if (built === undefined) {
    return { diagnostics };
  }
  attachDeferredSlideLink(built.node, deferredSlideLink, context);
  return { node: built.node, diagnostics };
}

export const connectorLineImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "cxnSp",
  read: (context) => readLineCxnSp(context),
};

export const presetLineShapeImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "sp",
  matches: isLineNormalizedSpElement,
  read: (context) => readLineSp(context),
};
