import { CommonNode } from "@pptxdk/core/document";
import { pxToEmu } from "@pptxdk/emit";
import {
  A_NS,
  DIAGRAM_NS,
  DSP_NS,
  MAXIMUM_COORDINATE_PX,
  MAXIMUM_WRITABLE_EMU,
  P_NS,
  REL_DIAGRAM_DATA,
  REL_DIAGRAM_DRAWING,
  R_NS,
  XMLNS_NS,
  XmlElementCursor,
  buildRelationshipIndex,
  emuToPx,
  extentEmuFragmentSchema,
  inspectUnsupportedGraphicFrame,
  offsetEmuFragmentSchema,
  parseXmlPart,
  peekShapeTreeObjectIdentity,
  readTransformXml,
  resolvePartPath,
  transformAttributesFragmentSchema,
  type ImportDiagnostic,
  type ImportLocation,
  type OmittedSlideObjectInfo,
  type ScopeConnectorResolutionContext,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
} from "@pptxdk/import";

import { imageImportHandler } from "../image/image-import.ts";
import {
  isLineNormalizedSpElement,
  readLineCxnSp,
  readLineSp,
} from "../line/line-import.ts";
import {
  readNonVisualCommonProperties,
  readShapeSp,
  type BoxCoordinateMapper,
} from "../shape/shape-import.ts";
import {
  attachDeferredSlideLink,
  consumeNonPlaceholderNvPr,
  readClickLink,
  type ClickLinkReadResult,
} from "../shared/identity.ts";
import { directElementChild } from "../shared/placeholder-import.ts";
import { textBoxImportHandler } from "../text/text-import.ts";
import {
  createGroupNode,
  type GroupDeclarationProps,
  type GroupStyle,
} from "./group-declaration.ts";

type PxBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

function readNodeStyleBox(node: CommonNode): PxBox | undefined {
  const style = (
    node.props as { readonly style?: Readonly<Record<string, unknown>> }
  ).style;
  if (
    style !== undefined &&
    typeof style["left"] === "number" &&
    typeof style["top"] === "number" &&
    typeof style["width"] === "number" &&
    typeof style["height"] === "number"
  ) {
    return {
      x: style["left"],
      y: style["top"],
      width: style["width"],
      height: style["height"],
    };
  }
  return undefined;
}

function nodeBoundsPx(node: CommonNode): PxBox | undefined {
  const styleBox = readNodeStyleBox(node);
  if (styleBox !== undefined) {
    return styleBox;
  }
  if (node.type === "group") {
    const childNodes = node.children.filter(
      (child): child is CommonNode => child instanceof CommonNode,
    );
    return memberUnionPx(childNodes);
  }
  return undefined;
}

function memberUnionPx(members: readonly CommonNode[]): PxBox | undefined {
  if (members.length === 0) return undefined;
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxRight = Number.NEGATIVE_INFINITY;
  let maxBottom = Number.NEGATIVE_INFINITY;
  for (const member of members) {
    const box = nodeBoundsPx(member);
    if (box === undefined) return undefined;
    minX = Math.min(minX, box.x);
    minY = Math.min(minY, box.y);
    maxRight = Math.max(maxRight, box.x + box.width);
    maxBottom = Math.max(maxBottom, box.y + box.height);
  }
  return {
    x: minX,
    y: minY,
    width: maxRight - minX,
    height: maxBottom - minY,
  };
}

const isValidMappedBox = (box: PxBox, allowZeroExtent = false): boolean => {
  if (
    !Number.isFinite(box.x) ||
    !Number.isFinite(box.y) ||
    Math.abs(box.x) > MAXIMUM_COORDINATE_PX ||
    Math.abs(box.y) > MAXIMUM_COORDINATE_PX ||
    !Number.isFinite(box.width) ||
    !Number.isFinite(box.height)
  ) {
    return false;
  }
  const cx = pxToEmu(box.width);
  const cy = pxToEmu(box.height);
  if (!Number.isSafeInteger(cx) || !Number.isSafeInteger(cy)) return false;
  if (cx > MAXIMUM_WRITABLE_EMU || cy > MAXIMUM_WRITABLE_EMU) return false;
  if (allowZeroExtent) {
    return cx >= 0 && cy >= 0 && (cx > 0 || cy > 0);
  }
  return cx >= 1 && cy >= 1;
};

const identityMapper: BoxCoordinateMapper = (box) => box;

function applyMemberMapper(
  node: CommonNode,
  memberMapper: BoxCoordinateMapper,
): void {
  const box = readNodeStyleBox(node);
  if (box !== undefined) {
    const mapped = memberMapper(box);
    node.style({
      position: "absolute",
      left: mapped.x,
      top: mapped.y,
      width: mapped.width,
      height: mapped.height,
    });
  }
}

/**
 * Converts slide-space coordinates computed in phase 1 into CSS containing-block-relative
 * coordinates. Only positioned groups establish a containing block for their descendants;
 * unpositioned groups inherit the nearest positioned ancestor's origin.
 */
function relativizeToContainingBlock(
  nodes: readonly CommonNode[],
  originX: number,
  originY: number,
): void {
  for (const node of nodes) {
    const style = (
      node.props as { readonly style?: Readonly<Record<string, unknown>> }
    ).style;
    const isPositioned =
      style !== undefined &&
      (style["position"] === "absolute" || style["position"] === "relative") &&
      typeof style["left"] === "number" &&
      typeof style["top"] === "number";

    if (node.type === "group") {
      const childNodes = node.children.filter(
        (child): child is CommonNode => child instanceof CommonNode,
      );
      if (isPositioned) {
        const groupOriginX = style["left"] as number;
        const groupOriginY = style["top"] as number;
        if (originX !== 0 || originY !== 0) {
          node.style({
            left: groupOriginX - originX,
            top: groupOriginY - originY,
          });
        }
        relativizeToContainingBlock(childNodes, groupOriginX, groupOriginY);
      } else {
        relativizeToContainingBlock(childNodes, originX, originY);
      }
      continue;
    }

    if (isPositioned && (originX !== 0 || originY !== 0)) {
      node.style({
        left: (style["left"] as number) - originX,
        top: (style["top"] as number) - originY,
      });
    }
  }
}

function readGroupSpSlideSpace(
  context: ShapeTreeImportContext,
  parentMapper: BoxCoordinateMapper,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex } = context;
  const diagnostics: ImportDiagnostic[] = [];

  const nvGrpSpPr = cursor.child(P_NS, "nvGrpSpPr");
  if (nvGrpSpPr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part, xpath: cursor.xpath, slideIndex },
    });
    return { diagnostics };
  }

  const cNvPr = nvGrpSpPr.child(P_NS, "cNvPr");
  const { objectId, objectName, link, deferredSlideLink } =
    readNonVisualCommonProperties(cNvPr, context, diagnostics);
  nvGrpSpPr.ignoreEmptyChild(P_NS, "cNvGrpSpPr");
  consumeNonPlaceholderNvPr(nvGrpSpPr);

  const location: ImportLocation = {
    part,
    slideIndex,
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };

  const grpSpPr = cursor.child(P_NS, "grpSpPr");
  if (grpSpPr === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: cursor.xpath },
    });
    return { diagnostics };
  }

  const xfrm = grpSpPr.child(A_NS, "xfrm");
  const off = xfrm?.child(A_NS, "off");
  const ext = xfrm?.child(A_NS, "ext");
  const chOff = xfrm?.child(A_NS, "chOff");
  const chExt = xfrm?.child(A_NS, "chExt");
  if (
    xfrm === undefined ||
    off === undefined ||
    ext === undefined ||
    chOff === undefined ||
    chExt === undefined
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: grpSpPr.xpath },
    });
    return { diagnostics };
  }

  const parsedXfrm = transformAttributesFragmentSchema.safeParse(
    xfrm.attrsObject(["rot", "flipH", "flipV"]),
  );
  const parsedOff = offsetEmuFragmentSchema.safeParse(
    off.attrsObject(["x", "y"]),
  );
  const parsedExt = extentEmuFragmentSchema.safeParse(
    ext.attrsObject(["cx", "cy"]),
  );
  const parsedChOff = offsetEmuFragmentSchema.safeParse(
    chOff.attrsObject(["x", "y"]),
  );
  const parsedChExt = extentEmuFragmentSchema.safeParse(
    chExt.attrsObject(["cx", "cy"]),
  );
  if (
    !parsedOff.success ||
    !parsedExt.success ||
    !parsedChOff.success ||
    !parsedChExt.success
  ) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: xfrm.xpath },
      attributeOrElement: "a:xfrm",
    });
    return { diagnostics };
  }

  const flipH = parsedXfrm.success && parsedXfrm.data.flipH === true;
  const flipV = parsedXfrm.success && parsedXfrm.data.flipV === true;

  const rotate =
    parsedXfrm.success &&
    parsedXfrm.data.rot !== undefined &&
    parsedXfrm.data.rot !== 0
      ? parsedXfrm.data.rot / 60_000
      : undefined;

  const offX = emuToPx(parsedOff.data.x);
  const offY = emuToPx(parsedOff.data.y);
  const extW = emuToPx(parsedExt.data.cx);
  const extH = emuToPx(parsedExt.data.cy);
  const chX = emuToPx(parsedChOff.data.x);
  const chY = emuToPx(parsedChOff.data.y);
  const chW = emuToPx(parsedChExt.data.cx);
  const chH = emuToPx(parsedChExt.data.cy);

  const scaleX = chW > 0 ? extW / chW : 1;
  const scaleY = chH > 0 ? extH / chH : 1;

  grpSpPr.ignoreEmptyChild(A_NS, "noFill");
  grpSpPr.ignoreEmptyChild(A_NS, "effectLst");
  for (const extraChild of grpSpPr.remainingElementChildren()) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: extraChild.xpath },
      namespaceUri: extraChild.namespaceUri,
      localName: extraChild.localName,
      mechanismHint: `group:${extraChild.localName}`,
    });
  }

  const memberMapper: BoxCoordinateMapper = (box) =>
    parentMapper({
      x: offX + (box.x - chX) * scaleX,
      y: offY + (box.y - chY) * scaleY,
      width: box.width * scaleX,
      height: box.height * scaleY,
    });

  const members: CommonNode[] = [];
  const groupObjectsBySpid = new Map<number, CommonNode>();
  const groupOmittedSpids = new Map<number, OmittedSlideObjectInfo>();
  const groupUsedAuthorIds = new Set<string>();
  const groupDeferredConnectors: Array<
    (ctx: ScopeConnectorResolutionContext) => void
  > = [];
  const recordGroupOmission = (
    spid: number | undefined,
    name: string | undefined,
  ): void => {
    if (spid !== undefined) {
      groupOmittedSpids.set(spid, {
        ...(name === undefined ? {} : { objectName: name }),
        reason: "corrupt",
      });
    }
  };
  const recordGroupMember = (
    spid: number | undefined,
    memberName: string | undefined,
    memberNode: CommonNode,
    memberXpath: string,
  ): void => {
    const memberBox = nodeBoundsPx(memberNode);
    if (
      memberBox !== undefined &&
      !isValidMappedBox(memberBox, memberNode.type === "line")
    ) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: memberXpath },
        attributeOrElement: "a:xfrm",
      });
      recordGroupOmission(spid, memberName);
      return;
    }
    if (spid !== undefined) {
      context.registerImportedObject?.(spid, memberNode);
      groupObjectsBySpid.set(spid, memberNode);
    }
    const authorId = memberNode.props["id"];
    if (typeof authorId === "string" && authorId.length > 0) {
      groupUsedAuthorIds.add(authorId);
    }
    members.push(memberNode);
  };

  for (const childCursor of cursor.remainingElementChildren()) {
    if (childCursor.namespaceUri !== P_NS) {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: childCursor.xpath },
        namespaceUri: childCursor.namespaceUri,
        localName: childCursor.localName,
      });
      continue;
    }

    const childContext: ShapeTreeImportContext = {
      ...context,
      cursor: childCursor,
      registerDeferredConnector: (resolve) => {
        groupDeferredConnectors.push(resolve);
      },
    };

    if (childCursor.localName === "grpSp") {
      const outcome = readGroupSpSlideSpace(childContext, memberMapper);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      const identity = peekShapeTreeObjectIdentity(childCursor.element);
      if (outcome.node !== undefined) {
        recordGroupMember(
          identity.objectId,
          identity.objectName,
          outcome.node,
          childCursor.xpath,
        );
      } else {
        recordGroupOmission(identity.objectId, identity.objectName);
      }
      continue;
    }

    if (childCursor.localName === "sp") {
      if (textBoxImportHandler.matches?.(childCursor)) {
        const outcome = textBoxImportHandler.read(childContext);
        if (outcome.diagnostics !== undefined) {
          diagnostics.push(...outcome.diagnostics);
        }
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        if (outcome.node !== undefined) {
          applyMemberMapper(outcome.node, memberMapper);
          recordGroupMember(
            identity.objectId,
            identity.objectName,
            outcome.node,
            childCursor.xpath,
          );
        } else {
          recordGroupOmission(identity.objectId, identity.objectName);
        }
        continue;
      }

      if (isLineNormalizedSpElement(childCursor)) {
        const outcome = readLineSp(childContext, memberMapper);
        if (outcome.diagnostics !== undefined) {
          diagnostics.push(...outcome.diagnostics);
        }
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        if (outcome.node !== undefined) {
          recordGroupMember(
            identity.objectId,
            identity.objectName,
            outcome.node,
            childCursor.xpath,
          );
        } else {
          recordGroupOmission(identity.objectId, identity.objectName);
        }
        continue;
      }

      const outcome = readShapeSp(childContext, memberMapper);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      const identity = peekShapeTreeObjectIdentity(childCursor.element);
      if (outcome.node !== undefined) {
        recordGroupMember(
          identity.objectId,
          identity.objectName,
          outcome.node,
          childCursor.xpath,
        );
      } else {
        recordGroupOmission(identity.objectId, identity.objectName);
      }
      continue;
    }

    if (childCursor.localName === "cxnSp") {
      const outcome = readLineCxnSp(childContext, memberMapper);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      const identity = peekShapeTreeObjectIdentity(childCursor.element);
      if (outcome.node !== undefined) {
        recordGroupMember(
          identity.objectId,
          identity.objectName,
          outcome.node,
          childCursor.xpath,
        );
      } else {
        recordGroupOmission(identity.objectId, identity.objectName);
      }
      continue;
    }

    if (childCursor.localName === "pic") {
      const outcome = imageImportHandler.read(childContext);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      const identity = peekShapeTreeObjectIdentity(childCursor.element);
      if (outcome.node !== undefined) {
        applyMemberMapper(outcome.node, memberMapper);
        recordGroupMember(
          identity.objectId,
          identity.objectName,
          outcome.node,
          childCursor.xpath,
        );
      } else {
        recordGroupOmission(identity.objectId, identity.objectName);
      }
      continue;
    }

    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: childCursor.xpath },
      namespaceUri: childCursor.namespaceUri,
      localName: childCursor.localName,
    });
  }

  if (members.length === 0) {
    if (diagnostics.length === 0) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: cursor.xpath },
      });
    }
    return { diagnostics };
  }

  const mappedGroupBox = parentMapper({
    x: offX,
    y: offY,
    width: extW,
    height: extH,
  });
  const unionPx = memberUnionPx(members);
  // Group resolution requires non-zero bounding width and height within writable DrawingML limits.
  if (
    (unionPx !== undefined && !isValidMappedBox(unionPx, false)) ||
    !isValidMappedBox(mappedGroupBox, true)
  ) {
    if (!diagnostics.some((d) => d.disposition === "omitted")) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: cursor.xpath },
        attributeOrElement: cursor.localName,
        rawValue:
          unionPx !== undefined && pxToEmu(unionPx.width) < 1
            ? "width=0"
            : unionPx !== undefined && pxToEmu(unionPx.height) < 1
              ? "height=0"
              : "bounds-overflow",
      });
    }
    return { diagnostics };
  }
  const matchesMemberUnion =
    unionPx !== undefined &&
    Math.abs(mappedGroupBox.x - unionPx.x) <= 1e-3 &&
    Math.abs(mappedGroupBox.y - unionPx.y) <= 1e-3 &&
    Math.abs(mappedGroupBox.width - unionPx.width) <= 1e-3 &&
    Math.abs(mappedGroupBox.height - unionPx.height) <= 1e-3;
  const isPositionedGroup =
    pxToEmu(mappedGroupBox.width) > 0 &&
    pxToEmu(mappedGroupBox.height) > 0 &&
    !matchesMemberUnion;

  // Resolve anchored connectors inside this group after member unionPx is computed so
  // removing left/top/width/height from connected Line nodes does not skew group bounds.
  if (groupDeferredConnectors.length > 0) {
    const scopeContext: ScopeConnectorResolutionContext = {
      scopeObjectsBySpid: groupObjectsBySpid,
      omittedSpids: groupOmittedSpids,
      usedAuthorIds: groupUsedAuthorIds,
      diagnostics,
    };
    for (const resolveConnector of groupDeferredConnectors) {
      resolveConnector(scopeContext);
    }
  }

  const groupStyle: GroupStyle = {
    ...(isPositionedGroup
      ? {
          position: "absolute" as const,
          left: mappedGroupBox.x,
          top: mappedGroupBox.y,
          width: mappedGroupBox.width,
          height: mappedGroupBox.height,
        }
      : {}),
    ...(rotate === undefined ? {} : { rotate }),
    ...(flipH ? { flipH: true } : {}),
    ...(flipV ? { flipV: true } : {}),
  };

  const props: GroupDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    ...(Object.keys(groupStyle).length === 0 ? {} : { style: groupStyle }),
  };

  const node = createGroupNode(props);
  node.append(...members);
  attachDeferredSlideLink(node, deferredSlideLink, context);
  return { node, diagnostics };
}

/**
 * Reads a Transitional `<p:grpSp>` recursively into an editable `GroupNode`,
 * applying affine `(chOff, chExt) -> (off, ext)` coordinate mapping and
 * converting positioned group descendants to containing-block-relative coordinates.
 */
export function readGroupSp(
  context: ShapeTreeImportContext,
  parentMapper: BoxCoordinateMapper = identityMapper,
): ShapeTreeImportOutcome {
  const outcome = readGroupSpSlideSpace(context, parentMapper);
  if (outcome.node !== undefined && parentMapper === identityMapper) {
    relativizeToContainingBlock([outcome.node], 0, 0);
  }
  return outcome;
}

type XmlElement = XmlElementCursor["element"];

function isDiagramGraphicFrame(cursor: XmlElementCursor): boolean {
  if (cursor.namespaceUri !== P_NS || cursor.localName !== "graphicFrame") {
    return false;
  }
  const graphic = directElementChild(cursor.element, A_NS, "graphic");
  if (graphic === undefined) return false;
  const graphicData = directElementChild(graphic, A_NS, "graphicData");
  return graphicData?.getAttribute("uri") === DIAGRAM_NS;
}

function findDescendantByLocalName(
  root: XmlElement,
  namespaceUri: string,
  localName: string,
): XmlElement | undefined {
  for (let child = root.firstChild; child !== null; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if (
      (el.namespaceURI ?? "") === namespaceUri &&
      (el.localName ?? el.tagName) === localName
    ) {
      return el;
    }
    const nested = findDescendantByLocalName(el, namespaceUri, localName);
    if (nested !== undefined) return nested;
  }
  return undefined;
}

function resolveDiagramDrawingPartPath(
  cursor: XmlElementCursor,
  context: ShapeTreeImportContext,
): string | undefined {
  const { package: pkg, part, relationships, relationshipEntries } = context;
  if (relationships === undefined) return undefined;

  const graphic = directElementChild(cursor.element, A_NS, "graphic");
  const graphicData =
    graphic === undefined
      ? undefined
      : directElementChild(graphic, A_NS, "graphicData");
  const relIds =
    graphicData === undefined
      ? undefined
      : directElementChild(graphicData, DIAGRAM_NS, "relIds");
  const dmRelId = relIds?.getAttributeNS(R_NS, "dm") || undefined;

  if (dmRelId !== undefined) {
    const dataRel = relationships.get(dmRelId);
    if (
      dataRel !== undefined &&
      dataRel.Type === REL_DIAGRAM_DATA &&
      dataRel.TargetMode !== "External"
    ) {
      const dataPartPath = resolvePartPath(part, dataRel.Target);
      const dataXml =
        dataPartPath === undefined ? undefined : pkg.readTextPart(dataPartPath);
      if (dataPartPath !== undefined && dataXml !== undefined) {
        try {
          const dataRoot = parseXmlPart(dataPartPath, dataXml);
          const dataModelExt = findDescendantByLocalName(
            dataRoot,
            DSP_NS,
            "dataModelExt",
          );
          const drawingRelId = dataModelExt?.getAttribute("relId") ?? undefined;
          if (drawingRelId !== undefined && drawingRelId !== "") {
            const drawingRel = relationships.get(drawingRelId);
            if (
              drawingRel !== undefined &&
              drawingRel.Type === REL_DIAGRAM_DRAWING &&
              drawingRel.TargetMode !== "External"
            ) {
              return resolvePartPath(part, drawingRel.Target);
            }
          }
          // Valid dataModel XML without a matching dataModelExt relId explicitly has no drawing cache.
          return undefined;
        } catch {
          // Malformed dataModel falls through to absence/single-rel fallback below.
        }
      }
    }
  }

  const dataRels = (relationshipEntries ?? []).filter(
    (entry) =>
      entry.Type === REL_DIAGRAM_DATA && entry.TargetMode !== "External",
  );
  const singleDrawingRels = (relationshipEntries ?? []).filter(
    (entry) =>
      entry.Type === REL_DIAGRAM_DRAWING && entry.TargetMode !== "External",
  );
  if (dataRels.length <= 1 && singleDrawingRels.length === 1) {
    return resolvePartPath(part, singleDrawingRels[0]!.Target);
  }
  return undefined;
}

function cloneDiagramNodeAsPresentation(
  source: XmlElement,
  isRootSpTree: boolean,
): XmlElement {
  const doc = source.ownerDocument!;
  const ns = source.namespaceURI ?? "";
  const local = source.localName ?? source.tagName;
  const isDsp = ns === DSP_NS;
  const targetNs = isDsp ? P_NS : ns;
  const targetLocal =
    isDsp && isRootSpTree && local === "spTree" ? "grpSp" : local;
  const targetQualified = isDsp ? `p:${targetLocal}` : source.tagName;
  const cloned =
    targetNs === ""
      ? doc.createElement(targetQualified)
      : doc.createElementNS(targetNs, targetQualified);

  for (let i = 0; i < source.attributes.length; i += 1) {
    const attr = source.attributes.item(i);
    if (attr === null) continue;
    const attrNs = attr.namespaceURI ?? "";
    if (
      attrNs === XMLNS_NS ||
      attr.name === "xmlns" ||
      attr.prefix === "xmlns"
    ) {
      continue;
    }
    const attrLocal = attr.localName ?? attr.name;
    if (isDsp && attrLocal === "modelId" && attrNs === "") {
      continue;
    }
    if (attrNs === "") {
      cloned.setAttribute(attrLocal, attr.value);
    } else {
      cloned.setAttributeNS(attrNs, attr.name, attr.value);
    }
  }

  for (
    let child = source.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType === 3 || child.nodeType === 4) {
      cloned.appendChild(doc.createTextNode(child.nodeValue ?? ""));
      continue;
    }
    if (child.nodeType !== 1) continue;
    const childEl = child as XmlElement;
    const childNs = childEl.namespaceURI ?? "";
    const childLocal = childEl.localName ?? childEl.tagName;
    if (
      childNs === DSP_NS &&
      (childLocal === "style" || childLocal === "txXfrm")
    ) {
      continue;
    }
    cloned.appendChild(cloneDiagramNodeAsPresentation(childEl, false));
  }

  return cloned;
}

function appendOrPopulateGroupXfrm(
  grpSp: XmlElement,
  frameTransform: NonNullable<ReturnType<typeof readTransformXml>>,
  objectName: string | undefined,
): boolean {
  const doc = grpSp.ownerDocument!;
  let nvGrpSpPr = directElementChild(grpSp, P_NS, "nvGrpSpPr");
  if (nvGrpSpPr === undefined) {
    nvGrpSpPr = doc.createElementNS(P_NS, "p:nvGrpSpPr");
    grpSp.insertBefore(nvGrpSpPr, grpSp.firstChild);
  }
  let cNvPr = directElementChild(nvGrpSpPr, P_NS, "cNvPr");
  if (cNvPr === undefined) {
    cNvPr = doc.createElementNS(P_NS, "p:cNvPr");
    cNvPr.setAttribute("id", "0");
    cNvPr.setAttribute("name", objectName ?? "");
    nvGrpSpPr.insertBefore(cNvPr, nvGrpSpPr.firstChild);
  } else if (
    objectName !== undefined &&
    (cNvPr.getAttribute("name") ?? "").trim() === ""
  ) {
    cNvPr.setAttribute("name", objectName);
  }
  if (directElementChild(nvGrpSpPr, P_NS, "cNvGrpSpPr") === undefined) {
    nvGrpSpPr.appendChild(doc.createElementNS(P_NS, "p:cNvGrpSpPr"));
  }
  if (directElementChild(nvGrpSpPr, P_NS, "nvPr") === undefined) {
    nvGrpSpPr.appendChild(doc.createElementNS(P_NS, "p:nvPr"));
  }

  let grpSpPr = directElementChild(grpSp, P_NS, "grpSpPr");
  if (grpSpPr === undefined) {
    grpSpPr = doc.createElementNS(P_NS, "p:grpSpPr");
    grpSp.appendChild(grpSpPr);
  }
  const existingXfrm = directElementChild(grpSpPr, A_NS, "xfrm");
  const hasFullXfrm =
    existingXfrm !== undefined &&
    directElementChild(existingXfrm, A_NS, "off") !== undefined &&
    directElementChild(existingXfrm, A_NS, "ext") !== undefined &&
    directElementChild(existingXfrm, A_NS, "chOff") !== undefined &&
    directElementChild(existingXfrm, A_NS, "chExt") !== undefined;
  if (hasFullXfrm) {
    return true;
  }
  if (existingXfrm !== undefined) {
    grpSpPr.removeChild(existingXfrm);
  }
  const xfrm = doc.createElementNS(A_NS, "a:xfrm");
  if (frameTransform.rotate !== undefined && frameTransform.rotate !== 0) {
    xfrm.setAttribute(
      "rot",
      String(Math.round(frameTransform.rotate * 60_000)),
    );
  }
  if (frameTransform.flipH) xfrm.setAttribute("flipH", "1");
  if (frameTransform.flipV) xfrm.setAttribute("flipV", "1");

  const off = doc.createElementNS(A_NS, "a:off");
  off.setAttribute("x", String(frameTransform.xEmu));
  off.setAttribute("y", String(frameTransform.yEmu));
  const ext = doc.createElementNS(A_NS, "a:ext");
  ext.setAttribute("cx", String(frameTransform.cxEmu));
  ext.setAttribute("cy", String(frameTransform.cyEmu));
  const chOff = doc.createElementNS(A_NS, "a:chOff");
  chOff.setAttribute("x", "0");
  chOff.setAttribute("y", "0");
  const chExt = doc.createElementNS(A_NS, "a:chExt");
  chExt.setAttribute("cx", String(frameTransform.cxEmu));
  chExt.setAttribute("cy", String(frameTransform.cyEmu));
  xfrm.appendChild(off);
  xfrm.appendChild(ext);
  xfrm.appendChild(chOff);
  xfrm.appendChild(chExt);
  grpSpPr.insertBefore(xfrm, grpSpPr.firstChild);
  return false;
}

/**
 * Recovers an external SmartArt `<p:graphicFrame>` (`DIAGRAM_NS`) as a native editable `Group`
 * when its pre-computed `ppt/diagrams/drawing*.xml` cache is readable and non-empty.
 */
function readDiagramGraphicFrame(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, package: pkg } = context;
  const identity = peekShapeTreeObjectIdentity(cursor.element);

  const omitUnsupported = (): ShapeTreeImportOutcome => {
    const frameInfo = inspectUnsupportedGraphicFrame(
      cursor,
      part,
      context.relationships ?? new Map(),
      context.relationshipEntries ?? [],
    );
    cursor.suppressUnconsumed();
    return {
      diagnostics: [
        {
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: {
            part,
            xpath: cursor.xpath,
            slideIndex,
            ...identity,
          },
          namespaceUri: cursor.namespaceUri,
          localName: cursor.localName,
          ...(frameInfo.mechanismHint === undefined
            ? {}
            : { mechanismHint: frameInfo.mechanismHint }),
          ...(frameInfo.cascaded.length === 0
            ? {}
            : { cascaded: frameInfo.cascaded }),
        },
      ],
    };
  };

  const drawingPartPath = resolveDiagramDrawingPartPath(cursor, context);
  const drawingXml =
    drawingPartPath === undefined
      ? undefined
      : pkg.readTextPart(drawingPartPath);
  if (drawingPartPath === undefined || drawingXml === undefined) {
    return omitUnsupported();
  }

  let drawingRootEl: XmlElement;
  try {
    drawingRootEl = parseXmlPart(drawingPartPath, drawingXml);
  } catch {
    return omitUnsupported();
  }
  if (
    (drawingRootEl.namespaceURI ?? "") !== DSP_NS ||
    (drawingRootEl.localName ?? drawingRootEl.tagName) !== "drawing"
  ) {
    return omitUnsupported();
  }
  const dspSpTree = directElementChild(drawingRootEl, DSP_NS, "spTree");
  if (dspSpTree === undefined) {
    return omitUnsupported();
  }

  const diagnostics: ImportDiagnostic[] = [];
  const nvGraphicFramePr = cursor.child(P_NS, "nvGraphicFramePr");
  let clickResult: ClickLinkReadResult | undefined;
  const cNvPr = nvGraphicFramePr?.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    cNvPr.attrsObject(["id", "name"]);
    const hlinkClick = cNvPr.child(A_NS, "hlinkClick");
    if (hlinkClick !== undefined) {
      clickResult = readClickLink(hlinkClick, context, identity);
      diagnostics.push(...clickResult.diagnostics);
    }
  }
  const xfrmCursor = cursor.child(P_NS, "xfrm");
  if (xfrmCursor === undefined) {
    return omitUnsupported();
  }
  const frameTransform = readTransformXml(
    xfrmCursor,
    { part, xpath: xfrmCursor.xpath, slideIndex, ...identity },
    diagnostics,
  );
  if (frameTransform === undefined) {
    return omitUnsupported();
  }

  const clonedGrpSp = cloneDiagramNodeAsPresentation(dspSpTree, true);
  const hadOwnXfrm = appendOrPopulateGroupXfrm(
    clonedGrpSp,
    frameTransform,
    identity.objectName,
  );
  const drawingRelIndex = buildRelationshipIndex(pkg, drawingPartPath);
  diagnostics.push(...drawingRelIndex.diagnostics);

  const grpCursor = XmlElementCursor.createRoot(
    clonedGrpSp,
    drawingPartPath,
    "/dsp:drawing/dsp:spTree",
  );
  const parentMapper: BoxCoordinateMapper = hadOwnXfrm
    ? (box) => ({
        x: frameTransform.x + box.x,
        y: frameTransform.y + box.y,
        width: box.width,
        height: box.height,
      })
    : identityMapper;

  const {
    registerImportedObject: _ignoredReg,
    claimedPlaceholderKeys: _ignoredClaimed,
    ...restContext
  } = context;
  const drawingContext: ShapeTreeImportContext = {
    ...restContext,
    cursor: grpCursor,
    part: drawingPartPath,
    relationships: drawingRelIndex.byId,
    relationshipEntries: drawingRelIndex.entries,
  };
  const outcome = readGroupSpSlideSpace(drawingContext, parentMapper);
  if (outcome.diagnostics !== undefined) {
    diagnostics.push(...outcome.diagnostics);
  }
  if (outcome.node === undefined) {
    return omitUnsupported();
  }

  relativizeToContainingBlock([outcome.node], 0, 0);
  if (clickResult?.link !== undefined) {
    outcome.node.set({ link: clickResult.link });
  }
  attachDeferredSlideLink(
    outcome.node,
    clickResult?.deferredSlideLink,
    context,
  );
  cursor.suppressUnconsumed();

  diagnostics.push({
    code: "compatibility-fallback",
    category: "unsupported",
    disposition: "fallback",
    location: {
      part,
      xpath: cursor.xpath,
      slideIndex,
      ...identity,
    },
    preferredMechanism: DIAGRAM_NS,
    selectedFallback: "group",
    loss: {
      lost: "native-mechanism",
      retainedAs: "group",
    },
  });

  return { node: outcome.node, diagnostics };
}

export const groupImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "grpSp",
  alternateLocalNames: ["graphicFrame"],
  matches: (cursor) =>
    cursor.localName === "grpSp" || isDiagramGraphicFrame(cursor),
  read: (context) =>
    context.cursor.localName === "graphicFrame"
      ? readDiagramGraphicFrame(context)
      : readGroupSp(context),
};
