import { CommonNode } from "@pptxdk/core/document";
import { pxToEmu } from "@pptxdk/emit";
import {
  A_NS,
  extentEmuFragmentSchema,
  offsetEmuFragmentSchema,
  P_NS,
  transformAttributesFragmentSchema,
  emuToPx,
  peekShapeTreeObjectIdentity,
  type ImportDiagnostic,
  type ImportLocation,
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
} from "../shared/identity.ts";
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
  const groupUsedAuthorIds = new Set<string>();
  const groupDeferredConnectors: Array<
    (ctx: ScopeConnectorResolutionContext) => void
  > = [];
  const recordGroupMember = (
    spid: number | undefined,
    memberNode: CommonNode,
  ): void => {
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
      if (outcome.node !== undefined) {
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        recordGroupMember(identity.objectId, outcome.node);
      }
      continue;
    }

    if (childCursor.localName === "sp") {
      if (textBoxImportHandler.matches?.(childCursor)) {
        const outcome = textBoxImportHandler.read(childContext);
        if (outcome.diagnostics !== undefined) {
          diagnostics.push(...outcome.diagnostics);
        }
        if (outcome.node !== undefined) {
          applyMemberMapper(outcome.node, memberMapper);
          const identity = peekShapeTreeObjectIdentity(childCursor.element);
          recordGroupMember(identity.objectId, outcome.node);
        }
        continue;
      }

      if (isLineNormalizedSpElement(childCursor)) {
        const outcome = readLineSp(childContext, memberMapper);
        if (outcome.diagnostics !== undefined) {
          diagnostics.push(...outcome.diagnostics);
        }
        if (outcome.node !== undefined) {
          const identity = peekShapeTreeObjectIdentity(childCursor.element);
          recordGroupMember(identity.objectId, outcome.node);
        }
        continue;
      }

      const outcome = readShapeSp(childContext, memberMapper);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      if (outcome.node !== undefined) {
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        recordGroupMember(identity.objectId, outcome.node);
      }
      continue;
    }

    if (childCursor.localName === "cxnSp") {
      const outcome = readLineCxnSp(childContext, memberMapper);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      if (outcome.node !== undefined) {
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        recordGroupMember(identity.objectId, outcome.node);
      }
      continue;
    }

    if (childCursor.localName === "pic") {
      const outcome = imageImportHandler.read(childContext);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      if (outcome.node !== undefined) {
        applyMemberMapper(outcome.node, memberMapper);
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        recordGroupMember(identity.objectId, outcome.node);
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
      omittedSpids: new Map(),
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

export const groupImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "grpSp",
  read: (context) => readGroupSp(context),
};
