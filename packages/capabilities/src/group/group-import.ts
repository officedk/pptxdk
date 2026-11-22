import { CommonNode } from "@pptxdk/core/document";
import { pxToEmu } from "@pptxdk/emit";
import {
  A_NS,
  ExtentEmuFragmentSchema,
  OffsetEmuFragmentSchema,
  P_NS,
  TransformAttributesFragmentSchema,
  emuToPx,
  peekShapeTreeObjectIdentity,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type ShapeTreeImportHandler,
  type ShapeTreeImportOutcome,
} from "@pptxdk/import";

import { imageImportHandler } from "../image/image-import.ts";
import {
  readNonVisualCommonProperties,
  readShapeSp,
  type BoxCoordinateMapper,
} from "../shape/shape-import.ts";
import { consumeNonPlaceholderNvPr } from "../shared/identity.ts";
import { textBoxImportHandler } from "../text/text-import.ts";
import {
  createGroupNode,
  type GroupDeclarationProps,
} from "./group-declaration.ts";

function nodeBoundsEmu(node: CommonNode):
  | {
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
    }
  | undefined {
  const props = node.props as {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly bounds?: {
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
    };
  };
  if (props.bounds !== undefined) {
    return {
      x: pxToEmu(props.bounds.x),
      y: pxToEmu(props.bounds.y),
      width: pxToEmu(props.bounds.width),
      height: pxToEmu(props.bounds.height),
    };
  }
  if (
    typeof props.x === "number" &&
    typeof props.y === "number" &&
    typeof props.width === "number" &&
    typeof props.height === "number"
  ) {
    return {
      x: pxToEmu(props.x),
      y: pxToEmu(props.y),
      width: pxToEmu(props.width),
      height: pxToEmu(props.height),
    };
  }
  if (node.type === "group") {
    const childNodes = node.children.filter(
      (child): child is CommonNode => child instanceof CommonNode,
    );
    return memberUnionEmu(childNodes);
  }
  return undefined;
}

function memberUnionEmu(members: readonly CommonNode[]):
  | {
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
    }
  | undefined {
  if (members.length === 0) return undefined;
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxRight = Number.NEGATIVE_INFINITY;
  let maxBottom = Number.NEGATIVE_INFINITY;
  for (const member of members) {
    const box = nodeBoundsEmu(member);
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
  const props = node.props as {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
  };
  if (
    typeof props.x === "number" &&
    typeof props.y === "number" &&
    typeof props.width === "number" &&
    typeof props.height === "number"
  ) {
    node.set(
      memberMapper({
        x: props.x,
        y: props.y,
        width: props.width,
        height: props.height,
      }),
    );
  }
}

/**
 * Reads a Transitional `<p:grpSp>` recursively into an editable `GroupNode`,
 * applying affine `(chOff, chExt) -> (off, ext)` coordinate mapping for external groups.
 */
export function readGroupSp(
  context: ShapeTreeImportContext,
  parentMapper: BoxCoordinateMapper = identityMapper,
): ShapeTreeImportOutcome {
  const { cursor, part, slideIndex, package: pkg } = context;
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
  const { objectId, objectName, link } = readNonVisualCommonProperties(
    cNvPr,
    pkg,
    part,
    slideIndex,
    diagnostics,
  );
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

  const parsedXfrm = TransformAttributesFragmentSchema.safeParse(
    xfrm.attrsObject(["rot", "flipH", "flipV"]),
  );
  const parsedOff = OffsetEmuFragmentSchema.safeParse(
    off.attrsObject(["x", "y"]),
  );
  const parsedExt = ExtentEmuFragmentSchema.safeParse(
    ext.attrsObject(["cx", "cy"]),
  );
  const parsedChOff = OffsetEmuFragmentSchema.safeParse(
    chOff.attrsObject(["x", "y"]),
  );
  const parsedChExt = ExtentEmuFragmentSchema.safeParse(
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

  if (parsedXfrm.success) {
    if (parsedXfrm.data.flipH === true) {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: `${xfrm.xpath}/@flipH` },
        namespaceUri: A_NS,
        localName: "flipH",
        mechanismHint: "group-flipH",
      });
    }
    if (parsedXfrm.data.flipV === true) {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: `${xfrm.xpath}/@flipV` },
        namespaceUri: A_NS,
        localName: "flipV",
        mechanismHint: "group-flipV",
      });
    }
  }

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
    };

    if (childCursor.localName === "grpSp") {
      const outcome = readGroupSp(childContext, memberMapper);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      if (outcome.node !== undefined) {
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        if (identity.objectId !== undefined) {
          context.registerImportedObject?.(identity.objectId, outcome.node);
        }
        members.push(outcome.node);
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
          if (identity.objectId !== undefined) {
            context.registerImportedObject?.(identity.objectId, outcome.node);
          }
          members.push(outcome.node);
        }
        continue;
      }

      const outcome = readShapeSp(childContext, memberMapper);
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      if (outcome.node !== undefined) {
        const identity = peekShapeTreeObjectIdentity(childCursor.element);
        if (identity.objectId !== undefined) {
          context.registerImportedObject?.(identity.objectId, outcome.node);
        }
        members.push(outcome.node);
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
        if (identity.objectId !== undefined) {
          context.registerImportedObject?.(identity.objectId, outcome.node);
        }
        members.push(outcome.node);
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
  const unionEmu = memberUnionEmu(members);
  const preserveExplicitGroupBox =
    pxToEmu(mappedGroupBox.width) > 0 &&
    pxToEmu(mappedGroupBox.height) > 0 &&
    (unionEmu === undefined ||
      unionEmu.x !== pxToEmu(mappedGroupBox.x) ||
      unionEmu.y !== pxToEmu(mappedGroupBox.y) ||
      unionEmu.width !== pxToEmu(mappedGroupBox.width) ||
      unionEmu.height !== pxToEmu(mappedGroupBox.height));

  const props: GroupDeclarationProps = {
    ...(objectName === undefined ? {} : { name: objectName }),
    ...(link === undefined ? {} : { link }),
    ...(preserveExplicitGroupBox
      ? {
          bounds: {
            x: mappedGroupBox.x,
            y: mappedGroupBox.y,
            width: mappedGroupBox.width,
            height: mappedGroupBox.height,
          },
        }
      : {}),
    ...(rotate === undefined ? {} : { style: { rotate } }),
  };

  const node = createGroupNode(props);
  node.append(...members);
  return { node, diagnostics };
}

export const groupImportHandler: ShapeTreeImportHandler = {
  target: "shape-tree",
  localName: "grpSp",
  read: (context) => readGroupSp(context),
};
