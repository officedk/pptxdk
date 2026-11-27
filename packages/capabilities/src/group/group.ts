import {
  CapabilityError,
  parseCapabilityLayoutContract,
  textStyleFrom,
  type BoundContainerCollaboration,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type ObjectCapabilityModule,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ObjectResolution,
  type Rect,
  type ResolvedObject,
  type SealedCanvasMechanism,
} from "@pptxdk/core/capability";
import { pxToEmu, xmlEscape } from "@pptxdk/emit";
import { defineCanvasMechanism } from "@pptxdk/import";

import {
  motionContainerAnnotations,
  validateMotionFlag,
} from "../motion/motion-api.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "../shared/identity.ts";
import { groupCodegen } from "./group-codegen.ts";
import {
  groupDeclarationPropsKeys,
  groupStyleKeys,
  groupTransformShape,
} from "./group-declaration.ts";
import { groupImportHandler } from "./group-import.ts";
import { groupXml } from "./group-xml.ts";

export {
  createGroupNode,
  groupDeclarationPropsSchema,
  groupStyleSchema,
  type GroupDeclarationProps,
  type GroupStyle,
} from "./group-declaration.ts";

const allowedParameters: readonly string[] = groupDeclarationPropsKeys;
const emuPerPixel = 9525;
const maximumGroupExtentEmu = 2_147_483_647;

type EmuBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

type MotionNativeMember = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly authorId?: string;
  readonly box?: NonNullable<ResolvedObject["box"]>;
  readonly members?: readonly MotionNativeMember[];
};

function motionMembers(
  objects: readonly ResolvedObject[],
): readonly MotionNativeMember[] {
  return objects.map((object) => ({
    id: object.id,
    kind: object.kind,
    name: object.name,
    ...(object.authorId === undefined ? {} : { authorId: object.authorId }),
    ...(object.box === undefined ? {} : { box: object.box }),
    ...(object.members === undefined
      ? {}
      : { members: motionMembers(object.members) }),
  }));
}

function bounds(members: readonly ResolvedObject[]): {
  readonly pixels: { x: number; y: number; width: number; height: number };
  readonly emu: EmuBox;
} {
  if (members.length === 0) {
    throw new CapabilityError(
      "empty-group",
      "Group requires at least one native object member",
    );
  }
  const memberBoxes = members.map((member) => {
    const box = member.box;
    if (
      box === undefined ||
      !Number.isFinite(box.x) ||
      !Number.isFinite(box.y) ||
      !Number.isFinite(box.width) ||
      !Number.isFinite(box.height) ||
      box.width <= 0 ||
      box.height <= 0
    ) {
      throw new CapabilityError(
        "member-without-geometry",
        `Group member ${member.name} is missing composable geometry bounds`,
      );
    }
    const x = pxToEmu(box.x);
    const y = pxToEmu(box.y);
    return {
      x,
      y,
      right: x + pxToEmu(box.width),
      bottom: y + pxToEmu(box.height),
    };
  });
  const x = Math.min(...memberBoxes.map((box) => box.x));
  const y = Math.min(...memberBoxes.map((box) => box.y));
  const right = Math.max(...memberBoxes.map((box) => box.right));
  const bottom = Math.max(...memberBoxes.map((box) => box.bottom));
  const width = right - x;
  const height = bottom - y;
  if (width < 1) {
    throw new CapabilityError(
      "invalid-group-geometry",
      "Computed Group width is 0px; width must be greater than 0px",
    );
  }
  if (height < 1) {
    throw new CapabilityError(
      "invalid-group-geometry",
      "Computed Group height is 0px; height must be greater than 0px",
    );
  }
  if (width > maximumGroupExtentEmu || height > maximumGroupExtentEmu) {
    throw new CapabilityError(
      "invalid-group-geometry",
      "Group member bounds exceed writable PresentationML dimensions",
    );
  }
  const emu = { x, y, width, height };
  return {
    emu,
    pixels: {
      x: x / emuPerPixel,
      y: y / emuPerPixel,
      width: width / emuPerPixel,
      height: height / emuPerPixel,
    },
  };
}

const groupAllowedStyleKeys = new Set<string>(groupStyleKeys);

/** Closed resolved group model produced in resolve() and consumed by emitGroupModel and paintGroupModel. */
type ResolvedGroupModel = {
  readonly escapedName: string;
  readonly box: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  readonly emu: EmuBox;
  readonly angle: number;
  readonly rotation?: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
};

function roundSvgNumber(value: number): string {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function emitGroupModel(
  model: ResolvedGroupModel,
  emitContext: EmitContext,
): EmitOutput {
  return {
    xml: groupXml(
      emitContext.id,
      model.escapedName,
      model.emu,
      emitContext.childrenXml,
      model.rotation,
      nativeLinkXml(emitContext.nativeLink),
      model.flipH,
      model.flipV,
    ),
  };
}

function paintGroupModel(
  model: ResolvedGroupModel,
  paintContext: PaintContext,
): PaintOutput {
  const childrenSvg = paintContext.childrenSvg ?? "";
  const cxNum = model.box.x + model.box.width / 2;
  const cyNum = model.box.y + model.box.height / 2;
  const cx = roundSvgNumber(cxNum);
  const cy = roundSvgNumber(cyNum);
  const transform =
    model.flipH || model.flipV
      ? `translate(${cx} ${cy}) rotate(${roundSvgNumber(model.angle)}) scale(${model.flipH ? -1 : 1} ${model.flipV ? -1 : 1}) translate(${roundSvgNumber(-cxNum)} ${roundSvgNumber(-cyNum)})`
      : `rotate(${roundSvgNumber(model.angle)} ${cx} ${cy})`;
  return {
    svg: `<g transform="${transform}">${childrenSvg}</g>`,
  };
}

/** Native canvas mechanism for PresentationML `<p:grpSp>` native groups. */
export const groupMechanism = defineCanvasMechanism({
  key: "group",
  role: "canvas",
  emit: emitGroupModel,
  paint: paintGroupModel,
  import: groupImportHandler,
});

/** Preserves members as native objects and writes group bounds according to final member geometry; layout participation uses generic container collaboration. */
export const groupCapability: ObjectCapabilityModule<
  BoundObjectCapability,
  readonly [SealedCanvasMechanism<"group">]
> = {
  declaration: "group",
  codegen: groupCodegen,
  mechanisms: [groupMechanism] as const,

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !allowedParameters.includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Group contains unsupported parameter ${unknown}`,
      );
    }
    validateMotionFlag(input.props["motion"]);
    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "grpSp",
      props: input.props,
      allowedStyleKeys: groupAllowedStyleKeys,
      allowPadding: true,
      isContainer: true,
      path: context.path,
    });

    const rawStyle = input.props["style"];
    const styleRecord =
      typeof rawStyle === "object" &&
      rawStyle !== null &&
      !Array.isArray(rawStyle)
        ? (rawStyle as Record<string, unknown>)
        : {};

    const rawRotation = styleRecord["rotate"];
    if (
      rawRotation !== undefined &&
      !groupTransformShape.rotate.safeParse(rawRotation).success
    ) {
      throw new CapabilityError(
        "invalid-visual-style",
        "rotate must be a finite angle",
      );
    }
    const rotation = rawRotation as number | undefined;
    const rawFlipH = styleRecord["flipH"];
    if (
      rawFlipH !== undefined &&
      !groupTransformShape.flipH.safeParse(rawFlipH).success
    ) {
      throw new CapabilityError(
        "invalid-visual-style",
        "flipH must be a boolean",
      );
    }
    const rawFlipV = styleRecord["flipV"];
    if (
      rawFlipV !== undefined &&
      !groupTransformShape.flipV.safeParse(rawFlipV).success
    ) {
      throw new CapabilityError(
        "invalid-visual-style",
        "flipV must be a boolean",
      );
    }
    const flipH = Boolean(rawFlipH);
    const flipV = Boolean(rawFlipV);

    const theme = context.theme.with(textStyleFrom(styleRecord));

    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = {
      ...(input.props["enter"] !== undefined
        ? { enter: input.props["enter"] }
        : {}),
      ...(input.props["emphasis"] !== undefined
        ? { emphasis: input.props["emphasis"] }
        : {}),
      ...(input.props["motion"] !== undefined
        ? { motion: input.props["motion"] }
        : {}),
    };

    const container: BoundContainerCollaboration = {
      children: input.children,
      theme,
      emitsObject: true,
      childComposition: "members",
      restrictsMemberContinuity: true,
      observesLayout: false,
      propagatesBaseline: false,
    };

    const native: BoundObjectNativeCollaboration = {
      ...(morphKey === undefined ? {} : { morphKey }),
      ...(requestedName === undefined ? {} : { hasExplicitName: true }),
      ...(link === undefined ? {} : { link }),
    };

    return {
      contract,
      container,
      ...(morphKey !== undefined ||
      requestedName !== undefined ||
      link !== undefined
        ? { native }
        : {}),
      resolve(resolveContext: ObjectResolveContext): ObjectResolution {
        const resolvedChildren = resolveContext.resolveChildren(input.children);
        const geometry = bounds(resolvedChildren.objects);
        const layoutBox = resolveContext.layoutBox;
        const isPositioned = contract.positioned === true;
        const fixedAxes = contract.fixedAxes ?? {
          width: false,
          height: false,
        };
        const sizedAxes = resolveContext.layoutSizedAxes ?? fixedAxes;
        const hasLayoutWidth = sizedAxes.width;
        const hasLayoutHeight = sizedAxes.height;

        let outputGeometry: Rect;
        if (isPositioned) {
          const memberBounds = geometry.pixels;
          if (layoutBox !== undefined) {
            const resolvedWidth =
              fixedAxes.width || (hasLayoutWidth && layoutBox.width > 0)
                ? layoutBox.width
                : memberBounds.width > 0
                  ? Math.max(
                      layoutBox.width,
                      memberBounds.width,
                      memberBounds.x + memberBounds.width - layoutBox.x,
                    )
                  : layoutBox.width;
            const resolvedHeight =
              fixedAxes.height || (hasLayoutHeight && layoutBox.height > 0)
                ? layoutBox.height
                : memberBounds.height > 0
                  ? Math.max(
                      layoutBox.height,
                      memberBounds.height,
                      memberBounds.y + memberBounds.height - layoutBox.y,
                    )
                  : layoutBox.height;
            outputGeometry = {
              x: layoutBox.x,
              y: layoutBox.y,
              width: resolvedWidth,
              height: resolvedHeight,
            };
          } else {
            outputGeometry = memberBounds;
          }
        } else {
          const memberBounds = geometry.pixels;
          const useLayoutWidth =
            fixedAxes.width ||
            (hasLayoutWidth && layoutBox !== undefined && layoutBox.width > 0);
          const useLayoutHeight =
            fixedAxes.height ||
            (hasLayoutHeight &&
              layoutBox !== undefined &&
              layoutBox.height > 0);
          if (layoutBox !== undefined && useLayoutWidth && useLayoutHeight) {
            outputGeometry = layoutBox;
          } else if (
            layoutBox !== undefined &&
            layoutBox.width > 0 &&
            layoutBox.height > 0
          ) {
            const x = useLayoutWidth
              ? layoutBox.x
              : memberBounds.width > 0
                ? Math.min(layoutBox.x, memberBounds.x)
                : layoutBox.x;
            const y = useLayoutHeight
              ? layoutBox.y
              : memberBounds.height > 0
                ? Math.min(layoutBox.y, memberBounds.y)
                : layoutBox.y;
            const right = useLayoutWidth
              ? layoutBox.x + layoutBox.width
              : memberBounds.width > 0
                ? Math.max(
                    layoutBox.x + layoutBox.width,
                    memberBounds.x + memberBounds.width,
                  )
                : layoutBox.x + layoutBox.width;
            const bottom = useLayoutHeight
              ? layoutBox.y + layoutBox.height
              : memberBounds.height > 0
                ? Math.max(
                    layoutBox.y + layoutBox.height,
                    memberBounds.y + memberBounds.height,
                  )
                : layoutBox.y + layoutBox.height;
            outputGeometry = {
              x,
              y,
              width: right - x,
              height: bottom - y,
            };
          } else if (
            layoutBox !== undefined &&
            (useLayoutWidth || useLayoutHeight)
          ) {
            outputGeometry = {
              x: useLayoutWidth ? layoutBox.x : memberBounds.x,
              y: useLayoutHeight ? layoutBox.y : memberBounds.y,
              width: useLayoutWidth ? layoutBox.width : memberBounds.width,
              height: useLayoutHeight ? layoutBox.height : memberBounds.height,
            };
          } else {
            outputGeometry = memberBounds;
          }
        }
        if (outputGeometry.width === 0 || pxToEmu(outputGeometry.width) === 0) {
          throw new CapabilityError(
            "invalid-group-geometry",
            "Computed Group width is 0px; width must be greater than 0px",
          );
        }
        if (outputGeometry.width < 0 || pxToEmu(outputGeometry.width) < 1) {
          throw new CapabilityError(
            "invalid-group-geometry",
            `Computed Group width is ${outputGeometry.width}px; width must be greater than 0px`,
          );
        }
        if (
          outputGeometry.height === 0 ||
          pxToEmu(outputGeometry.height) === 0
        ) {
          throw new CapabilityError(
            "invalid-group-geometry",
            "Computed Group height is 0px; height must be greater than 0px",
          );
        }
        if (outputGeometry.height < 0 || pxToEmu(outputGeometry.height) < 1) {
          throw new CapabilityError(
            "invalid-group-geometry",
            `Computed Group height is ${outputGeometry.height}px; height must be greater than 0px`,
          );
        }
        const outputEmu = {
          x: pxToEmu(outputGeometry.x),
          y: pxToEmu(outputGeometry.y),
          width: pxToEmu(outputGeometry.width),
          height: pxToEmu(outputGeometry.height),
        };
        if (
          outputEmu.width > maximumGroupExtentEmu ||
          outputEmu.height > maximumGroupExtentEmu
        ) {
          throw new CapabilityError(
            "invalid-group-geometry",
            "Group bounds exceed writable PresentationML dimensions",
          );
        }
        const name =
          resolveContext.morphName ??
          requestedName ??
          `Group ${resolveContext.id - 1}`;
        let escapedName: string;
        try {
          escapedName = xmlEscape(name);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Group.name contains characters not allowed in XML 1.0",
          );
        }
        const members = motionMembers(resolvedChildren.objects);
        const annotations = motionContainerAnnotations(motionProps, {
          kind: "group",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: false,
          filled: false,
          box: outputGeometry,
          rotationDegrees: rotation ?? 0,
          members,
        });
        const angle = rotation ?? 0;
        const resolvedModel: ResolvedGroupModel = {
          escapedName,
          box: outputGeometry,
          emu: outputEmu,
          angle,
          ...(rotation === undefined ? {} : { rotation }),
          flipH,
          flipV,
        };
        const transform: PreviewTransform = {
          box: outputGeometry,
          angle,
          flipH,
          flipV,
        };
        return {
          observation: {
            id: resolveContext.id,
            kind: "group",
            name,
            box: outputGeometry,
            members: resolvedChildren.objects,
          },
          transform,
          extensions: [],
          ...(annotations.length === 0 ? {} : { annotations }),
          emit(emitContext: EmitContext): EmitOutput {
            return groupMechanism.emit(resolvedModel, emitContext);
          },
          paint(paintContext: PaintContext): PaintOutput {
            return groupMechanism.paint(resolvedModel, paintContext);
          },
        };
      },
    };
  },
} satisfies ObjectCapabilityModule;
