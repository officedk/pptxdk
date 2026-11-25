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
  type GeometryBox,
  type ObjectCapabilityModule,
  type ObjectObservation,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ObjectResolution,
  type ThemeScope,
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
  groupDeclarationPropsSchema,
  groupStyleSchema,
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
  readonly box?: NonNullable<ObjectObservation["box"]>;
  readonly members?: readonly MotionNativeMember[];
};

function motionMembers(
  objects: readonly ObjectObservation[],
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

function bounds(members: readonly ObjectObservation[]): {
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
  if (
    width < 1 ||
    height < 1 ||
    width > maximumGroupExtentEmu ||
    height > maximumGroupExtentEmu
  ) {
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

const groupAllowedStyleKeys = new Set<string>(
  Object.keys(groupStyleSchema.unwrap().shape),
);

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
    ),
  };
}

function paintGroupModel(
  model: ResolvedGroupModel,
  paintContext: PaintContext,
): PaintOutput {
  const childrenSvg = paintContext.childrenSvg ?? "";
  const cx = roundSvgNumber(model.box.x + model.box.width / 2);
  const cy = roundSvgNumber(model.box.y + model.box.height / 2);
  return {
    svg: `<g transform="rotate(${roundSvgNumber(model.angle)} ${cx} ${cy})">${childrenSvg}</g>`,
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
export const groupCapability = {
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
      !groupStyleSchema.unwrap().shape.rotate.safeParse(rawRotation).success
    ) {
      throw new CapabilityError(
        "invalid-visual-style",
        "rotate must be a finite angle",
      );
    }
    const rotation = rawRotation as number | undefined;

    let explicitBox:
      | {
          readonly x: number;
          readonly y: number;
          readonly width: number;
          readonly height: number;
        }
      | undefined;
    if (input.props["bounds"] !== undefined) {
      const parsedBounds = groupDeclarationPropsSchema
        .unwrap()
        .shape.bounds.safeParse(input.props["bounds"]);
      if (!parsedBounds.success) {
        throw new CapabilityError(
          "invalid-group-geometry",
          parsedBounds.error.issues[0]?.message ?? "Invalid Group bounds",
        );
      }
      explicitBox = parsedBounds.data;
    }

    const theme = context.theme.with(
      input.props["theme"] as ThemeScope | undefined,
      textStyleFrom(styleRecord),
    );

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
      checksCanvasOverflow: false,
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
        const position = contract.layoutStyle.position;
        const isPositioned = position === "relative" || position === "absolute";
        const hasExplicitWidth =
          contract.layoutStyle.width !== undefined ||
          ((contract.layoutStyle.left !== undefined ||
            contract.layoutStyle.inset !== undefined) &&
            (contract.layoutStyle.right !== undefined ||
              contract.layoutStyle.inset !== undefined));
        const hasExplicitHeight =
          contract.layoutStyle.height !== undefined ||
          ((contract.layoutStyle.top !== undefined ||
            contract.layoutStyle.inset !== undefined) &&
            (contract.layoutStyle.bottom !== undefined ||
              contract.layoutStyle.inset !== undefined));
        const hasLayoutWidth =
          hasExplicitWidth ||
          contract.layoutStyle.flex !== undefined ||
          contract.layoutStyle.flexGrow !== undefined ||
          contract.layoutStyle.flexBasis !== undefined;
        const hasLayoutHeight = hasExplicitHeight;

        let outputGeometry: GeometryBox;
        if (
          !isPositioned &&
          explicitBox !== undefined &&
          !hasExplicitWidth &&
          !hasExplicitHeight
        ) {
          // Transitional bridge (Tickets 03-10): imported groups with Group.bounds keep
          // slide-space members un-offset until Ticket 06 migrates group-import.
          outputGeometry = explicitBox;
        } else if (isPositioned) {
          const memberBounds = geometry.pixels;
          if (layoutBox !== undefined) {
            const resolvedWidth =
              hasExplicitWidth || layoutBox.width > 0
                ? layoutBox.width
                : memberBounds.width > 0
                  ? Math.max(
                      memberBounds.width,
                      memberBounds.x + memberBounds.width - layoutBox.x,
                    )
                  : 0;
            const resolvedHeight =
              hasExplicitHeight || layoutBox.height > 0
                ? layoutBox.height
                : memberBounds.height > 0
                  ? Math.max(
                      memberBounds.height,
                      memberBounds.y + memberBounds.height - layoutBox.y,
                    )
                  : 0;
            outputGeometry = {
              x: layoutBox.x,
              y: layoutBox.y,
              width: resolvedWidth,
              height: resolvedHeight,
            };
          } else {
            outputGeometry = explicitBox ?? memberBounds;
          }
        } else {
          const memberBounds = geometry.pixels;
          if (
            layoutBox !== undefined &&
            hasLayoutWidth &&
            hasLayoutHeight &&
            layoutBox.width > 0 &&
            layoutBox.height > 0
          ) {
            outputGeometry = layoutBox;
          } else if (
            layoutBox !== undefined &&
            layoutBox.width > 0 &&
            layoutBox.height > 0
          ) {
            const x = hasLayoutWidth
              ? layoutBox.x
              : memberBounds.width > 0
                ? Math.min(layoutBox.x, memberBounds.x)
                : layoutBox.x;
            const y = hasLayoutHeight
              ? layoutBox.y
              : memberBounds.height > 0
                ? Math.min(layoutBox.y, memberBounds.y)
                : layoutBox.y;
            const right = hasLayoutWidth
              ? layoutBox.x + layoutBox.width
              : memberBounds.width > 0
                ? Math.max(
                    layoutBox.x + layoutBox.width,
                    memberBounds.x + memberBounds.width,
                  )
                : layoutBox.x + layoutBox.width;
            const bottom = hasLayoutHeight
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
            (hasLayoutWidth || hasLayoutHeight)
          ) {
            outputGeometry = {
              x: hasLayoutWidth ? layoutBox.x : memberBounds.x,
              y: hasLayoutHeight ? layoutBox.y : memberBounds.y,
              width:
                hasLayoutWidth && layoutBox.width > 0
                  ? layoutBox.width
                  : memberBounds.width,
              height:
                hasLayoutHeight && layoutBox.height > 0
                  ? layoutBox.height
                  : memberBounds.height,
            };
          } else {
            outputGeometry = memberBounds;
          }
        }
        if (outputGeometry.width <= 0 || outputGeometry.height <= 0) {
          throw new CapabilityError(
            "invalid-group-geometry",
            "Group layout width and height must be greater than 0",
          );
        }
        const outputEmu = {
          x: pxToEmu(outputGeometry.x),
          y: pxToEmu(outputGeometry.y),
          width: pxToEmu(outputGeometry.width),
          height: pxToEmu(outputGeometry.height),
        };
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
        };
        const transform: PreviewTransform = {
          box: outputGeometry,
          angle,
          flipH: false,
          flipV: false,
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
