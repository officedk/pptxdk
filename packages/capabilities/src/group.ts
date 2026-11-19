import {
  CapabilityError,
  layoutStyleKeys,
  paddingStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  textStyleFrom,
  textStyleKeys,
  xmlEscape,
  type BoundContainerCollaboration,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type LayoutStyle,
  type LinkTarget,
  type ObjectCapabilityModule,
  type ObjectObservation,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type PreviewTransform,
  type ResolvedBoundObjectCapability,
  type ResolvedObject,
  type TextStyle,
  type ThemeScope,
  type VisualStyle,
} from "@react-pptx/core/capability";
import { CommonContainer } from "@react-pptx/core/document";

import { groupXml } from "./group-xml.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "./identity.ts";
import type { MorphProps } from "./morph-api.ts";
import {
  motionContainerAnnotations,
  validateMotionFlag,
  type MotionProps,
} from "./motion-api.ts";

/** Native Group preserves member relationships; style allows it to arrange as a layout child, with output bounds computed from the union of member geometries. */
export type GroupDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly style?: LayoutStyle & TextStyle & Pick<VisualStyle, "rotate">;
    readonly theme?: ThemeScope;
  };

/** Creates the editable native group shared by the command factory and the Group JSX element. */
export function createGroupNode(
  props: GroupDeclarationProps,
): CommonContainer<
  "group",
  GroupDeclarationProps,
  NonNullable<GroupDeclarationProps["style"]>
> {
  return new CommonContainer<
    "group",
    GroupDeclarationProps,
    NonNullable<GroupDeclarationProps["style"]>
  >("group", "rp:group", props);
}

const allowedParameters = [
  "id",
  "name",
  "morphKey",
  "link",
  "style",
  "theme",
  "enter",
  "emphasis",
  "motion",
] as const;
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

const groupAllowedStyleKeys = new Set<string>([
  ...layoutStyleKeys,
  ...paddingStyleKeys,
  ...textStyleKeys,
  "rotate",
]);

/** Preserves members as native objects and writes group bounds according to final member geometry; layout participation uses generic container collaboration. */
export const groupCapability: ObjectCapabilityModule = {
  declaration: "rp:group",

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): ResolvedBoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Group contains unsupported parameter ${unknown}`,
        context.path,
      );
    }
    validateMotionFlag(input.props["motion"]);
    const requestedName = parseOptionalName(input.props["name"], context.path);
    const morphKey = parseMorphKey(input.props["morphKey"], context.path);
    const link = parseLinkTarget(input.props["link"], context.path);

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
      (typeof rawRotation !== "number" || !Number.isFinite(rawRotation))
    ) {
      throw new CapabilityError(
        "invalid-visual-style",
        "rotate must be a finite angle",
        context.path,
      );
    }
    const rotation = rawRotation;

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
      layoutMode: "when-styled",
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
      resolve(resolveContext: ObjectResolveContext): ResolvedObject {
        const resolvedChildren = resolveContext.resolveChildren(input.children);
        const geometry = bounds(resolvedChildren.objects);
        const outputGeometry = resolveContext.layoutBox ?? geometry.pixels;
        if (outputGeometry.width <= 0 || outputGeometry.height <= 0) {
          throw new CapabilityError(
            "invalid-group-geometry",
            "Group layout width and height must be greater than 0",
            resolveContext.path,
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
            resolveContext.path,
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
            return {
              xml: groupXml(
                emitContext.id,
                escapedName,
                outputEmu,
                emitContext.childrenXml,
                rotation,
                nativeLinkXml(emitContext.nativeLink),
              ),
            };
          },
          paint(paintContext: PaintContext): PaintOutput {
            const childrenSvg = paintContext.childrenSvg ?? "";
            const cx = outputGeometry.x + outputGeometry.width / 2;
            const cy = outputGeometry.y + outputGeometry.height / 2;
            return {
              svg: `<g transform="rotate(${angle} ${cx} ${cy})">${childrenSvg}</g>`,
            };
          },
        };
      },
    };
  },
};
